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
function CCollaborativeEditing()
{
	AscCommon.CCollaborativeEditingBase.call(this);
    this.PosExtChangesX = [];
    this.PosExtChangesY = [];
    this.ScaleX = null;
    this.ScaleY = null;
}

CCollaborativeEditing.prototype = Object.create(AscCommon.CCollaborativeEditingBase.prototype);
CCollaborativeEditing.prototype.constructor = CCollaborativeEditing;

CCollaborativeEditing.prototype.GetEditorApi = function()
{
    return Asc.editor;
};
CCollaborativeEditing.prototype.GetDrawingDocument = function()
{
    return Asc.editor.getDrawingDocument();
};
CCollaborativeEditing.prototype.GetPresentation = function()
{
    return Asc.editor.private_GetLogicDocument();
};
CCollaborativeEditing.prototype.Send_Changes = function(IsUserSave, AdditionalInfo, IsUpdateInterface, isAfterAskSave)
{
    // Recalculate positions
    this.Refresh_DCChanges();
    this.RefreshPosExtChanges();
    // Generate our changes
    let StartPoint = ( null === AscCommon.History.SavedIndex ? 0 : AscCommon.History.SavedIndex + 1 );
    let LastPoint  = -1;
    if ( this.m_nUseType <= 0 )
    {
        // (remove unnecessary points beforehand)
        AscCommon.History.Clear_Redo();
        LastPoint = AscCommon.History.Points.length - 1;
    }
    else
    {
        LastPoint = AscCommon.History.Index;
    }
    // Calculate how many changes do not need to be sent to the server
	let SumIndex = 0;
	let StartPoint2 = Math.min( StartPoint, LastPoint + 1 );
    for ( let PointIndex = 0; PointIndex < StartPoint2; PointIndex++ )
    {
		let Point = AscCommon.History.Points[PointIndex];
        SumIndex += Point.Items.length;
    }
	let deleteIndex = ( null === AscCommon.History.SavedIndex ? null : SumIndex );

	let aChanges = [], aChanges2 = [];
    for ( let PointIndex = StartPoint; PointIndex <= LastPoint; PointIndex++ )
    {
		let Point = AscCommon.History.Points[PointIndex];

        AscCommon.History.Update_PointInfoItem(PointIndex, StartPoint, LastPoint, SumIndex, deleteIndex);
        for ( let Index = 0; Index < Point.Items.length; Index++ )
        {
			let Item = Point.Items[Index];
			let oChanges = new AscCommon.CCollaborativeChanges();
            oChanges.Set_FromUndoRedo( Item.Class, Item.Data, Item.Binary );
            aChanges2.push(Item.Data);
            aChanges.push( oChanges.m_pData );
        }
    }


    // While the user is alone, we don't clear their locks until another user joins
	let bCollaborative = this.getCollaborativeEditing();

	let num_arr = [];
	let drawingDocument = this.GetDrawingDocument();
	let presentation = this.GetPresentation();
    if (bCollaborative)
	{
		let map = this.Release_Locks();

		let UnlockCount2 = this.m_aNeedUnlock2.length;
		for (let Index = 0; Index < UnlockCount2; Index++ )
		{
			let Class = this.m_aNeedUnlock2[Index];
			Class.Lock.Set_Type( AscCommon.c_oAscLockTypes.kLockTypeNone, false);
            if(Class instanceof AscCommon.CCore)
            {
				Asc.editor.sendEvent("asc_onLockCore", false);
            }

			let check_obj = null;
			if(Class.getObjectType)
			{
				if(Class.isDrawing && AscCommon.isRealObject(Class.parent))
				{
					let nParentNum = Class.getParentNum();
					if(nParentNum !== -1)
					{

						map[nParentNum] = true;
					}

					check_obj =
					{
						"type": c_oAscLockTypeElemPresentation.Object,
						"slideId": Class.parent.Get_Id(),
						"objId": Class.Get_Id(),
						"guid": Class.Get_Id()
					};
				}
				else if(AscFormat.isSlideLikeObject(Class))
				{
					check_obj =
					{
						"type": c_oAscLockTypeElemPresentation.Slide,
						"val": Class.Get_Id(),
						"guid": Class.Get_Id()
					};
				}
				else if(Class instanceof AscCommon.CComment){
					if(Class.Parent && Class.Parent.slide){
						if(Class.Parent.slide === presentation){
							check_obj =
							{
								"type": c_oAscLockTypeElemPresentation.Slide,
								"val": presentation.commentsLock.Get_Id(),
								"guid": presentation.commentsLock.Get_Id()
							};
						}
						else {
                            if(Class.Parent.slide.deleteLock){
                                check_obj =
                                {
                                    "type": c_oAscLockTypeElemPresentation.Object,
                                    "slideId": Class.Parent.slide.deleteLock.Get_Id(),
                                    "objId": Class.Get_Id(),
                                    "guid": Class.Get_Id()
                                };
                                map[Class.Parent.slide.num] = true;
                            }
						}
					}
				}
				if(check_obj)
					Asc.editor.CoAuthoringApi.releaseLocks( check_obj );
			}
		}


		if(Asc.editor.WordControl.m_oDrawingDocument.IsLockObjectsEnable)
		{
			for(let key in map)
			{
				if(map.hasOwnProperty(key))
				{
					num_arr.push(parseInt(key, 10));
				}
			}
			num_arr.sort(AscCommon.fSortAscending);
		}
		this.m_aNeedUnlock.length  = 0;
		this.m_aNeedUnlock2.length = 0;
	}

    if (0 < aChanges.length || null !== deleteIndex)
	{
		this.CoHistory.AddOwnChanges(aChanges2, deleteIndex);
		Asc.editor.CoAuthoringApi.saveChanges(aChanges, deleteIndex, AdditionalInfo, editor.canUnlockDocument2, bCollaborative);
        AscCommon.History.CanNotAddChanges = true;
    } else
		Asc.editor.CoAuthoringApi.unLockDocument(!!isAfterAskSave, editor.canUnlockDocument2, null, bCollaborative);
	Asc.editor.canUnlockDocument2 = false;

    if ( -1 === this.m_nUseType )
    {
        // Clear Undo/Redo only during collaborative editing
        AscCommon.History.Clear();
        AscCommon.History.SavedIndex = null;
    }
    else if ( 0 === this.m_nUseType )
    {
        // Clear Undo/Redo only during collaborative editing
        AscCommon.History.Clear();
        AscCommon.History.SavedIndex = null;

        this.m_nUseType = 1;
    }
    else
    {
        // Update the last save point in history
        AscCommon.History.Reset_SavedIndex(IsUserSave);
    }

    for(let i = 0; i < num_arr.length; ++i)
    {
		Asc.editor.WordControl.m_oDrawingDocument.OnRecalculateSlide(num_arr[i]);
    }
    if(num_arr.length > 0)
    {
		Asc.editor.WordControl.m_oDrawingDocument.OnEndRecalculate();
    }
    let oSlide = presentation.GetCurrentSlide();
    if(oSlide && oSlide.notesShape){
		Asc.editor.WordControl.m_oDrawingDocument.Notes_OnRecalculate(presentation.CurPage, oSlide.NotesWidth, oSlide.getNotesHeight());
    }
    presentation.Document_UpdateInterfaceState();
    presentation.Document_UpdateUndoRedoState();

    // drawingDocument.ClearCachePages();
    //    drawingDocument.FirePaint();
};

CCollaborativeEditing.prototype.Release_Locks = function()
{
    let map_redraw = {};
    let UnlockCount = this.m_aNeedUnlock.length;
	let presentation = this.GetPresentation();
    for ( let Index = 0; Index < UnlockCount; Index++ )
    {
        let CurLockType = this.m_aNeedUnlock[Index].Lock.Get_Type();
		let Class =  this.m_aNeedUnlock[Index];
        if  ( AscCommon.c_oAscLockTypes.kLockTypeOther3 != CurLockType && AscCommon.c_oAscLockTypes.kLockTypeOther != CurLockType )
        {
			Class.Lock.Set_Type( AscCommon.c_oAscLockTypes.kLockTypeNone, false);
            if ( Class instanceof AscCommonSlide.PropLocker )
            {
				let object = AscCommon.g_oTableId.Get_ById(Class.objectId);
                if(object instanceof AscCommonSlide.CPresentation)
                {
                    if(Class === presentation.themeLock)
                    {
						Asc.editor.sendEvent("asc_onUnLockDocumentTheme");
                    }
                    else if(Class === presentation.schemeLock)
                    {
						Asc.editor.sendEvent("asc_onUnLockDocumentSchema");
                    }
                    else if(Class === presentation.slideSizeLock)
                    {
						Asc.editor.sendEvent("asc_onUnLockDocumentProps");
                    }
                    else if(Class === presentation.viewPrLock)
                    {
						Asc.editor.sendEvent("asc_onUnLockViewProps");
                    }
                    else if (Class === presentation.hdrFtrLock)
                    {
						Asc.editor.sendEvent("asc_onUnLockSlideHdrFtrApplyToAll");
                    }
                }
            }
            if(Class instanceof AscCommon.CComment)
            {
				Asc.editor.sync_UnLockComment(Class.Get_Id());
                if(Class.Parent && Class.Parent.slide && presentation !== Class.Parent.slide)
                {
                    map_redraw[Class.Parent.slide.num] = true;
                }
            }
            if(Class instanceof AscCommon.CCore)
            {
				Asc.editor.sendEvent("asc_onLockCore", false);
            }
        }
        else if ( AscCommon.c_oAscLockTypes.kLockTypeOther3 === CurLockType )
        {
			Class.Lock.Set_Type( AscCommon.c_oAscLockTypes.kLockTypeOther, false);
        }
		if(Class.isDrawing)
		{
			let nIdx = Class.getParentNum();
			if(nIdx !== -1)
			{
				map_redraw[nIdx] = true;
			}
		}
    }
    return map_redraw;
};

CCollaborativeEditing.prototype.OnEnd_Load_Objects = function()
{
    // This function is called when external objects (images and fonts) are loaded

    // Release the lock
    AscCommon.CollaborativeEditing.Set_GlobalLock(false);
    AscCommon.CollaborativeEditing.Set_GlobalLockSelection(false);

    // Start full document recalculation
    let LogicDocument = this.GetPresentation();

		const arrChanges = this.CoHistory.GetAllChanges();
		const oRecalcData = LogicDocument.History.Get_RecalcData(null, arrChanges, this.m_nRecalcIndexStart, this.m_nRecalcIndexEnd);
    LogicDocument.Recalculate(oRecalcData);
    LogicDocument.Document_UpdateSelectionState();
    LogicDocument.Document_UpdateInterfaceState();

	Asc.editor.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.ApplyChanges);
};

CCollaborativeEditing.prototype.OnEnd_CheckLock = function(DontLockInFastMode)
{
	let drawingDocument = this.GetDrawingDocument();
    let aIds = [];
	let Count = this.m_aCheckLocks.length;
    for ( let Index = 0; Index < Count; Index++ )
    {
		let oItem = this.m_aCheckLocks[Index];

        if ( true === oItem ) // must compare by value and type
            return true;
        else if ( false !== oItem )
            aIds.push( oItem );
    }


    if (true === DontLockInFastMode && true === this.Is_Fast())
        return false;

    if ( aIds.length > 0 )
    {
        // Send request to server with list of Ids
		Asc.editor.CoAuthoringApi.askLock( aIds, this.OnCallback_AskLock );

        // Set global lock only during collaborative editing
        if ( -1 === this.m_nUseType )
		{
			this.Set_GlobalLock(true);
		}
        else
        {
            // Iterate through array and mark as locked by us
            Count = this.m_aCheckLocks.length;
            for ( let Index = 0; Index < Count; Index++ )
            {
				let oItem = this.m_aCheckLocks[Index];
				let items = [];
                switch(oItem["type"])
                {
                    case c_oAscLockTypeElemPresentation.Object:
                    {
                        items.push(oItem["objId"]);
                        items.push(oItem["slideId"]);
                        break;
                    }
                    case c_oAscLockTypeElemPresentation.Slide:
                    {
                        items.push(oItem["val"]);
                        break;
                    }
                    case c_oAscLockTypeElemPresentation.Presentation:
                    {
                        break;
                    }
                }

                for(let i = 0; i < items.length; ++i)
                {
					let item = items[i];
                    if ( true !== item && false !== item ) // must compare by value and type
                    {
						let Class = AscCommon.g_oTableId.Get_ById( item );
                        if ( null != Class )
                        {
                            Class.Lock.Set_Type( AscCommon.c_oAscLockTypes.kLockTypeMine, false );
                            this.Add_Unlock2( Class );
                        }
                    }
                }
            }

            this.m_aCheckLocks.length = 0;
        }
    }

    return false;
};
CCollaborativeEditing.prototype.OnCallback_AskLock = function(result)
{
	let presentation = AscCommon.CollaborativeEditing.GetPresentation();
    if (true === AscCommon.CollaborativeEditing.Get_GlobalLock())
    {
        if (!Asc.editor.checkLongActionCallback(AscCommon.CollaborativeEditing.OnCallback_AskLock, result))
            return;

        // Release global lock
        AscCommon.CollaborativeEditing.Set_GlobalLock(false);

        if (result["lock"])
        {
            // Iterate through array and mark as locked by us

			let Count = AscCommon.CollaborativeEditing.m_aCheckLocks.length;
            for ( let Index = 0; Index < Count; Index++ )
            {
				let oItem = AscCommon.CollaborativeEditing.m_aCheckLocks[Index];
				let item;
                switch(oItem["type"])
                {
                    case c_oAscLockTypeElemPresentation.Object:
                    {
                        item = oItem["objId"];
                        break;
                    }
                    case c_oAscLockTypeElemPresentation.Slide:
                    {
                        item = oItem["val"];
                        break;
                    }
                    case c_oAscLockTypeElemPresentation.Presentation:
                    {
                        break;
                    }
                }
                if ( true !== oItem && false !== oItem ) // must compare by value and type
                {
					let Class = AscCommon.g_oTableId.Get_ById( item );
                    if ( null != Class )
                    {
                        Class.Lock.Set_Type( AscCommon.c_oAscLockTypes.kLockTypeMine );
                        AscCommon.CollaborativeEditing.Add_Unlock2( Class );
                    }
                }
            }
        }
        else if (result["error"])
        {
            // If chart editing started but we were told it cannot be edited,
            // send message to close the chart editor.
            if (Asc.editor.frameManager.isLoadingChartEditor)
                Asc.editor.sync_closeChartEditor();

            if (Asc.editor.frameManager.isLoadingOleEditor)
              Asc.editor.sync_closeOleEditor();

            // Rollback one step and remove this last point from Undo/Redo
			presentation.Document_Undo();
            AscCommon.History.Clear_Redo();
        }

    }
	Asc.editor.frameManager.endLoadChartEditor();
	Asc.editor.frameManager.endLoadOleEditor();
};

CCollaborativeEditing.prototype.AddPosExtChanges = function(Item, ChangeObject)
{
    if(ChangeObject.IsHorizontal())
    {
        this.PosExtChangesX.push(Item);
    }
    else
    {
        this.PosExtChangesY.push(Item);
    }
};


CCollaborativeEditing.prototype.RewritePosExtChanges = function(changesArr, scale)
{
    for(let i = 0; i < changesArr.length; ++i)
    {
		let changes = changesArr[i];
		let data = changes.Data;
        data.New *= scale;
        data.Old *= scale;
		let Binary_Writer = AscCommon.History.BinaryWriter;
		let Binary_Pos    = Binary_Writer.GetCurPosition();
		if (Asc.editor.binaryChanges) {
			Binary_Writer.WriteWithLen(this, function () {
				Binary_Writer.WriteString2(changes.Class.Get_Id());
				Binary_Writer.WriteLong(changes.Data.Type);
				changes.Data.WriteToBinary(Binary_Writer);
			});
		} else {
			Binary_Writer.WriteString2(changes.Class.Get_Id());
			Binary_Writer.WriteLong(changes.Data.Type);
			changes.Data.WriteToBinary(Binary_Writer);
		}

		let Binary_Len = Binary_Writer.GetCurPosition() - Binary_Pos;

        changes.Binary.Pos = Binary_Pos;
        changes.Binary.Len = Binary_Len;
    }
};

CCollaborativeEditing.prototype.RefreshPosExtChanges = function()
{
    if(this.ScaleX != null && this.ScaleY != null)
    {
        this.RewritePosExtChanges(this.PosExtChangesX, this.ScaleX);
        this.RewritePosExtChanges(this.PosExtChangesY, this.ScaleY);
    }
    this.PosExtChangesX.length = 0;
    this.PosExtChangesY.length = 0;
    this.ScaleX = null;
    this.ScaleY = null;
};

CCollaborativeEditing.prototype.Update_ForeignCursorsPositions = function()
{
	let DrawingDocument = this.GetDrawingDocument();
	let oPresentation = this.GetPresentation();
	let oTargetDocContentOrTable;
	let oCurController = oPresentation.GetCurrentController();
    if(oCurController){
        oTargetDocContentOrTable = oCurController.getTargetDocContent(undefined, true);
    }
    if(!oTargetDocContentOrTable){
        for (let UserId in this.m_aForeignCursors){
            DrawingDocument.Collaborative_RemoveTarget(UserId);
        }
        return;
    }
	let bTable = (oTargetDocContentOrTable instanceof AscCommonWord.CTable);
    for (let UserId in this.m_aForeignCursors){
		let DocPos = this.m_aForeignCursors[UserId];
        if (!DocPos || DocPos.length <= 0)
            continue;

        this.m_aForeignCursorsPos.Update_DocumentPosition(DocPos);

        let Run      = DocPos[DocPos.length - 1].Class;
		let InRunPos = DocPos[DocPos.length - 1].Position;
        this.Update_ForeignCursorPosition(UserId, Run, InRunPos, false, oTargetDocContentOrTable, bTable);
    }
};

CCollaborativeEditing.prototype.Update_ForeignCursorPosition = function(UserId, Run, InRunPos, isRemoveLabel, oTargetDocContentOrTable, bTable){
    if (!(Run instanceof AscCommonWord.ParaRun))
        return;

    let DrawingDocument = this.GetDrawingDocument();
    let oPresentation = this.GetPresentation();
    let Paragraph = Run.GetParagraph();
    if (!Paragraph || !Paragraph.Parent){
        DrawingDocument.Collaborative_RemoveTarget(UserId);
        return;
    }

    if(!bTable){
        if(oTargetDocContentOrTable !== Paragraph.Parent){
            DrawingDocument.Collaborative_RemoveTarget(UserId);
            return;
        }
    }
    else{
        if(!Paragraph.Parent.Parent || !Paragraph.Parent.Parent.Row ||
            !Paragraph.Parent.Parent.Row.Table || Paragraph.Parent.Parent.Row.Table !== oTargetDocContentOrTable){
            DrawingDocument.Collaborative_RemoveTarget(UserId);
            return;
        }
    }

	let ParaContentPos = Paragraph.Get_PosByElement(Run);
    if (!ParaContentPos){
        DrawingDocument.Collaborative_RemoveTarget(UserId);
        return;
    }
    ParaContentPos.Update(InRunPos, ParaContentPos.GetDepth() + 1);
	let XY = Paragraph.Get_XYByContentPos(ParaContentPos);
    if (XY && XY.Height > 0.001){
		let ShortId = this.m_aForeignCursorsId[UserId] ? this.m_aForeignCursorsId[UserId] : UserId;
        DrawingDocument.Collaborative_UpdateTarget(UserId, ShortId, XY.X, XY.Y, XY.Height, oPresentation.CurPage, Paragraph.Get_ParentTextTransform());
        this.Add_ForeignCursorXY(UserId, XY.X, XY.Y, XY.PageNum, XY.Height, Paragraph, isRemoveLabel);
        if (true === this.m_aForeignCursorsToShow[UserId]){
            this.Show_ForeignCursorLabel(UserId);
            this.Remove_ForeignCursorToShow(UserId);
        }
    }
    else{
        DrawingDocument.Collaborative_RemoveTarget(UserId);
        this.Remove_ForeignCursorXY(UserId);
        this.Remove_ForeignCursorToShow(UserId);
    }
};

CCollaborativeEditing.prototype.private_RecalculateDocument = function(arrChanges){
    this.GetPresentation().Recalculate(AscCommon.History.Get_RecalcData(null, arrChanges));
};

CCollaborativeEditing.prototype.private_UpdateForeignCursor = function(CursorInfo, UserId, Show, UserShortId)
{
    this.GetPresentation().Update_ForeignCursor(CursorInfo, UserId, Show, UserShortId);
};
CCollaborativeEditing.prototype._PreUndo = function() {
	let logicDocument = this.m_oLogicDocument;

	logicDocument.DrawingDocument.EndTrackTable(null, true);
	logicDocument.TurnOffCheckChartSelection();

	return this.private_SaveDocumentState();
};
CCollaborativeEditing.prototype._PostUndo = function(state, changes) {
	this.private_RestoreDocumentState(state);
	this.private_RecalculateDocument(changes);

	let logicDocument = this.m_oLogicDocument;
	logicDocument.TurnOnCheckChartSelection();
	logicDocument.UpdateSelection();
	logicDocument.UpdateInterface();
	logicDocument.UpdateRulers();
};

//--------------------------------------------------------export----------------------------------------------------
window['AscCommon'] = window['AscCommon'] || {};
window['AscCommon'].SlideCollaborativeEditing = CCollaborativeEditing;
