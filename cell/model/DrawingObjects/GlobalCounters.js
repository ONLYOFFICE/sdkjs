/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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

(function(window, undefined){

    /**
     *
     * @constructor
     * @extends {AscCommon.CCollaborativeEditingBase}
     */
    function CCollaborativeEditing()
    {
        AscCommon.CCollaborativeEditingBase.call(this);
        this.WaitImages = {};
    }

    CCollaborativeEditing.prototype = Object.create(AscCommon.CCollaborativeEditingBase.prototype);
    CCollaborativeEditing.prototype.constructor = CCollaborativeEditing;

    CCollaborativeEditing.prototype.GetEditorApi = function()
    {
        return Asc.editor;
    };
    CCollaborativeEditing.prototype.GetDrawingDocument = function()
    {
        var oApi = this.GetEditorApi();
        if(!oApi)
        {
            return null;
        }
        var oWB = oApi.wbModel;
        if(!oWB)
        {
            return null;
        }
        return oWB.DrawingDocument;
    };

    CCollaborativeEditing.prototype.Have_OtherChanges = function()
    {
        return false;
    };

    CCollaborativeEditing.prototype.Start_CollaborationEditing = function()
    {
    };

    CCollaborativeEditing.prototype.Add_User = function(UserId)
    {
    };

    CCollaborativeEditing.prototype.Find_User = function(UserId)
    {
    };

    CCollaborativeEditing.prototype.Remove_User = function(UserId)
    {
    };

    CCollaborativeEditing.prototype.Add_Changes = function(Changes)
    {
    };

    CCollaborativeEditing.prototype.Add_Unlock = function(LockClass)
    {
    };

    CCollaborativeEditing.prototype.Add_Unlock2 = function(Lock)
    {
    };

    CCollaborativeEditing.prototype.Apply_OtherChanges = function()
    {
    };


    CCollaborativeEditing.prototype.Apply_Changes = function()
    {
    };

    CCollaborativeEditing.prototype.Send_Changes = function()
    {
    };

    CCollaborativeEditing.prototype.Release_Locks = function()
    {
    };

    CCollaborativeEditing.prototype.OnStart_Load_Objects = function()
    {
    };

    CCollaborativeEditing.prototype.OnEnd_Load_Objects = function()
    {
    };
    //-----------------------------------------------------------------------------------
    // Функции для работы с ссылками, у новых объектов
    //-----------------------------------------------------------------------------------
    CCollaborativeEditing.prototype.Apply_LinkData = function()
    {
        AscCommon.CCollaborativeEditingBase.prototype.Apply_LinkData.call(this);
        //stub for asc_nativeApplyChanges2(Load_Images clears m_aNewImages) remove it together with Load_Images above
        if (window["NATIVE_EDITOR_ENJINE"] === true)
        {
            window["NATIVE_EDITOR_ENJINE_NEW_IMAGES"] = this.m_aNewImages.concat();
        }
        this.Load_Images();
    };
    CCollaborativeEditing.prototype.CheckWaitingImages = function(aImages)
    {
        this.WaitImages = {};
        for(var i = 0; i < aImages.length; ++i)
        {
            this.WaitImages[aImages] = 1;
        }
    };
    CCollaborativeEditing.prototype.SendImagesCallback = function (aImages)
    {
        var oApi = Asc['editor'], bOldVal;
        if(aImages.length > 0)
        {
            bOldVal =  oApi.ImageLoader.bIsAsyncLoadDocumentImages;
            oApi.ImageLoader.bIsAsyncLoadDocumentImages = true;
            oApi.ImageLoader.LoadDocumentImages(aImages);
            oApi.ImageLoader.bIsAsyncLoadDocumentImages = bOldVal;
            this.WaitImages = {};
        }
    };
    CCollaborativeEditing.prototype.Load_Images = function(){
        var aImages = this.CollectImagesFromChanges();
        if(aImages.length > 0)
        {
            this.SendImagesUrlsFromChanges(aImages);
        }
        else
        {
            this.SendImagesCallback([].concat(this.m_aNewImages));
            this.m_aNewImages.length = 0;
    }
    };



    //--------------------------------------------------
    CCollaborativeEditing.prototype.Update_ForeignCursorsPositions = function()
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
    CCollaborativeEditing.prototype.Update_ForeignCursorPosition = function(UserId, Run, InRunPos, isRemoveLabel)
    {
        var DrawingDocument = Asc.editor.wbModel.DrawingDocument;

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
        ParaContentPos.Update(InRunPos, ParaContentPos.Get_Depth() + 1);

        var XY = Paragraph.Get_XYByContentPos(ParaContentPos);
        if (XY && XY.Height > 0.001)
        {
            var ShortId = this.m_aForeignCursorsId[UserId] ? this.m_aForeignCursorsId[UserId] : UserId;
            var sWSId = null;
            var oShape = Paragraph.Parent && Paragraph.Parent.Is_DrawingShape(true);
            if(oShape)
            {
                if(oShape.worksheet)
                {
                    sWSId = oShape.worksheet.Id;
                }
                else
                {
                    if(oShape.chart)
                    {
                        if(oShape.chart.worksheet)
                        {
                            sWSId = oShape.chart.worksheet.Id;
                        }
                    }
                }
            }
            DrawingDocument.Collaborative_UpdateTarget(UserId, ShortId, XY.X, XY.Y, XY.Height, sWSId, Paragraph.Get_ParentTextTransform());
            this.Add_ForeignCursorXY(UserId, XY.X, XY.Y, XY.PageNum, XY.Height, Paragraph, isRemoveLabel, sWSId);

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

    CCollaborativeEditing.prototype.Check_ForeignCursorsLabels = function(X, Y, SheetId) {
        var DrawingDocument = this.GetDrawingDocument();
        if(!DrawingDocument) {
            return;
        }
        var Px7 = DrawingDocument.GetMMPerDot(7);
        var Px3 = DrawingDocument.GetMMPerDot(3);

        for (var UserId in this.m_aForeignCursorsXY) {
            var Cursor = this.m_aForeignCursorsXY[UserId];
            if ((true === Cursor.Transform && Cursor.SheetId === SheetId && Cursor.X0 - Px3 < X && X < Cursor.X1 + Px3 && Cursor.Y0 - Px3 < Y && Y < Cursor.Y1 + Px3)
                || (Math.abs(X - Cursor.X) < Px7 && Cursor.Y - Px3 < Y && Y < Cursor.Y + Cursor.H + Px3 && Cursor.SheetId === SheetId)) {
                this.Show_ForeignCursorLabel(UserId);
            }
        }
    };
    //-----------------------------------------------------------------------------------
    // Функции для проверки корректности новых изменений
    //-----------------------------------------------------------------------------------
    CCollaborativeEditing.prototype.Check_MergeData = function()
    {
    };
    //-----------------------------------------------------------------------------------
    // Функции для проверки залоченности объектов
    //-----------------------------------------------------------------------------------
    CCollaborativeEditing.prototype.Get_GlobalLock = function()
    {
        return false;
    };
    CCollaborativeEditing.prototype.Set_GlobalLock = function(isLock)
    {

    };
    CCollaborativeEditing.prototype.Get_GlobalLockSelection = function()
    {
        return false;
    };
    CCollaborativeEditing.prototype.Set_GlobalLockSelection = function(isLock)
    {

    };

    CCollaborativeEditing.prototype.OnStart_CheckLock = function()
    {
    };

    CCollaborativeEditing.prototype.Add_CheckLock = function(oItem)
    {
    };

    CCollaborativeEditing.prototype.OnEnd_CheckLock = function()
    {
    };

    CCollaborativeEditing.prototype.OnCallback_AskLock = function(result)
    {
    };
    //-----------------------------------------------------------------------------------
    // Функции для работы с залоченными объектами, которые еще не были добавлены
    //-----------------------------------------------------------------------------------
    CCollaborativeEditing.prototype.Reset_NeedLock = function()
    {
    };

    CCollaborativeEditing.prototype.Add_NeedLock = function(Id, sUser)
    {
    };

    CCollaborativeEditing.prototype.Remove_NeedLock = function(Id)
    {
    };

    CCollaborativeEditing.prototype.Lock_NeedLock = function()
    {
    };
    //-----------------------------------------------------------------------------------
    // Функции для работы с новыми объектами, созданными на других клиентах
    //-----------------------------------------------------------------------------------
    CCollaborativeEditing.prototype.Clear_NewObjects = function()
    {
    };

    CCollaborativeEditing.prototype.Add_NewObject = function(Class)
    {
    };

    CCollaborativeEditing.prototype.OnEnd_ReadForeignChanges = function()
    {
    };
    CCollaborativeEditing.prototype.private_UpdateForeignCursor = function(CursorInfo, UserId, Show, UserShortId)
    {
        var oApi = this.GetEditorApi();
        if(!oApi)
        {
            return;
        }
        var oWB = oApi.wb;
        if(!oWB)
        {
            return;
        }
        oWB.Update_ForeignCursor(CursorInfo, UserId, Show, UserShortId);
    };
    //--------------------------------------------------------export----------------------------------------------------
    window['AscCommon'] = window['AscCommon'] || {};
    window['AscCommon'].CollaborativeEditing = new CCollaborativeEditing();
})(window);
