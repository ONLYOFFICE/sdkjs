/*
 *
 * (c) Copyright Ascensio System Limited 2010-2016
 *
 * This program is freeware. You can redistribute it and/or modify it under the terms of the GNU 
 * General Public License (GPL) version 3 as published by the Free Software Foundation (https://www.gnu.org/copyleft/gpl.html). 
 * In accordance with Section 7(a) of the GNU GPL its Section 15 shall be amended to the effect that 
 * Ascensio System SIA expressly excludes the warranty of non-infringement of any third-party rights.
 *
 * THIS PROGRAM IS DISTRIBUTED WITHOUT ANY WARRANTY; WITHOUT EVEN THE IMPLIED WARRANTY OF MERCHANTABILITY OR
 * FITNESS FOR A PARTICULAR PURPOSE. For more details, see GNU GPL at https://www.gnu.org/copyleft/gpl.html
 *
 * You can contact Ascensio System SIA by email at sales@onlyoffice.com
 *
 * The interactive user interfaces in modified source and object code versions of ONLYOFFICE must display 
 * Appropriate Legal Notices, as required under Section 5 of the GNU GPL version 3.
 *
 * Pursuant to Section 7  3(b) of the GNU GPL you must retain the original ONLYOFFICE logo which contains 
 * relevant author attributions when distributing the software. If the display of the logo in its graphic 
 * form is not reasonably feasible for technical reasons, you must include the words "Powered by ONLYOFFICE" 
 * in every copy of the program you distribute. 
 * Pursuant to Section 7  3(e) we decline to grant you any rights under trademark law for use of our trademarks.
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
    CCollaborativeEditing.superclass.constructor.call(this);
}

AscCommon.extendClass(CCollaborativeEditing, AscCommon.CCollaborativeEditingBase);

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
CCollaborativeEditing.prototype.Clear_LinkData = function()
{
    this.m_aLinkData.length = 0;
};

CCollaborativeEditing.prototype.Add_LinkData = function(Class, LinkData)
{
    this.m_aLinkData.push( { Class : Class, LinkData : LinkData } );
};

CCollaborativeEditing.prototype.Apply_LinkData = function()
{
    var Count = this.m_aLinkData.length;
    for ( var Index = 0; Index < Count; Index++ )
    {
        var Item = this.m_aLinkData[Index];
        Item.Class.Load_LinkData( Item.LinkData );
    }
    this.Clear_LinkData();
    this.Load_Images();
};
CCollaborativeEditing.prototype.Load_Images = function(){
    if(this.m_aNewImages.length > 0)
    {
        var old_val =  Asc["editor"].ImageLoader.bIsAsyncLoadDocumentImages;
        Asc["editor"].ImageLoader.bIsAsyncLoadDocumentImages = true;
        Asc["editor"].ImageLoader.LoadDocumentImages(this.m_aNewImages, null);
        AscCommon.CollaborativeEditing.m_aNewImages.length = 0;
        Asc["editor"].ImageLoader.bIsAsyncLoadDocumentImages = old_val;
    }
}
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

//-----------------------------------------------------------------------------------
// Функции для работы с массивом m_aDC
//-----------------------------------------------------------------------------------
CCollaborativeEditing.prototype.Add_NewDC = function(Class)
{
};

CCollaborativeEditing.prototype.Clear_DCChanges = function()
{
};

CCollaborativeEditing.prototype.Refresh_DCChanges = function()
{
};
//-----------------------------------------------------------------------------------
// Функции для работы с отметками изменений
//-----------------------------------------------------------------------------------

CCollaborativeEditing.prototype.Clear_CollaborativeMarks = function()
{
    for ( var Id in this.m_aChangedClasses )
    {
        this.m_aChangedClasses[Id].Clear_CollaborativeMarks();
    }

    // Очищаем массив
    this.m_aChangedClasses = {};
};

    //--------------------------------------------------------export----------------------------------------------------
    window['AscCommon'] = window['AscCommon'] || {};
    window['AscCommon'].CollaborativeEditing = new CCollaborativeEditing();
})(window);
