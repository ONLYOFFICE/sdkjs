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

AscDFH.changesFactory[AscDFH.historyitem_type_ChangeCustomXML]    = CChangesCustomXMLChange
AscDFH.changesFactory[AscDFH.historyitem_type_CustomXML]       = CChangesCustomXmlAdd;

AscDFH.changesRelationMap[AscDFH.historyitem_type_ChangeCustomXML] = [AscDFH.historyitem_type_ChangeCustomXML];
AscDFH.changesRelationMap[AscDFH.historyitem_type_CustomXML]    = [AscDFH.historyitem_type_CustomXML];

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesCustomXMLChange(Class, Old, New)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New);
}
CChangesCustomXMLChange.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesCustomXMLChange.prototype.constructor = CChangesCustomXMLChange;
CChangesCustomXMLChange.prototype.Type = AscDFH.historyitem_Comment_Change;
CChangesCustomXMLChange.prototype.WriteToBinary = function(Writer)
{
	// Variable : New data
	// Variable : Old data

	this.New.Write_ToBinary2(Writer);
	this.Old.Write_ToBinary2(Writer);
};
CChangesCustomXMLChange.prototype.ReadFromBinary = function(Reader)
{
	// Variable : New data
	// Variable : Old data

	debugger
	this.New = new AscCommon.CCommentData();
	this.Old = new AscCommon.CCommentData();
	this.New.Read_FromBinary2(Reader);
	this.Old.Read_FromBinary2(Reader);
};
CChangesCustomXMLChange.prototype.private_SetValue = function(Value)
{
	this.Class.Data = Value;
	editor.sync_ChangeCommentData(this.Class.Id, Value);
};


/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesCustomXmlAdd(Class, Id, xml)
{
	AscDFH.CChangesBase.call(this, Class);

	this.Id		= Id;
	this.xml	= xml;
}
CChangesCustomXmlAdd.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesCustomXmlAdd.prototype.constructor = CChangesCustomXmlAdd;
CChangesCustomXmlAdd.prototype.Type = AscDFH.historyitem_type_CustomXML;
CChangesCustomXmlAdd.prototype.Undo = function()
{
	// var oComment = this.Class.m_arrCommentsById[this.Id];
	// if (oComment)
	// {
	// 	delete this.Class.m_arrCommentsById[this.Id];
	// 	var oChangedComments = this.Class.UpdateCommentPosition(oComment);
	// 	editor.sync_RemoveComment(this.Id);
	// 	editor.sync_ChangeCommentLogicalPosition(oChangedComments, this.Class.GetCommentsPositionsCount());
	// }
};
CChangesCustomXmlAdd.prototype.Redo = function()
{
	// this.Class.m_arrCommentsById[this.Id] = this.Comment;
	// var oChangedComments = this.Class.UpdateCommentPosition(this.Comment);
	// editor.sync_AddComment(this.Id, this.Comment.Data);
	// editor.sync_ChangeCommentLogicalPosition(oChangedComments, this.Class.GetCommentsPositionsCount());
};
CChangesCustomXmlAdd.prototype.WriteToBinary = function(Writer)
{
	// String : Id комментария
	Writer.WriteString2(this.Id);
};
CChangesCustomXmlAdd.prototype.ReadFromBinary = function(Reader)
{
	// String : Id комментария
	this.Id      = Reader.GetString2();
	this.xml = AscCommon.g_oTableId.Get_ById(this.Id);
};
CChangesCustomXmlAdd.prototype.CreateReverseChange = function()
{
	//return new CChangesCommentsRemove(this.Class, this.Id, this.Comment);
};
CChangesCustomXmlAdd.prototype.Merge = function(oChange)
{

};
