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

AscDFH.changesFactory[AscDFH.historyitem_Endnotes_AddEndnote]    = CChangesEndnotesAddEndnote;
AscDFH.changesFactory[AscDFH.historyitem_Endnotes_RemoveEndnote] = CChangesEndnotesRemoveEndnote;

//----------------------------------------------------------------------------------------------------------------------
// Карта зависимости изменений
//----------------------------------------------------------------------------------------------------------------------
AscDFH.changesRelationMap[AscDFH.historyitem_Endnotes_AddEndnote]    = [AscDFH.historyitem_Endnotes_AddEndnote];
AscDFH.changesRelationMap[AscDFH.historyitem_Endnotes_RemoveEndnote] = [AscDFH.historyitem_Endnotes_RemoveEndnote];
//----------------------------------------------------------------------------------------------------------------------

/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesEndnotesAddEndnote(Class, Id)
{
	AscDFH.CChangesBase.call(this, Class);

	this.Id = Id;
}
CChangesEndnotesAddEndnote.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesEndnotesAddEndnote.prototype.constructor = CChangesEndnotesAddEndnote;
CChangesEndnotesAddEndnote.prototype.Type = AscDFH.historyitem_Endnotes_AddEndnote;
CChangesEndnotesAddEndnote.prototype.Undo = function()
{
	delete this.Class.Endnote[this.Id];
};
CChangesEndnotesAddEndnote.prototype.Redo = function()
{
	this.Class.Endnote[this.Id] = AscCommon.g_oTableId.Get_ById(this.Id);
};
CChangesEndnotesAddEndnote.prototype.WriteToBinary = function(Writer)
{
	// String : Id
	Writer.WriteString2(this.Id);
};
CChangesEndnotesAddEndnote.prototype.ReadFromBinary = function(Reader)
{
	// String : Id
	this.Id = Reader.GetString2();
};
CChangesEndnotesAddEndnote.prototype.CreateReverseChange = function()
{
	return null;
};
CChangesEndnotesAddEndnote.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return true;

	if (this.Type === oChange.Type && this.Id === oChange.Id)
		return false;

	return true;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesEndnotesRemoveEndnote(Class, Id)
{
	AscDFH.CChangesBase.call(this, Class);

	this.Id = Id;
}
CChangesEndnotesRemoveEndnote.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesEndnotesRemoveEndnote.prototype.constructor = CChangesEndnotesRemoveEndnote;
CChangesEndnotesRemoveEndnote.prototype.Type = AscDFH.historyitem_Endnotes_RemoveEndnote;
CChangesEndnotesRemoveEndnote.prototype.Undo = function()
{
	this.Class.Endnote[this.Id] = AscCommon.g_oTableId.Get_ById(this.Id);
};
CChangesEndnotesRemoveEndnote.prototype.Redo = function()
{
	delete this.Class.Endnote[this.Id];
};
CChangesEndnotesRemoveEndnote.prototype.WriteToBinary = function(Writer)
{
	// String : Id
	Writer.WriteString2(this.Id);
};
CChangesEndnotesRemoveEndnote.prototype.ReadFromBinary = function(Reader)
{
	// String : Id
	this.Id = Reader.GetString2();
};
CChangesEndnotesRemoveEndnote.prototype.CreateReverseChange = function()
{
	return new CChangesEndnotesAddEndnote(this.Class, this.Id);
};
CChangesEndnotesRemoveEndnote.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return true;

	if (this.Type === oChange.Type && this.Id === oChange.Id)
		return false;

	return true;
};
