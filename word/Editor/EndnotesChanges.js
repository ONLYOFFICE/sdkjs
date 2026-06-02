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
