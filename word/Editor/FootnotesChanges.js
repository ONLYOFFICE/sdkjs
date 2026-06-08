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

AscDFH.changesFactory[AscDFH.historyitem_Footnotes_AddFootnote]              = CChangesFootnotesAddFootnote;
AscDFH.changesFactory[AscDFH.historyitem_Footnotes_SetSeparator]             = CChangesFootnotesSetSeparator;
AscDFH.changesFactory[AscDFH.historyitem_Footnotes_SetContinuationSeparator] = CChangesFootnotesSetContinuationSeparator;
AscDFH.changesFactory[AscDFH.historyitem_Footnotes_SetContinuationNotice]    = CChangesFootnotesSetContinuationNotice;
AscDFH.changesFactory[AscDFH.historyitem_Footnotes_RemoveFootnote]           = CChangesFootnotesRemoveFootnote;

//----------------------------------------------------------------------------------------------------------------------
// Карта зависимости изменений
//----------------------------------------------------------------------------------------------------------------------
AscDFH.changesRelationMap[AscDFH.historyitem_Footnotes_AddFootnote]              = [AscDFH.historyitem_Footnotes_AddFootnote];
AscDFH.changesRelationMap[AscDFH.historyitem_Footnotes_SetSeparator]             = [AscDFH.historyitem_Footnotes_SetSeparator];
AscDFH.changesRelationMap[AscDFH.historyitem_Footnotes_SetContinuationSeparator] = [AscDFH.historyitem_Footnotes_SetContinuationSeparator];
AscDFH.changesRelationMap[AscDFH.historyitem_Footnotes_SetContinuationNotice]    = [AscDFH.historyitem_Footnotes_SetContinuationNotice];
AscDFH.changesRelationMap[AscDFH.historyitem_Footnotes_RemoveFootnote]           = [AscDFH.historyitem_Footnotes_RemoveFootnote];
//----------------------------------------------------------------------------------------------------------------------

/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesFootnotesAddFootnote(Class, Id)
{
	AscDFH.CChangesBase.call(this, Class);

	this.Id = Id;
}
CChangesFootnotesAddFootnote.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesFootnotesAddFootnote.prototype.constructor = CChangesFootnotesAddFootnote;
CChangesFootnotesAddFootnote.prototype.Type = AscDFH.historyitem_Footnotes_AddFootnote;
CChangesFootnotesAddFootnote.prototype.Undo = function()
{
	delete this.Class.Footnote[this.Id];
};
CChangesFootnotesAddFootnote.prototype.Redo = function()
{
	this.Class.Footnote[this.Id] = AscCommon.g_oTableId.Get_ById(this.Id);
};
CChangesFootnotesAddFootnote.prototype.WriteToBinary = function(Writer)
{
	// String : Id
	Writer.WriteString2(this.Id);
};
CChangesFootnotesAddFootnote.prototype.ReadFromBinary = function(Reader)
{
	// String : Id
	this.Id = Reader.GetString2();
};
CChangesFootnotesAddFootnote.prototype.CreateReverseChange = function()
{
	return null;
};
CChangesFootnotesAddFootnote.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return true;

	if (this.Type === oChange.Type && this.Id === oChange.Id)
		return false;

	return true;
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesFootnotesSetSeparator(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesFootnotesSetSeparator.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesFootnotesSetSeparator.prototype.constructor = CChangesFootnotesSetSeparator;
CChangesFootnotesSetSeparator.prototype.Type = AscDFH.historyitem_Footnotes_SetSeparator;
CChangesFootnotesSetSeparator.prototype.private_SetValue = function(Value)
{
	this.Class.Separator = Value;
};
CChangesFootnotesSetSeparator.prototype.WriteToBinary = function(Writer)
{
	// Long : Flags
	// 1bit : is new undefined?
	// 2bit : is old undefined?
	// String : New id (1 bit zero)
	// String : Old id (2 bit zero)

	var nFlags = 0;
	if (!this.New)
		nFlags |= 1;
	if (!this.Old)
		nFlags |= 2;

	Writer.WriteLong(nFlags);

	if (this.New)
		Writer.WriteString2(this.New.Get_Id());

	if (this.Old)
		Writer.WriteString2(this.Old.Get_Id());
};
CChangesFootnotesSetSeparator.prototype.ReadFromBinary = function(Reader)
{
	// Long : Flags
	// 1bit : is new undefined?
	// 2bit : is old undefined?
	// String : New id (1 bit zero)
	// String : Old id (2 bit zero)

	var nFlags = Reader.GetLong();

	if (nFlags & 1)
		this.New = null;
	else
		this.New = AscCommon.g_oTableId.Get_ById(Reader.GetString2());

	if (nFlags & 2)
		this.New = null;
	else
		this.New = AscCommon.g_oTableId.Get_ById(Reader.GetString2());
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesFootnotesSetContinuationSeparator(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesFootnotesSetContinuationSeparator.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesFootnotesSetContinuationSeparator.prototype.constructor = CChangesFootnotesSetContinuationSeparator;
CChangesFootnotesSetContinuationSeparator.prototype.Type = AscDFH.historyitem_Footnotes_SetContinuationSeparator;
CChangesFootnotesSetContinuationSeparator.prototype.private_SetValue = function(Value)
{
	this.Class.ContinuationSeparator = Value;
};
CChangesFootnotesSetContinuationSeparator.prototype.WriteToBinary = CChangesFootnotesSetSeparator.prototype.WriteToBinary;
CChangesFootnotesSetContinuationSeparator.prototype.ReadFromBinary = CChangesFootnotesSetSeparator.prototype.ReadFromBinary;
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseProperty}
 */
function CChangesFootnotesSetContinuationNotice(Class, Old, New, Color)
{
	AscDFH.CChangesBaseProperty.call(this, Class, Old, New, Color);
}
CChangesFootnotesSetContinuationNotice.prototype = Object.create(AscDFH.CChangesBaseProperty.prototype);
CChangesFootnotesSetContinuationNotice.prototype.constructor = CChangesFootnotesSetContinuationNotice;
CChangesFootnotesSetContinuationNotice.prototype.Type = AscDFH.historyitem_Footnotes_SetContinuationNotice;
CChangesFootnotesSetContinuationNotice.prototype.private_SetValue = function(Value)
{
	this.Class.ContinuationNotice = Value;
};
CChangesFootnotesSetContinuationNotice.prototype.WriteToBinary = CChangesFootnotesSetSeparator.prototype.WriteToBinary;
CChangesFootnotesSetContinuationNotice.prototype.ReadFromBinary = CChangesFootnotesSetSeparator.prototype.ReadFromBinary;
/**
 * @constructor
 * @extends {AscDFH.CChangesBase}
 */
function CChangesFootnotesRemoveFootnote(Class, Id)
{
	AscDFH.CChangesBase.call(this, Class);

	this.Id = Id;
}
CChangesFootnotesRemoveFootnote.prototype = Object.create(AscDFH.CChangesBase.prototype);
CChangesFootnotesRemoveFootnote.prototype.constructor = CChangesFootnotesRemoveFootnote;
CChangesFootnotesRemoveFootnote.prototype.Type = AscDFH.historyitem_Footnotes_RemoveFootnote;
CChangesFootnotesRemoveFootnote.prototype.Undo = function()
{
	this.Class.Footnote[this.Id] = AscCommon.g_oTableId.Get_ById(this.Id);
};
CChangesFootnotesRemoveFootnote.prototype.Redo = function()
{
	delete this.Class.Footnote[this.Id];
};
CChangesFootnotesRemoveFootnote.prototype.WriteToBinary = function(Writer)
{
	// String : Id
	Writer.WriteString2(this.Id);
};
CChangesFootnotesRemoveFootnote.prototype.ReadFromBinary = function(Reader)
{
	// String : Id
	this.Id = Reader.GetString2();
};
CChangesFootnotesRemoveFootnote.prototype.CreateReverseChange = function()
{
	return new CChangesFootnotesAddFootnote(this.Class, this.Id);
};
CChangesFootnotesRemoveFootnote.prototype.Merge = function(oChange)
{
	if (this.Class !== oChange.Class)
		return true;

	if (this.Type === oChange.Type && this.Id === oChange.Id)
		return false;

	return true;
};
