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
