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
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesFormFieldAddItem(Class, Pos, Items)
{
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, true);
}
CChangesFormFieldAddItem.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesFormFieldAddItem.prototype.constructor = CChangesFormFieldAddItem;
CChangesFormFieldAddItem.prototype.Type = AscDFH.historyitem_FieldContent_AddItem;
CChangesFormFieldAddItem.prototype.Undo = function()
{
	this.Class.Content.splice(this.Pos, this.Items.length);
};
CChangesFormFieldAddItem.prototype.Redo = function()
{
	let Array_start = this.Class.Content.slice(0, this.Pos);
	let Array_end   = this.Class.Content.slice(this.Pos);

	this.Class.Content = Array_start.concat(this.Items, Array_end);
};
CChangesFormFieldAddItem.prototype.private_WriteItem = function(Writer, Item)
{
	Writer.WriteString2(Item.Get_Id());
};
CChangesFormFieldAddItem.prototype.private_ReadItem = function(Reader)
{
	return AscCommon.g_oTableId.Get_ById(Reader.GetString2());
};
CChangesFormFieldAddItem.prototype.Load = function(Color)
{
	var oFieldContent = this.Class;
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		var Pos     = oFieldContent.m_oContentChanges.Check(AscCommon.contentchanges_Add, this.PosArray[nIndex]);
		var Element = this.Items[nIndex];

		if (null != Element)
		{
			oFieldContent.Content.splice(Pos, 0, Element);
			AscCommon.CollaborativeEditing.Update_DocumentPositionsOnAdd(oFieldContent, Pos);
		}
	}
};
CChangesFormFieldAddItem.prototype.IsRelated = function(oChanges)
{
	if (this.Class === oChanges.Class && (AscDFH.historyitem_FieldContent_AddItem === oChanges.Type || AscDFH.historyitem_FieldContent_RemoveItem === oChanges.Type))
		return true;

	return false;
};
CChangesFormFieldAddItem.prototype.CreateReverseChange = function()
{
	return this.private_CreateReverseChange(CChangesFormFieldRemoveItem);
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesFormFieldRemoveItem(Class, Pos, Items)
{
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, false);
}
CChangesFormFieldRemoveItem.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesFormFieldRemoveItem.prototype.constructor = CChangesFormFieldRemoveItem;
CChangesFormFieldRemoveItem.prototype.Type = AscDFH.historyitem_FieldContent_RemoveItem;
CChangesFormFieldRemoveItem.prototype.Undo = function()
{
	let Array_start = this.Class.Content.slice(0, this.Pos);
	let Array_end   = this.Class.Content.slice(this.Pos);

	this.Class.Content = Array_start.concat(this.Items, Array_end);
};
CChangesFormFieldRemoveItem.prototype.Redo = function()
{
	this.Class.Content.splice(this.Pos, this.Items.length);
};
CChangesFormFieldRemoveItem.prototype.private_WriteItem = function(Writer, Item)
{
	Writer.WriteString2(Item.Get_Id());
};
CChangesFormFieldRemoveItem.prototype.private_ReadItem = function(Reader)
{
	return AscCommon.g_oTableId.Get_ById(Reader.GetString2());
};
CChangesFormFieldRemoveItem.prototype.Load = function(Color)
{
	var oFieldContent = this.Class;
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		var ChangesPos = oFieldContent.m_oContentChanges.Check(AscCommon.contentchanges_Remove, this.PosArray[nIndex]);

		if (false === ChangesPos)
			continue;

		oFieldContent.Content.splice(ChangesPos, 1);
		AscCommon.CollaborativeEditing.Update_DocumentPositionsOnRemove(oFieldContent, ChangesPos, 1);
	}
};
CChangesFormFieldRemoveItem.prototype.IsRelated = function(oChanges)
{
	if (this.Class === oChanges.Class && (AscDFH.historyitem_FieldContent_AddItem === oChanges.Type || AscDFH.historyitem_FieldContent_RemoveItem === oChanges.Type))
		return true;

	return false;
};
CChangesFormFieldRemoveItem.prototype.CreateReverseChange = function()
{
	return this.private_CreateReverseChange(CChangesFormFieldAddItem);
};

AscDFH.changesFactory[AscDFH.historyitem_FieldContent_AddItem]    = CChangesFormFieldAddItem;
AscDFH.changesFactory[AscDFH.historyitem_FieldContent_RemoveItem] = CChangesFormFieldRemoveItem;
