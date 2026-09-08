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

AscDFH.changesFactory[AscDFH.historyitem_DocumentContent_AddItem]    = CChangesDocumentContentAddItem;
AscDFH.changesFactory[AscDFH.historyitem_DocumentContent_RemoveItem] = CChangesDocumentContentRemoveItem;

//----------------------------------------------------------------------------------------------------------------------
// Карта зависимости изменений
//----------------------------------------------------------------------------------------------------------------------
AscDFH.changesRelationMap[AscDFH.historyitem_DocumentContent_AddItem]    = [
	AscDFH.historyitem_DocumentContent_AddItem,
	AscDFH.historyitem_DocumentContent_RemoveItem
];
AscDFH.changesRelationMap[AscDFH.historyitem_DocumentContent_RemoveItem] = [
	AscDFH.historyitem_DocumentContent_AddItem,
	AscDFH.historyitem_DocumentContent_RemoveItem
];
//----------------------------------------------------------------------------------------------------------------------

/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesDocumentContentAddItem(Class, Pos, Items)
{
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, true);
}
CChangesDocumentContentAddItem.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesDocumentContentAddItem.prototype.constructor = CChangesDocumentContentAddItem;
CChangesDocumentContentAddItem.prototype.Type = AscDFH.historyitem_DocumentContent_AddItem;
CChangesDocumentContentAddItem.prototype.Undo = function()
{
	var oDocument = this.Class;
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let pos = true !== this.UseArray ? this.Pos : this.PosArray[nIndex];
		const removedElement = oDocument.Content[pos];
		oDocument.UpdateSectionsBeforeRemove([removedElement], false);
		removedElement.OnDetach();
		removedElement.UpdatePresentationOutline();

		var elements = oDocument.Content.splice(pos, 1);
		oDocument.private_RecalculateNumbering(elements);
		oDocument.private_ReindexContent(pos);
		oDocument.private_UpdateSelectionPosOnRemove(pos, 1);

		let prev = pos > 0 ? oDocument.Content[pos - 1] : null;
		let next = pos < oDocument.Content.length ? oDocument.Content[pos] : null;

		if (prev)
			prev.Set_DocumentNext(next);

		if (next)
			next.Set_DocumentPrev(prev);

		removedElement.Set_DocumentNext(null);
		removedElement.Set_DocumentPrev(null);
		removedElement.SetParent(null);
	}
	this.Class.Recalculated = false;
};
CChangesDocumentContentAddItem.prototype.Redo = function()
{
	var oDocument = this.Class;
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		var element = this.Items[nIndex];
		let pos     = true !== this.UseArray ? this.Pos + nIndex : this.PosArray[nIndex];

		oDocument.Content.splice(pos, 0, element);
		oDocument.private_RecalculateNumbering([element]);
		oDocument.private_ReindexContent(pos);
		oDocument.private_UpdateSelectionPosOnAdd(pos, 1);

		let prev = pos > 0 ? oDocument.Content[pos - 1] : null;
		let next = pos + 1 < oDocument.Content.length ? oDocument.Content[pos + 1] : null;

		if (prev)
			prev.Set_DocumentNext(element);
		if (next)
			next.Set_DocumentPrev(element);

		element.Set_DocumentPrev(prev);
		element.Set_DocumentNext(next);
		element.SetParent(oDocument);

		oDocument.UpdateSectionsAfterAdd([element]);
		element.OnAttach();
		element.UpdatePresentationOutline();
	}
	this.Class.Recalculated = false;
};
CChangesDocumentContentAddItem.prototype.private_WriteItem = function(Writer, Item)
{
	Writer.WriteString2(Item.Get_Id());
};
CChangesDocumentContentAddItem.prototype.private_ReadItem = function(Reader)
{
	return AscCommon.g_oTableId.Get_ById(Reader.GetString2());
};
CChangesDocumentContentAddItem.prototype.Load = function(Color)
{
	if (this.PosArray.length <= 0 || this.Items.length <= 0)
		return;

	var oDocument = this.Class;
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let pos     = oDocument.m_oContentChanges.Check(AscCommon.contentchanges_Add, true !== this.UseArray ? this.Pos + nIndex : this.PosArray[nIndex]);
		var element = this.Items[nIndex];

		pos = Math.min(pos, oDocument.Content.length);

		if (!element)
			continue;

		oDocument.Content.splice(pos, 0, element);
		oDocument.private_RecalculateNumbering([element]);
		oDocument.private_ReindexContent(pos);

		let prev = pos > 0 ? oDocument.Content[pos - 1] : null;
		let next = pos + 1 < oDocument.Content.length ? oDocument.Content[pos + 1] : null;

		if (prev)
			prev.Set_DocumentNext(element);
		if (next)
			next.Set_DocumentPrev(element);

		element.Set_DocumentPrev(prev);
		element.Set_DocumentNext(next);
		element.SetParent(oDocument);

		AscCommon.CollaborativeEditing.Update_DocumentPositionsOnAdd(oDocument, pos);
		oDocument.UpdateSectionsAfterAdd([element]);
		element.OnAttach();
		element.UpdatePresentationOutline();
	}

	this.Class.Recalculated = false;
};
CChangesDocumentContentAddItem.prototype.IsRelated = function(oChanges)
{
	if (this.Class === oChanges.Class && (AscDFH.historyitem_DocumentContent_AddItem === oChanges.Type || AscDFH.historyitem_DocumentContent_RemoveItem === oChanges.Type))
		return true;

	return false;
};
CChangesDocumentContentAddItem.prototype.CreateReverseChange = function()
{
	return this.private_CreateReverseChange(CChangesDocumentContentRemoveItem);
};
/**
 * @constructor
 * @extends {AscDFH.CChangesBaseContentChange}
 */
function CChangesDocumentContentRemoveItem(Class, Pos, Items)
{
	AscDFH.CChangesBaseContentChange.call(this, Class, Pos, Items, false);
}
CChangesDocumentContentRemoveItem.prototype = Object.create(AscDFH.CChangesBaseContentChange.prototype);
CChangesDocumentContentRemoveItem.prototype.constructor = CChangesDocumentContentRemoveItem;
CChangesDocumentContentRemoveItem.prototype.Type = AscDFH.historyitem_DocumentContent_RemoveItem;
CChangesDocumentContentRemoveItem.prototype.Undo = function()
{
	if (!this.Items || this.Items.length <= 0)
		return;

	var oDocument = this.Class;

	var arrayStart = oDocument.Content.slice(0, this.Pos);
	var arrayEnd   = oDocument.Content.slice(this.Pos);

	oDocument.private_RecalculateNumbering(this.Items);
	oDocument.private_ReindexContent(this.Pos);
	oDocument.Content = arrayStart.concat(this.Items, arrayEnd);

	oDocument.private_UpdateSelectionPosOnAdd(this.Pos, this.Items.length);

	let startIndex = Math.max(this.Pos - 1, 0);
	let endIndex   = Math.min(oDocument.Content.length - 1, this.Pos + this.Items.length + 1);
	for (let index = startIndex; index <= endIndex; ++index)
	{
		let element = oDocument.Content[index];
		element.Set_DocumentPrev(index > 0 ? oDocument.Content[index - 1] : null);
		element.Set_DocumentNext(index < oDocument.Content.length - 1 ? oDocument.Content[index + 1] : null);
		element.SetParent(oDocument);
	}
	this.Class.Recalculated = false;
	oDocument.UpdateSectionsAfterAdd(this.Items);
	for (let i = 0; i < this.Items.length; ++i)
	{
		const element = this.Items[i];
		element.OnAttach();
		element.UpdatePresentationOutline();
	}

};
CChangesDocumentContentRemoveItem.prototype.Redo = function()
{
	if (!this.Items || this.Items.length <= 0)
		return;

	var oDocument = this.Class;

	let removedElements = oDocument.Content.slice(this.Pos, this.Pos + this.Items.length);
	oDocument.UpdateSectionsBeforeRemove(removedElements, false);

	for (let i = 0; i < removedElements.length; ++i) {
		const removedElement = removedElements[i];
		removedElement.OnDetach();
		removedElement.UpdatePresentationOutline();
	}

	let elements = oDocument.Content.splice(this.Pos, this.Items.length);
	oDocument.private_RecalculateNumbering(elements);
	oDocument.private_ReindexContent(this.Pos);
	oDocument.private_UpdateSelectionPosOnRemove(this.Pos, this.Items.length);

	let pos  = this.Pos;
	let prev = pos > 0 ? oDocument.Content[pos - 1] : null;
	let next = pos < oDocument.Content.length ? oDocument.Content[pos] : null;

	if (prev)
		prev.Set_DocumentNext(next);

	if (next)
		next.Set_DocumentPrev(prev);

	for (let i = 0; i < elements.length; ++i)
	{
		elements[i].Set_DocumentNext(null);
		elements[i].Set_DocumentPrev(null);
		elements[i].SetParent(null);
	}
	this.Class.Recalculated = false;
};
CChangesDocumentContentRemoveItem.prototype.private_WriteItem = function(Writer, Item)
{
	Writer.WriteString2(Item.Get_Id());
};
CChangesDocumentContentRemoveItem.prototype.private_ReadItem = function(Reader)
{
	return AscCommon.g_oTableId.Get_ById(Reader.GetString2());
};
CChangesDocumentContentRemoveItem.prototype.Load = function(Color)
{
	var oDocument = this.Class;
	for (var nIndex = 0, nCount = this.Items.length; nIndex < nCount; ++nIndex)
	{
		let pos = oDocument.m_oContentChanges.Check(AscCommon.contentchanges_Remove, true !== this.UseArray ? this.Pos : this.PosArray[nIndex]);

		// действие совпало, не делаем его
		if (false === pos)
			continue;

		let removedElement = oDocument.Content[pos];
		oDocument.UpdateSectionsBeforeRemove([removedElement], false);
		removedElement.OnDetach();
		removedElement.UpdatePresentationOutline();

		let elements = oDocument.Content.splice(pos, 1);
		oDocument.private_RecalculateNumbering(elements);
		AscCommon.CollaborativeEditing.Update_DocumentPositionsOnRemove(oDocument, pos, 1);

		let prev = pos > 0 ? oDocument.Content[pos - 1] : null;
		let next = pos < oDocument.Content.length ? oDocument.Content[pos] : null;

		if (prev)
			prev.Set_DocumentNext(next);

		if (next)
			next.Set_DocumentPrev(prev);

		removedElement.Set_DocumentNext(null);
		removedElement.Set_DocumentPrev(null);
		removedElement.SetParent(null);

		oDocument.private_ReindexContent(pos);
	}
	this.Class.Recalculated = false;
};
CChangesDocumentContentRemoveItem.prototype.IsRelated = function(oChanges)
{
	if (this.Class === oChanges.Class && (AscDFH.historyitem_DocumentContent_AddItem === oChanges.Type || AscDFH.historyitem_DocumentContent_RemoveItem === oChanges.Type))
		return true;

	return false;
};
CChangesDocumentContentRemoveItem.prototype.CreateReverseChange = function()
{
	return this.private_CreateReverseChange(CChangesDocumentContentAddItem);
};
