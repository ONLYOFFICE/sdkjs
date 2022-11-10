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

(function (window)
{
	/**
	 * Класс работающий с общей историей в совместном редактировании
	 * @param {AscCommon.CCollaborativeEditingBase} coEditing
	 * @constructor
	 */
	function CCollaborativeHistory(coEditing)
	{
		this.CoEditing = coEditing;

		this.Changes   = []; // Список всех изменений
		this.OwnRanges = []; // Диапазоны собственных изменений

		this.SyncIndex = -1; // Позиция в массиве изменений, которые согласованы с сервером

	}

	CCollaborativeHistory.prototype.AddChange = function(change)
	{
		this.Changes.push(change);
	};
	CCollaborativeHistory.prototype.AddOwnChanges = function(ownChanges, deleteIndex)
	{
		if (null !== deleteIndex)
			this.Changes.length = this.SyncIndex + deleteIndex;
		else
			this.SyncIndex = this.Changes.length;

		// TODO: Пока мы делаем это как одну точку, которую надо откатить. Надо пробежаться по массиву и разбить его
		//       по отдельным действиям. В принципе, данная схема срабатывает в быстром совместном редактировании,
		//       так что как правило две точки не успевают попасть в одно сохранение.
		if (ownChanges.length > 0)
		{
			this.OwnChangesRanges.push(new COwnRange(this.Changes.length, ownChanges.length));
			this.Changes = this.Changes.concat(ownChanges);
		}
	};
	/**
	 * Откатываем заданное количество действий
	 * @param {number} count
	 * @returns {[]} возвращаем массив откаченных действий
	 */
	CCollaborativeHistory.prototype.UndoGlobalChanges = function(count)
	{
		count = Math.min(count, this.Changes.length);

		let index = this.Changes.length - 1;
		let changeArray = [];
		while (index >= this.Changes.length - count)
		{
			let change = this.Changes[index--];
			if (!change)
				continue;

			if (change.IsContentChange())
			{
				let simpleChanges = change.ConvertToSimpleChanges();
				for (let simpleIndex = simpleChanges.length - 1; simpleIndex >= 0; --simpleIndex)
				{
					simpleChanges[simpleIndex].Undo();
					changeArray.push(simpleChanges[simpleIndex]);
				}
			}
			else
			{
				change.Undo();
				changeArray.push(change);
			}
		}

		this.Changes.length = this.Changes.length - count;
		return changeArray;
	};
	/**
	 * Отменяем все действия, попавшие в последнюю точку истории
	 * @returns {[]} возвращаем массив отмененных действий
	 */
	CCollaborativeHistory.prototype.UndoGlobalPoint = function()
	{
		let count = 0;
		for (let index = this.Changes.length - 1; index > 0; --index, ++count)
		{
			let change = this.Changes[index];
			if (!change)
				continue;

			if (change.IsDescriptionChange())
			{
				count++;
				break;
			}
		}

		return count ? this.UndoGlobalChanges(count) : [];
	};

	CCollaborativeHistory.prototype.GetEmptyContentChanges = function()
	{
		let changes = [];
		for (let index = this.Changes.length - 1; index >= 0; --index)
		{
			let tempChange = this.Changes[index];

			if (tempChange.IsContentChange() && tempChange.GetItemsCount() <= 0)
				changes.push(tempChange);
		}
		return changes;
	};

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//------------------------------------------------------------------------------------------------------------------
	function COwnRange(position, length)
	{
		this.Position = position;
		this.Length   = length;
	}

	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].CCollaborativeHistory = CCollaborativeHistory;

})(window);
