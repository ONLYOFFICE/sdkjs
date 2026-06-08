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

(function(window)
{
	/**
	 * Класс, для работы с позицией внутри содержимого параграфа
	 * @constructor
	 */
	function CParagraphContentPos()
	{
		this.Data  = [0, 0, 0];
		this.Depth = 0;
		this.bPlaceholder = false;
	}

	CParagraphContentPos.prototype.Add = function(nPos)
	{
		this.Data[this.Depth] = nPos;
		this.Depth++;
	};
	CParagraphContentPos.prototype.Update = function(nPos, nDepth)
	{
		this.Data[nDepth] = nPos;
		this.Depth        = nDepth + 1;
	};
	CParagraphContentPos.prototype.Update2 = function(nPos, nDepth)
	{
		this.Data[nDepth] = nPos;
	};
	/**
	 * Коипруем данные позиции
	 */
	CParagraphContentPos.prototype.Set = function(oPos)
	{
		let nDepth = oPos.Depth;

		for (let nPos = 0; nPos < nDepth; ++nPos)
			this.Data[nPos] = oPos.Data[nPos];

		this.Depth = nDepth;

		if (this.Data.length > this.Depth)
			this.Data.length = this.Depth;
	};
	CParagraphContentPos.prototype.Get = function(nDepth)
	{
		return this.Data[nDepth];
	};
	CParagraphContentPos.prototype.Reset = function()
	{
		this.Data  = [0, 0, 0];
		this.Depth = 0;
	};
	/**
	 * Получаем текущую глубину позиции
	 * @returns {number}
	 */
	CParagraphContentPos.prototype.GetDepth = function()
	{
		return this.Depth - 1;
	};
	/**
	 * В данной функции мы устанавливаем глубину позиции (при этом не меняя сам массив позиции)
	 * @param {number} nDepth
	 */
	CParagraphContentPos.prototype.SetDepth = function(nDepth)
	{
		this.Depth = Math.max(0, Math.min(nDepth + 1, this.Data.length - 1));
	};
	/**
	 * Уменьшаем глубину на заданное значение
	 * @param {number} nCount
	 */
	CParagraphContentPos.prototype.DecreaseDepth = function(nCount)
	{
		this.Depth = Math.max(0, this.Depth - nCount);
	};
	CParagraphContentPos.prototype.Copy = function()
	{
		let oPos = new CParagraphContentPos();
		oPos.Set(this);
		return oPos;
	};
	/**
	 * Сравниваем текущую позицию с заданной
	 * @param {CParagraphContentPos} oPos
	 * @returns {number} 0 - позиции совпадают, 1 - текущая позиция дальше заданной, -1 - текущая позиция до заданной
	 */
	CParagraphContentPos.prototype.Compare = function(oPos)
	{
		if (!this.IsValid() || !oPos || !oPos.IsValid())
			return -2;

		let nDepth = 0;

		let nLen1   = this.Depth;
		let nLen2   = oPos.Depth;
		let nLenMin = Math.min(nLen1, nLen2);

		while (nDepth < nLenMin)
		{
			if (this.Data[nDepth] > oPos.Data[nDepth])
				return 1;
			else if (this.Data[nDepth] < oPos.Data[nDepth])
				return -1;

			++nDepth;
		}

		if (nLen1 > nLen2)
			return 1;
		else if (nLen1 < nLen2)
			return -1;

		return 0;
	};
	/**
	 * Получаем позицию NearestPos
	 * @param oParagraph
	 */
	CParagraphContentPos.prototype.ToAnchorPos = function(oParagraph)
	{
		if (!oParagraph)
			return;

		var oNearPos = {
			Paragraph  : oParagraph,
			ContentPos : this,
			transform  : oParagraph.Get_ParentTextTransform()
		};

		oParagraph.Check_NearestPos(oNearPos);
		return oNearPos;
	};
	/**
	 * Проверяем позиции на совпадение
	 * @param {CParagraphContentPos} oPos
	 * @returns {boolean}
	 */
	CParagraphContentPos.prototype.IsEqual = function(oPos)
	{
		return (oPos && 0 === this.Compare(oPos));
	};
	/**
	 * Проверяем, что текущая позиция является частью заданной
	 * @param oPos {CParagraphContentPos}
	 * @returns {boolean}
	 */
	CParagraphContentPos.prototype.IsPartOf = function(oPos)
	{
		if (!oPos
			|| this.IsEmpty()
			|| oPos.IsEmpty()
			|| !this.IsValid()
			|| !oPos.IsValid()
			|| this.Depth > oPos.Depth)
			return false;

		let nDepth = 0;
		let nLen   = this.Depth;
		while (nDepth < nLen)
		{
			if (this.Data[nDepth] !== oPos.Data[nDepth])
				return false;

			++nDepth;
		}

		return true;
	};
	CParagraphContentPos.prototype.IsValid = function()
	{
		return (this.Depth <= this.Data.length);
	};
	CParagraphContentPos.prototype.IsEmpty = function()
	{
		return (0 === this.Depth);
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].CParagraphContentPos = CParagraphContentPos;

})(window);
