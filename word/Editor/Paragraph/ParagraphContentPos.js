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
