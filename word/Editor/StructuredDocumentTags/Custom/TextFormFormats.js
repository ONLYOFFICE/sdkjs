/*
 * (c) Copyright Ascensio System SIA 2010-2022
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

(function(window)
{
	/**
	 * Базовый класс для всех форматов
	 * @constructor
	 */
	function CTextFormFormatBase()
	{
		this.Set = []; // Специальный параметр для возможного ограничения на ввод символов
	}
	CTextFormFormatBase.prototype.Check = function(sText)
	{
		if (sText && this.Set.length)
		{
			for (let oIter = sText.getUnicodeIterator(); iter.check(); iter.next())
			{
				if (-1 === this.Set.indexOf(oIter.value()))
					return false;
			}
		}

		return true;
	};
	/**
	 * Класс представляющий пустой формат
	 * @constructor
	 * @extends {CTextFormFormatBase}
	 */
	function CTextFormFormatNone()
	{
		CTextFormFormatBase.call(this);
	}
	CTextFormFormatNone.prototype = Object.create(AscWord.CRunElementBase.prototype);
	CTextFormFormatNone.prototype.constructor = CTextFormFormatBase;
	CTextFormFormatNone.prototype.Check = function(sText)
	{
		return this.constructor.prototype.Check.call(sText);
	};

	/**
	 * Только числа
	 * @constructor
	 * @extends {CTextFormFormatBase}
	 */
	function CTextFormFormatDecimal()
	{
		CTextFormFormatBase.call(this);
	}
	CTextFormFormatDecimal.prototype = Object.create(AscWord.CRunElementBase.prototype);
	CTextFormFormatDecimal.prototype.constructor = CTextFormFormatBase;
	CTextFormFormatDecimal.prototype.Check = function(sText)
	{
		return true;
	};
	/**
	 * Формат - маска
	 * @param sMask {string}
	 * @constructor
	 * @extends {CTextFormFormatBase}
	 */
	function CTextFormFormatLetters(sMask)
	{
		CTextFormFormatBase.call(this);
	}
	CTextFormFormatMask.prototype = Object.create(AscWord.CRunElementBase.prototype);
	CTextFormFormatMask.prototype.constructor = CTextFormFormatBase;
	CTextFormFormatMask.prototype.Check = function(sText)
	{
		return true;
	};
	/**
	 * Формат - маска
	 * @param sMask {string}
	 * @constructor
	 * @extends {CTextFormFormatBase}
	 */
	function CTextFormFormatMask(sMask)
	{
		CTextFormFormatBase.call(this);
	}
	CTextFormFormatMask.prototype = Object.create(AscWord.CRunElementBase.prototype);
	CTextFormFormatMask.prototype.constructor = CTextFormFormatBase;
	CTextFormFormatMask.prototype.Check = function(sText)
	{
		return true;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].CTextFormFormatNone    = CTextFormFormatNone;
	window['AscWord'].CTextFormFormatDecimal = CTextFormFormatDecimal;
	window['AscWord'].CTextFormFormatMask    = CTextFormFormatMask;



})(window);
