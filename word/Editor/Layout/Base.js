/*
 * (c) Copyright Ascensio System SIA 2010-2023
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
	const DEFAULT_PAGE_WIDTH  = 210;
	const DEFAULT_PAGE_HEIGHT = 297;

	/**
	 * Базовый класс для вида документа
	 * @param oLogicDocument {CDocument}
	 * @constructor
	 */
	function CDocumentLayoutBase(oLogicDocument)
	{
		this.LogicDocument = oLogicDocument;
		this.SectInfo      = new CDocumentSectionsInfoElement(oLogicDocument.GetLastSection(), 0);
	}
	CDocumentLayoutBase.prototype.IsPrintMode = function()
	{
		return false;
	};
	CDocumentLayoutBase.prototype.IsReadMode = function()
	{
		return false;
	};
	/**
	 * Рисуем заголовки
	 * @returns {boolean}
	 */
	CDocumentLayoutBase.prototype.IsHeaders = function()
	{
		return true;
	};
	/**
	 * Дополнительный коэффициент для шрифта
	 * returns {number}
	 */
	CDocumentLayoutBase.prototype.GetFontScale = function()
	{
		return 1;
	};
	/**
	 * Получаем колонтитулы и секцию на заданной странице
	 * @param nPageAbs
	 * @param isFirst
	 * @param isEven
	 * @returns {{Header : null, SectPr : AscWord.CSectionPr, Footer : null}}
	 */
	CDocumentLayoutBase.prototype.GetSectionHdrFtr = function(nPageAbs, isFirst, isEven)
	{
		return {
			Header : null,
			Footer : null,
			SectPr : this.LogicDocument.GetLastSection()
		}
	};
	/**
	 * Получаем границы, внутри короторых должно быть расчитано содержимое основной части документа
	 * @param nPageAbs {number}
	 * @param oSectPr {AscWord.CSectionPr}
	 * @returns {{X : number, Y : number, XLimit : number, YLimit : number}}
	 */
	CDocumentLayoutBase.prototype.GetPageContentFrame = function(nPageAbs, oSectPr)
	{
		return {
			X      : 0,
			Y      : 0,
			XLimit : DEFAULT_PAGE_WIDTH,
			YLimit : DEFAULT_PAGE_HEIGHT
		};
	};
	/**
	 * Получаем границы содержимого по заданной колонке и заданной странице
	 * @param nPageAbs {number}
	 * @param nColumnAbs {number}
	 * @param oSectPr {AscWord.CSectionPr}
	 * @returns {{ColumnSpaceBefore : number, X : number, ColumnSpaceAfter : number, Y : number, XLimit : number, YLimit : number}}
	 */
	CDocumentLayoutBase.prototype.GetColumnContentFrame = function(nPageAbs, nColumnAbs, oSectPr)
	{
		return {
			X                 : 0,
			Y                 : 0,
			XLimit            : DEFAULT_PAGE_WIDTH,
			YLimit            : DEFAULT_PAGE_HEIGHT,
			ColumnSpaceBefore : 0,
			ColumnSpaceAfter  : 0
		};
	};
	/**
	 * Получаем настройки секции на заданной странице, или заданного элемента
	 * @param nPageAbs {number} Если номер элемента не задан, тогда получаем по заданному номеру страницы
	 * @param nContentIndex {?number} Если задан номер элемента, то ориентируемся на него
	 * @returns {AscWord.CSectionPr}
	 */
	CDocumentLayoutBase.prototype.GetSection = function(nPageAbs, nContentIndex)
	{
		return this.LogicDocument.GetLastSection();
	};
	CDocumentLayoutBase.prototype.GetSectionByPos = function(nContentIndex)
	{
		return this.LogicDocument.GetLastSection();
	};
	CDocumentLayoutBase.prototype.GetSectionInfo = function(nContentIndex)
	{
		return this.SectInfo;
	};
	/**
	 * Получаем номер секции в общем списке секций
	 * @param oSectPr {AscWord.CSectionPr}
	 * @returns {number}
	 */
	CDocumentLayoutBase.prototype.GetSectionIndex = function(oSectPr)
	{
		return 0;
	};
	/**
	 * Получаем время (в миллисекундах) доступное для однократного синхронного пересчета страниц
	 * @returns {number}
	 */
	CDocumentLayoutBase.prototype.GetCalculateTimeLimit = function()
	{
		return 10;
	};
	/**
	 * Получаем коэффициент изменения заданной секции
	 * @param oSectPr
	 * @returns {number}
	 */
	CDocumentLayoutBase.prototype.GetScaleBySection = function(oSectPr)
	{
		return 1;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.CDocumentLayoutBase = CDocumentLayoutBase;
	AscWord.DEFAULT_PAGE_HEIGHT = DEFAULT_PAGE_HEIGHT;
	AscWord.DEFAULT_PAGE_WIDTH  = DEFAULT_PAGE_WIDTH;

})(window);
