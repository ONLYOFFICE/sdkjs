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
		this.SectionsInfo  = oLogicDocument.GetSections();
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
	 * @returns {{Header : null, SectPr : AscWord.SectPr, Footer : null}}
	 */
	CDocumentLayoutBase.prototype.GetSectionHdrFtr = function(nPageAbs, isFirst, isEven)
	{
		let oLogicDocument = this.LogicDocument;
		
		let docPage      = oLogicDocument.GetPage(nPageAbs);
		let sectionIndex = docPage.GetSection(0).GetIndex();
		let sectPr       = this.SectionsInfo.GetSectPrByIndex(sectionIndex);
		let startSectPr  = sectPr;
		
		isEven  = isEven && sectPr.IsEvenAndOdd();
		isFirst = isFirst && sectPr.IsTitlePage();
		
		let oHeader = null;
		let oFooter = null;
		while (sectionIndex >= 0)
		{
			sectPr = this.SectionsInfo.GetSectPrByIndex(sectionIndex--);
			
			if (!oHeader)
				oHeader = sectPr.GetHdrFtr(true, isFirst, isEven);
			
			if (!oFooter)
				oFooter = sectPr.GetHdrFtr(false, isFirst, isEven);
			
			if (oHeader && oFooter)
				break;
		}
		
		return {
			Header : oHeader,
			Footer : oFooter,
			SectPr : startSectPr
		};
	};
	/**
	 * Получаем границы, внутри короторых должно быть расчитано содержимое основной части документа
	 * @param nPageAbs {number}
	 * @param oSectPr {AscWord.SectPr}
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
	 * @param oSectPr {AscWord.SectPr}
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
	 * Получаем настройки секции на заданной странице
	 * @param pageAbs {number}
	 * @returns {AscWord.SectPr}
	 */
	CDocumentLayoutBase.prototype.GetPageStartSection = function(pageAbs)
	{
		let docPage = this.LogicDocument.GetPage(pageAbs);
		let pageSection = docPage ? docPage.GetSection(0) : null;
		let sectPr = pageSection ? pageSection.GetSectPr() : this.LogicDocument.GetFinalSectPr();
		return this.CheckSectPr(sectPr);
	};
	CDocumentLayoutBase.prototype.GetSectionByElement = function(element)
	{
		return this.SectionsInfo.GetSectPrByElement(element);
	};
	CDocumentLayoutBase.prototype.CheckSectPr = function(sectPr)
	{
		return sectPr;
	};
	CDocumentLayoutBase.prototype.GetFinalSectPr = function()
	{
		return this.LogicDocument.SectPr;
	};
	/**
	 * Получаем номер секции в общем списке секций
	 * @param oSectPr {AscWord.SectPr}
	 * @returns {number}
	 */
	CDocumentLayoutBase.prototype.GetSectionIndex = function(oSectPr)
	{
		return this.SectionsInfo.Find(oSectPr);
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
	/**
	 * @returns {number}
	 */
	CDocumentLayoutBase.prototype.calculateIndent = function(ind, sectPr)
	{
		return ind;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.CDocumentLayoutBase = CDocumentLayoutBase;
	AscWord.DEFAULT_PAGE_HEIGHT = DEFAULT_PAGE_HEIGHT;
	AscWord.DEFAULT_PAGE_WIDTH  = DEFAULT_PAGE_WIDTH;

})(window);
