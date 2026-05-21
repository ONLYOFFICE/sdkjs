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
	 * @param {AscWord.Paragraph.WrapRecalcState} wrapState
	 * @constructor
	 */
	function AlignRecalcState(wrapState)
	{
		this.wrapState     = wrapState;
	    this.X             = 0; // Текущая позиция по горизонтали
	    this.Y             = 0; // Текущая позиция по вертикали
	    this.XEnd          = 0; // Предельная позиция по горизонтали
	    this.JustifyWord   = 0; // Добавочная ширина символов
	    this.JustifySpace  = 0; // Добавочная ширина пробелов
	    this.SpacesCounter = 0; // Счетчик пробелов с добавочной шириной (чтобы пробелы в конце строки не трогать)
	    this.SpacesSkip    = 0; // Количество пробелов, которые мы пропускаем в начале строки
	    this.LettersSkip   = 0; // Количество букв, которые мы пропускаем (из-за таба)
	    this.LastW         = 0; // Ширина последнего элемента (необходимо для позиционирования картинки)
	    this.Paragraph     = undefined;
	    this.RecalcResult  = 0x00;//recalcresult_NextElement;

	    this.Y0            = 0; // Верхняя граница строки
	    this.Y1            = 0; // Нижняя граница строки

	    this.CurPage       = 0;
	    this.PageY         = 0;
	    this.PageX         = 0;

	    this.RecalcFast    = false; // Если пересчет быстрый, тогда все "плавающие" объекты мы не трогаем
	    this.RecalcFast2   = false; // Второй вариант быстрого пересчета

		this.ComplexFields = new AscWord.ParagraphComplexFieldStack();
	}
	AlignRecalcState.prototype.IsFastRangeRecalc = function()
	{
		return this.RecalcFast;
	};
	AlignRecalcState.prototype.getLogicDocument = function()
	{
		return this.wrapState.Paragraph.GetLogicDocument();
	};
	AlignRecalcState.prototype.getDocumentSettings = function()
	{
		let logicDocument = this.Paragraph.GetLogicDocument();
		if (logicDocument && logicDocument.IsDocumentEditor())
			return logicDocument.getDocumentSettings();

		return AscWord.DEFAULT_DOCUMENT_SETTINGS;
	};
	AlignRecalcState.prototype.getCompatibilityMode = function()
	{
		return this.getDocumentSettings().getCompatibilityMode();
	};
	AlignRecalcState.prototype.getLineTop = function()
	{
		let p = this.Paragraph;
		return p.Pages[this.wrapState.Page].Y + p.Lines[this.wrapState.Line].Top;
	};
	AlignRecalcState.prototype.getLineBottom = function()
	{
		let p = this.Paragraph;
		return p.Pages[this.wrapState.Page].Y + p.Lines[this.wrapState.Line].Bottom;
	};
	AlignRecalcState.prototype.isParagraphStartFromNewPage = function()
	{
		return this.Paragraph.IsStartFromNewPage();
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.Paragraph.AlignRecalcState = AlignRecalcState;

})(window);
