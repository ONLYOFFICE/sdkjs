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
	 * Класс представляющий собой разделитель (который в основном используется для сносок).
	 * @constructor
	 * @extends {AscWord.CRunElementBase}
	 */
	function CRunSeparator()
	{
		AscWord.CRunElementBase.call(this);
		this.LineW = 0;
	}
	CRunSeparator.prototype = Object.create(AscWord.CRunElementBase.prototype);
	CRunSeparator.prototype.constructor = CRunSeparator;

	CRunSeparator.prototype.Type     = para_Separator;
	CRunSeparator.prototype.Get_Type = function()
	{
		return para_Separator;
	};
	CRunSeparator.prototype.Draw     = function(X, Y, Context, PDSE)
	{
		var l = X, t = PDSE.LineTop, r = X + this.GetWidth(), b = PDSE.BaseLine;

		Context.p_color(0, 0, 0, 255);
		Context.drawHorLineExt(c_oAscLineDrawingRule.Center, (t + b) / 2, l, r, this.LineW, 0, 0);

		if (editor && editor.ShowParaMarks && Context.DrawFootnoteRect)
		{
			Context.DrawFootnoteRect(X, PDSE.LineTop, this.GetWidth(), PDSE.BaseLine - PDSE.LineTop);
		}
	};
	CRunSeparator.prototype.Measure  = function(Context, TextPr)
	{
		this.Width        = (50 * AscWord.TEXTWIDTH_DIVIDER) | 0;
		this.WidthVisible = (50 * AscWord.TEXTWIDTH_DIVIDER) | 0;

		this.LineW = (TextPr.FontSize / 18) * g_dKoef_pt_to_mm;
	};
	CRunSeparator.prototype.Copy     = function()
	{
		return new CRunSeparator();
	};
	CRunSeparator.prototype.UpdateWidth = function(PRS)
	{
		let contentFrame = PRS.Paragraph.GetPageContentFrame(PRS.Page);
		
		let width = (Math.min(50, (contentFrame.XLimit - contentFrame.X)) * AscWord.TEXTWIDTH_DIVIDER) | 0;
		
		this.Width        = width;
		this.WidthVisible = width;
	};
	CRunSeparator.prototype.IsNeedSaveRecalculateObject = function()
	{
		return true;
	};
	CRunSeparator.prototype.SaveRecalculateObject = function(isCopy)
	{
		return {
			Width : this.Width
		};
	};
	CRunSeparator.prototype.LoadRecalculateObject = function(oRecalcObj)
	{
		this.Width        = oRecalcObj.Width;
		this.WidthVisible = oRecalcObj.Width;
	};
	CRunSeparator.prototype.PrepareRecalculateObject = function()
	{
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunSeparator = CRunSeparator;

})(window);
