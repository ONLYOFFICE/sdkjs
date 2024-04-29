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
	 * Класс представляющий собой длинный разделитель (который в основном используется для сносок).
	 * @constructor
	 * @extends {AscWord.CRunElementBase}
	 */
	function CRunContinuationSeparator()
	{
		AscWord.CRunElementBase.call(this);
		this.LineW = 0;
	}
	CRunContinuationSeparator.prototype = Object.create(AscWord.CRunElementBase.prototype);
	CRunContinuationSeparator.prototype.constructor = CRunContinuationSeparator;

	CRunContinuationSeparator.prototype.Type         = para_ContinuationSeparator;
	CRunContinuationSeparator.prototype.Get_Type     = function()
	{
		return para_ContinuationSeparator;
	};
	CRunContinuationSeparator.prototype.Draw         = function(X, Y, Context, PDSE)
	{
		var l = X, t = PDSE.LineTop, r = X + this.GetWidth(), b = PDSE.BaseLine;

		Context.p_color(0, 0, 0, 255);
		Context.drawHorLineExt(c_oAscLineDrawingRule.Center, (t + b) / 2, l, r, this.LineW, 0, 0);

		if (editor && editor.ShowParaMarks && Context.DrawFootnoteRect)
		{
			Context.DrawFootnoteRect(X, PDSE.LineTop, this.GetWidth(), PDSE.BaseLine - PDSE.LineTop);
		}
	};
	CRunContinuationSeparator.prototype.Measure      = function(Context, TextPr)
	{
		this.Width        = (50 * AscWord.TEXTWIDTH_DIVIDER) | 0;
		this.WidthVisible = (50 * AscWord.TEXTWIDTH_DIVIDER) | 0;

		this.LineW = (TextPr.FontSize / 18) * g_dKoef_pt_to_mm;
	};
	CRunContinuationSeparator.prototype.Copy         = function()
	{
		return new CRunContinuationSeparator();
	};
	CRunContinuationSeparator.prototype.UpdateWidth = function(PRS)
	{
		var oPara    = PRS.Paragraph;
		var nCurPage = PRS.Page;

		oPara.Parent.Update_ContentIndexing();
		var oLimits = oPara.Parent.Get_PageContentStartPos2(oPara.PageNum, oPara.ColumnNum, nCurPage, oPara.Index);

		var nWidth = (Math.max(oLimits.XLimit - PRS.X, 50) * AscWord.TEXTWIDTH_DIVIDER) | 0;

		this.Width        = nWidth;
		this.WidthVisible = nWidth;
	};
	CRunContinuationSeparator.prototype.IsNeedSaveRecalculateObject = function()
	{
		return true;
	};
	CRunContinuationSeparator.prototype.SaveRecalculateObject = function(isCopy)
	{
		return {
			Width : this.Width
		};
	};
	CRunContinuationSeparator.prototype.LoadRecalculateObject = function(oRecalcObj)
	{
		this.Width        = oRecalcObj.Width;
		this.WidthVisible = oRecalcObj.Width;
	};
	CRunContinuationSeparator.prototype.PrepareRecalculateObject = function()
	{
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunContinuationSeparator = CRunContinuationSeparator;

})(window);
