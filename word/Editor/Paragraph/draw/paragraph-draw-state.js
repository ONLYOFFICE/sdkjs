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
	 * The main paragraph drawing state
	 * @constructor
	 */
	function ParagraphDrawState()
	{
		AscWord.ParagraphRecalculateStateBase.call(this);
		
		this.highlightState  = new AscWord.ParagraphHighlightDrawState(this);
		this.runElementState = new AscWord.ParagraphContentDrawState(this);
		this.lineState       = new AscWord.ParagraphLineDrawState(this);
		
		this.paragraph  = null;
		this.graphics   = null;
		this.theme      = null;
		this.colorMap   = null;
		this.bgColor    = null;
		this.compiledPr = null;
	}
	
	ParagraphDrawState.prototype = Object.create(AscWord.ParagraphRecalculateStateBase.prototype);
	ParagraphDrawState.prototype.constructor = ParagraphDrawState;
	
	ParagraphDrawState.prototype.init = function(paragraph, graphics)
	{
		this.paragraph = paragraph;
		this.graphics  = graphics;
		
		this.theme      = paragraph.getTheme();
		this.colorMap   = paragraph.getColorMap();
		this.compiledPr = paragraph.getCompiledPr();
		
		let paraParent = paragraph.GetParent();
		let bgColor    = null;
		if (this.compiledPr.ParaPr.Shd && !this.compiledPr.ParaPr.Shd.IsNil())
		{
			bgColor = this.compiledPr.ParaPr.Shd.GetSimpleColor(this.theme, this.colorMap);
			if (bgColor.Auto && paraParent)
				bgColor = paraParent.Get_TextBackGroundColor();
		}
		else if (paraParent)
		{
			bgColor = paraParent.Get_TextBackGroundColor();
		}
		this.bgColor = bgColor;
		
		this.highlightState.init(paragraph, graphics);
		this.runElementState.init();
		this.lineState.init(paragraph, graphics);
	};
	ParagraphDrawState.prototype.getHighlightState = function()
	{
		return this.highlightState;
	};
	ParagraphDrawState.prototype.getRunElementState = function()
	{
		return this.runElementState;
	};
	ParagraphDrawState.prototype.getLineState = function()
	{
		return this.lineState;
	};
	ParagraphDrawState.prototype.getParagraph = function()
	{
		return this.paragraph;
	};
	ParagraphDrawState.prototype.getGraphics = function()
	{
		return this.graphics;
	};
	ParagraphDrawState.prototype.getTheme = function()
	{
		return this.theme;
	};
	ParagraphDrawState.prototype.getColorMap = function()
	{
		return this.colorMap;
	};
	ParagraphDrawState.prototype.getBgColor = function()
	{
		return this.bgColor;
	};
	ParagraphDrawState.prototype.getParagraphCompiledPr = function()
	{
		return this.compiledPr;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.ParagraphDrawState = ParagraphDrawState;
	
})(window);

