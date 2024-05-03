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

