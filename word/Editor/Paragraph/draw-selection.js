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
	const EPSILON = AscCommon.TwipsToMM(1);
	
	/**
	 * Class for handling bidirectional flow of text or other content
	 * @param {AscWord.Paragraph} paragraph
	 * @constructor
	 */
	function ParagraphDrawSelectionState(paragraph)
	{
		this.paragraph = paragraph
		
		this.y = 0;
		this.h = 0;
		this.x = 0;
		
		this.bidi = new AscWord.BidiFlow(this);
		
		this.page  = 0;
		this.range = 0;
		this.line  = 0;
		
		this.selectionRanges = [];
		
		this.beginElement = null;
		this.endElement   = null;
		this.beginInfo    = null;
		this.endInfo      = null;
		
		this.anchoredObjects = [];
		
		this.complexFields = new AscWord.ParagraphComplexFieldStack();
	}
	ParagraphDrawSelectionState.prototype.resetPage = function(page)
	{
		this.page = page;
		
		this.complexFields.resetPage(this.paragraph, page);
	};
	ParagraphDrawSelectionState.prototype.resetLine = function(line)
	{
		this.line = line;
		
		let p = this.paragraph;
		this.y = p.Pages[this.page].Y + p.Lines[this.line].Top;
		this.h = p.Lines[this.line].Bottom - p.Lines[this.line].Top;
	};
	ParagraphDrawSelectionState.prototype.beginRange = function(range)
	{
		this.range = range;
		
		this.x = this.paragraph.Lines[this.line].Ranges[this.range].XVisible;
		
		if (this.paragraph.Numbering.checkRange(this.line, this.range))
			this.x += this.paragraph.Numbering.WidthVisible;
		
		this.bidi.begin(this.paragraph.isRtlDirection());
		
		this.beginElement = null;
		this.endElement   = null;
		
		this.flowObjects = [];
	};
	ParagraphDrawSelectionState.prototype.endRange = function()
	{
		this.bidi.end();
	};
	ParagraphDrawSelectionState.prototype.getSelectionRanges = function()
	{
		return this.selectionRanges;
	};
	ParagraphDrawSelectionState.prototype.getAnchoredObjects = function()
	{
		return this.anchoredObjects;
	};
	ParagraphDrawSelectionState.prototype.getBeginInfo = function()
	{
		return this.beginInfo;
	};
	ParagraphDrawSelectionState.prototype.getEndInfo = function()
	{
		return this.endInfo;
	};
	/**
	 * @param element {AscWord.CRunElementBase}
	 * @param isSelected {boolean}
	 */
	ParagraphDrawSelectionState.prototype.handleRunElement = function(element, isSelected)
	{
		if (!this.complexFields.checkRunElement(element))
			return;
		
		if (para_Drawing === element.Type && !element.IsInline())
		{
			if (isSelected)
				this.anchoredObjects.push(element);
			
			return;
		}
		
		if (isSelected)
		{
			if (!this.beginElement)
				this.beginElement = element;
			
			this.endElement = element;
		}
		
		this.bidi.add([element, isSelected], element.getBidiType());
	};
	ParagraphDrawSelectionState.prototype.handleBidiFlow = function(data, direction)
	{
		let element    = data[0];
		let isSelected = data[1];
		
		let w = element.GetWidthVisible();
		if (isSelected)
		{
			let lastRange = this.selectionRanges.length ? this.selectionRanges[this.selectionRanges.length - 1] : null;
			if (lastRange && Math.abs(lastRange.x + lastRange.w - this.x) < EPSILON)
				lastRange.w += w;
			else
				this.selectionRanges.push({x : this.x, w : w, y : this.y, h : this.h});
			
			// TODO: Сейчас мы используем начальный типа символа, нужно использовать тип расчитанный алгоритомом bidi
			if (element === this.beginElement)
			{
				this.beginInfo = {
					x : direction === AscBidi.DIRECTION.R ? this.x + w : this.x,
					w : 0,
					y : this.y,
					h : this.h
				};
			}
			
			if (element === this.endElement)
			{
				this.endInfo = {
					x : direction === AscBidi.DIRECTION.R ? this.x : this.x + w,
					w : 0,
					y : this.y,
					h : this.h
				};
			}
		}
		
		this.x += w;
	};
	ParagraphDrawSelectionState.prototype.handleMathElement = function(mathElement, isSelected)
	{
		this.bidi.end();
		
		let w = mathElement.GetWidth(this.line, this.range);
		
		if (isSelected)
		{
			let lastRange = this.selectionRanges.length ? this.selectionRanges[this.selectionRanges.length - 1] : null;
			if (lastRange && Math.abs(lastRange.x + lastRange.w - this.x) < EPSILON)
				lastRange.w += w;
			else
				this.selectionRanges.push({x : this.x, w : w, y : this.y, h : this.h});
		}
		
		this.x += w;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.ParagraphDrawSelectionState = ParagraphDrawSelectionState;
	
})(window);


