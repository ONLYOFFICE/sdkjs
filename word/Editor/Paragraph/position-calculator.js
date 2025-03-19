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
	const BLACK_COLOR = new AscWord.CDocumentColor(0, 0, 0);
	
	/**
	 * Class for calculating the current position of the cursor
	 * @param {AscWord.Paragraph} paragraph
	 * @constructor
	 */
	function ParagraphPositionCalculator(paragraph)
	{
		this.paragraph = paragraph
		
		this.page = 0;
		this.line = 0;
		this.range = 0;
		
		this.x = 0;
		this.y = 0;
		
		this.startX = 0;
		
		this.bidi = new AscWord.BidiFlow(this);
		
		this.isNextCurrent = false;
		this.nextRun       = null;
		this.stop          = false;
		
		this.posInfo = {
			x     : 0,
			y     : 0,
			run   : null,
			mathY : -1
		};
		
		// Track prev element for the case when cursor is placed between elements with different directions
		this.prev = {
			x         : 0,
			y         : 0,
			run       : null,
			element   : null,
			inRunPos  : 0,
			find      : false,
			usePos    : false,
			direction : AscBidi.DIRECTION.L
		};
		
		this.complexFields = new AscWord.ParagraphComplexFieldStack();
	}
	ParagraphPositionCalculator.prototype.reset = function(page, line, range)
	{
		this.prev.element = null;
		this.prev.find    = false;
		
		this.isNextCurrent = false;
		this.nextRun       = null;
		this.stop          = false;
		
		this.page  = page;
		this.line  = line;
		this.range = range;
		
		let p = this.paragraph;
		
		this.x = p.Lines[line].Ranges[range].XVisible;
		this.y = p.Pages[page].Y + p.Lines[line].Y;
		
		this.startX = this.x;
		
		if (p.Numbering.checkRange(range, line))
			this.x += p.Numbering.WidthVisible;
		
		this.bidi.begin(this.paragraph.isRtlDirection());
		
		this.complexFields.resetRange(this.paragraph, this.line, this.range);
	};
	ParagraphPositionCalculator.prototype.setNextCurrent = function(run, lastCombItem)
	{
		this.isNextCurrent = true;
		this.nextRun       = run;
		
		if (lastCombItem && lastCombItem.RGap)
		{
			this.bidi.end();
			
			this.posInfo.x = this.x;
			this.posInfo.y = this.y;
			this.posInfo.run = run;
			
			if (lastCombItem.getBidiType() & AscBidi.FLAG.RTL)
				this.posInfo.x -= lastCombItem.GetWidthVisible();
			
			if (lastCombItem.RGapCount)
				this.posInfo.x -= lastCombItem.RGapCount * lastCombItem.RGapShift - (lastCombItem.RGapShift - lastCombItem.RGapCharWidth) / 2;
			else
				this.posInfo.x -= lastCombItem.RGap;
			
			this.isNextCurrent = false;
		}
	};
	ParagraphPositionCalculator.prototype.handleRunElement = function(element, run, isCurrent, isNearFootnoteRef, inRunPos)
	{
		if (!this.complexFields.checkRunElement(element))
			return;
		
		if (para_Drawing === element.Type && !element.IsInline())
		{
			if (isCurrent)
				this.setNextCurrent(run);
			return;
		}
		
		if (this.isNextCurrent)
		{
			isCurrent = true;
			run = this.nextRun;
			this.isNextCurrent = false;
			this.nextRun = null;
		}
		
		if (isCurrent)
			this.prev.find = true;
		
		if (!this.prev.find)
		{
			this.prev.element   = element;
			this.prev.run       = run;
			this.prev.inRunPos  = inRunPos;
			this.prev.usePos    = false;
			this.prev.direction = (element.getBidiType() & AscBidi.FLAG.RTL) ? AscBidi.DIRECTION.R : AscBidi.DIRECTION.L;
		}
		
		this.bidi.add([element, run, isCurrent, isNearFootnoteRef], element.getBidiType());
	};
	ParagraphPositionCalculator.prototype.handleBidiFlow = function(data, direction)
	{
		let element   = data[0];
		let run       = data[1];
		let isCurrent = data[2];
		
		if (element === this.prev.element)
		{
			this.prev.x         = this.x;
			this.prev.y         = this.y;
			this.prev.usePos    = true;
			this.prev.direction = direction;
		}
		
		let w = element.GetWidthVisible();

		// If we encountered the last entered element then we place the cursor after it
		let lastElement = this.prev.element && this.prev.find && this.prev.usePos && this.prev.element.IsText() ? this.prev.element : null;
		if (lastElement && AscWord.checkAsYouTypeEnterText(this.prev.run, this.prev.inRunPos + 1, lastElement.GetCodePoint()))
		{
			this.posInfo.x = this.prev.x;
			this.posInfo.y = this.prev.y;
			
			if (AscBidi.DIRECTION.L === this.prev.direction)
				this.posInfo.x += lastElement.GetWidthVisible();
			
			this.posInfo.run = this.prev.run;
			
			// for comb forms
			if (element.LGap)
				this.posInfo.x += element.LGap;
			
			this.stop = true;
		}
		
		if (isCurrent && !this.stop)
		{
			this.posInfo.x = this.x;
			this.posInfo.y = this.y;
			
			if (direction === AscBidi.DIRECTION.R)
				this.posInfo.x += w;
			
			this.posInfo.run = run;
			
			// for comb forms
			if (element.LGap)
				this.posInfo.x += element.LGap;
		}
		
		this.x += w;
	};
	ParagraphPositionCalculator.prototype.handleParaMath = function(math, isCurrent)
	{
		this.bidi.end();
		
		if (this.isNextCurrent && !this.stop)
		{
			this.posInfo.x   = this.x;
			this.posInfo.y   = this.y;
			this.posInfo.run = this.nextRun;
			
			this.isNextCurrent = false;
			this.nextRun       = null;
		}
		
		let mathW = math.Root.GetWidth(this.line, this.range);
		let x = this.x;
		
		if (isCurrent)
			math.Root.recalculateCursorPosition(this, true);
		
		this.x = x + mathW;
	};
	ParagraphPositionCalculator.prototype.handleMathRun = function(run, isCurrentRun, currentPos)
	{
		this.bidi.end();
		this.bidi.begin(this.paragraph.isRtlDirection());
		
		if (!isCurrentRun)
			return;
		
		let paraMathLocation = run.ParaMath.GetLinePosition(this.line, this.range);
		
		this.x = paraMathLocation.x;
		this.y = paraMathLocation.y;
		
		let mathY = this.y;
		
		if (run.IsEmpty())
		{
			this.x += run.pos.x;
			this.y += run.pos.y;
		}
		else
		{
			let w = 0;
			if (!run.Content[currentPos])
			{
				currentPos = run.Content.length - 1;
				w          = run.Content[currentPos].GetWidthVisible();
			}
			
			let elementLocation = run.Content[currentPos].GetLocationOfLetter();
			this.x += elementLocation.x + w;
			this.y += elementLocation.y;
		}
		
		// TODO: Пометить данное место, как текущее
		this.posInfo.x     = this.x;
		this.posInfo.y     = this.y;
		this.posInfo.mathY = mathY;
		this.posInfo.run   = run;
	};
	ParagraphPositionCalculator.prototype.handleMathElement = function(element)
	{
		let w = element.GetWidth(this.line, this.range);
		this.x += w;
	};
	ParagraphPositionCalculator.prototype.getXY = function()
	{
		this.finalize();
		return {x : this.posInfo.x, y : this.posInfo.y};
	};
	ParagraphPositionCalculator.prototype.getTargetXY = function()
	{
		this.finalize();
		let run = this.posInfo.run;
		if (!run)
			return {x : this.posInfo.x, y : this.posInfo.y, h : 0, ascent : 0};
		
		let textPr = run.getCompiledPr();
		let isNearFootnoteRef = run.IsCurPosNearFootEndnoteReference();
		
		let fontCoef = isNearFootnoteRef ? 1 : textPr.getFontCoef();
		AscCommon.g_oTextMeasurer.SetTextPr(textPr, this.paragraph.getTheme());
		AscCommon.g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII, fontCoef);
		
		let textHeight = AscCommon.g_oTextMeasurer.GetHeight();
		let descent    = Math.abs(AscCommon.g_oTextMeasurer.GetDescender());
		let ascent     = textHeight - descent;
		
		let y = this.posInfo.y - ascent - run.getYOffset();
		if (!isNearFootnoteRef)
		{
			if (AscCommon.vertalign_SubScript === textPr.VertAlign)
				y -= textPr.FontSize * g_dKoef_pt_to_mm * AscCommon.vaKSub;
			else if (AscCommon.vertalign_SuperScript === textPr.VertAlign)
				y -= textPr.FontSize * g_dKoef_pt_to_mm * AscCommon.vaKSuper;
		}
		
		let runParent = run.Parent;
		if (run.IsMathRun() && runParent && runParent.bRoot && runParent.bMath_OneLine)
		{
			let mathBounds = runParent.Get_Bounds();
			
			let y0 = y;
			let y1 = y + textHeight;
			
			let _y = runParent.pos.y - runParent.size.ascent;
			
			let _y0 = this.posInfo.mathY + _y - 0.2 * mathBounds.H;
			let _y1 = this.posInfo.mathY + _y + 1.4 * mathBounds.H;
			
			y0 = Math.max(y0, _y0);
			y1 = Math.min(y1, _y1);
			
			y = y0;
			textHeight = y1 - y0;
		}
		
		return {x : this.posInfo.x, y : y, h : textHeight, ascent : ascent};
	};
	/**
	 * @returns {AscWord.CDocumentColor}
	 */
	ParagraphPositionCalculator.prototype.getTargetColor = function()
	{
		if (!this.finalize())
			return BLACK_COLOR;
		
		let p      = this.paragraph;
		let textPr = this.posInfo.run.getCompiledPr();
		let color  = textPr.Color;
		
		if (textPr.TextFill)
		{
			textPr.TextFill.check(p.getTheme(), p.getColorMap());
			let RGBA = textPr.TextFill.getRGBAColor();
			return new AscWord.CDocumentColor(RGBA.R, RGBA.G, RGBA.B);
		}
		else if (textPr.Unifill)
		{
			textPr.Unifill.check(p.getTheme(), p.getColorMap());
			let RGBA = textPr.Unifill.getRGBAColor();
			return new AscWord.CDocumentColor(RGBA.R, RGBA.G, RGBA.B);
		}
		else if (color.IsAuto())
		{
			return this.posInfo.run.getAutoColor();
		}
		
		return color;
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	/**
	 * @returns {boolean}
	 */
	ParagraphPositionCalculator.prototype.finalize = function()
	{
		this.bidi.end();
		
		if (this.isNextCurrent && !this.stop)
		{
			this.posInfo.x   = this.paragraph.isRtlDirection() ? this.startX : this.x;
			this.posInfo.y   = this.y;
			this.posInfo.run = this.nextRun;
			
			this.isNextCurrent = false;
			this.nextRun       = null;
		}
		
		return !!this.posInfo.run;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.ParagraphPositionCalculator = ParagraphPositionCalculator;
	
})(window);


