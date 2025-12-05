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
	const EPSILON  = 0.001;
	const MAX_DIFF = 1000000;
	
	/**
	 * Class for searching position in the paragraph
	 * @constructor
	 */
	function ParagraphSearchPositionXY()
	{
		this.line  = 0;
		this.range = 0;
		this.page  = 0;
		
		this.paragraph  = null;
		this.centerMode = true;  // Search the closest position (relative to the middle of the element), or we search a position beyond the specified x-coordinate
		this.stepEnd    = false; // Search for position beyond the mark of paragraph
		
		this.curX    = 0;
		this.curY    = 0;
		this.x       = 0;
		this.y       = 0;
		this.diffX   = MAX_DIFF;
		this.diffAbs = MAX_DIFF;
		
		this.numbering = false;
		this.inText    = false;
		this.inTextX   = false;
		this.paraEnd   = false;
		
		this.bidiFlow = new AscWord.BidiFlow(this);
		
		this.emptyRun = null;
		this.emptyRunHandler = {};
		
		// TODO: Unite with CRunWithPosition class
		this.pos     = null;
		this.posInfo = {
			run : null,
			pos : 0
		};
		
		this.inTextPos     = null;
		this.inTextPosInfo = {
			run : null,
			pos : 0
		};
		
		this.complexFields = new AscWord.ParagraphComplexFieldStack();
	}
	ParagraphSearchPositionXY.prototype.init = function(paragraph, stepEnd, centerMode)
	{
		this.paragraph  = paragraph;
		this.stepEnd    = undefined !== stepEnd ? stepEnd : false;
		this.centerMode = undefined !== centerMode ? centerMode : true;
		
		this.bidiFlow.begin(paragraph.isRtlDirection());
	};
	ParagraphSearchPositionXY.prototype.reset = function()
	{
		this.bidiFlow.end();
		this.emptyRun = null;
		this.emptyRunHandler = {};
	};
	ParagraphSearchPositionXY.prototype.setDiff = function(diff)
	{
		this.diffX   = Math.abs(diff);
		this.diffAbs = diff;
	};
	ParagraphSearchPositionXY.prototype.searchByXY = function(x, y, page)
	{
		if (this.correctPageAndLineNumber(page))
			this.line = this.calculateLineNumber(y, this.page);
		else
			y = undefined;
		
		this.searchByLine(x, this.line, page, y);
	};
	ParagraphSearchPositionXY.prototype.searchByLine = function(x, line, page, y)
	{
		this.pos       = null;
		this.inTextPos = null;
		
		this.line = line;
		this.correctPageAndLineNumber(page);
		
		this.range = this.calculateRangeNumber(x);
		if (-1 === this.range)
			return;
		
		this.complexFields.resetRange(this.paragraph, this.line, this.range);
		
		let para = this.paragraph;
		let paraRange = para.Lines[this.line].Ranges[this.range];
		
		this.y    = undefined === y ? 0 : y;
		this.x    = x;
		this.curX = paraRange.XVisible;
		this.curY = 0;
		
		let isRtl = para.isRtlDirection();
		if (para.Numbering.checkRange(this.range, this.line) && !isRtl)
			this.checkNumbering(isRtl);
		
		let startPos = paraRange.StartPos;
		let endPos   = paraRange.EndPos;
		
		for (let pos = startPos; pos <= endPos; ++pos)
		{
			para.Content[pos].getParagraphContentPosByXY(this);
		}
		
		this.bidiFlow.end();
		
		if (para.Numbering.checkRange(this.range, this.line) && isRtl)
			this.checkNumbering(isRtl);
		
		this.checkRangeBounds(x, paraRange);
		
		this.checkInText()
		
		if (this.diffX > MAX_DIFF - 1)
		{
			this.pos       = para.Get_StartRangePos2(this.line, this.range);
			this.inTextPos = this.pos.Copy();
		}
	};
	ParagraphSearchPositionXY.prototype.handleRun = function(run)
	{
		if (this.emptyRun)
			return;
		
		if (!run.IsEmpty())
			this.emptyRun = null;
		else
			this.emptyRun = run;
	};
	ParagraphSearchPositionXY.prototype.handleEmptyRun = function(run)
	{
		let curX = this.curX;
		if (run.IsMathRun())
		{
			let mathPos = run.ParaMath.GetLinePosition(this.line, this.range);
			curX        = mathPos.x + run.pos.x;
		}
		
		let diff = this.x - curX;
		if (this.checkPosition(diff))
		{
			this.setDiff(diff);
			this.posInfo.run = run;
			this.posInfo.pos = run.GetElementsCount();
		}
	};
	ParagraphSearchPositionXY.prototype.handleParaMath = function(math)
	{
		if (this.emptyRun)
		{
			this.handleEmptyRun(this.emptyRun);
			this.emptyRun = null;
		}
		
		let curX = this.curX;
		let mathW = math.Root.GetWidth(this.line, this.range);
		
		if ((curX <= this.x && this.x < curX + mathW) || this.diffX > MAX_DIFF - 1)
		{
			let diffX = this.diffX;
			this.setDiff(MAX_DIFF);
			
			math.Root.getParagraphContentPosByXY(this);
			
			if (this.inText)
				this.diffX = EPSILON;
			
			// TODO: Пересмотреть данную проверку. Надо выяснить насколько сильно она вообще нужна
			// Если мы попадаем в формулу, тогда не ищем позицию вне ее. За исключением, случая когда формула идет в начале
			// строки. Потому что в последнем случае из формулы 100% придет true, а позиция, возможно, находится за формулой.
			if (this.diffX < MAX_DIFF - 1 && diffX < MAX_DIFF - 1)
				this.diffX = 0;
			else if (this.diffX > MAX_DIFF - 1)
				this.diffX = diffX;
		}
		
		// Такое возможно, если все элементы до этого (в том числе и этот) были пустыми, тогда, чтобы не возвращать
		// неправильную позицию вернем позицию начала данного элемента.
		if (this.diffX > MAX_DIFF - 1)
		{
			this.pos = this.getStartPosOfElement(math);
			this.setDiff(0);
		}
		
		this.curX = curX + mathW;
		
		this.reset();
	};
	ParagraphSearchPositionXY.prototype.handleMathBase = function(base)
	{
		if (this.emptyRun)
		{
			this.handleEmptyRun(this.emptyRun);
			this.emptyRun = null;
		}
		
		if (!base.Content.length)
			return;
		
		let startPos = 0;
		let endPos   = base.Content.length - 1;
		
		if (!base.bOneLine)
		{
			let rangePos = base.getRangePos(this.line, this.range);
			startPos = rangePos[0];
			endPos   = rangePos[1];
		}
		
		let x = this.x;
		let y = this.y;
		
		let targetPos    = -1;
		let targetBounds = null;
		let diff         = null;
		for (let pos = 0; pos < base.Content.length; ++pos)
		{
			if (pos < startPos || pos > endPos)
				continue;
			
			let bounds = base.Content[pos].Get_LineBound(this.line, this.range);
			if (!bounds || bounds.W < EPSILON || bounds.H < EPSILON)
				continue;
			
			if (bounds.X <= x && x <= bounds.X + bounds.W && bounds.Y <= y && y <= bounds.Y + bounds.H)
			{
				targetPos    = pos;
				targetBounds = bounds;
				break;
			}
			else
			{
				// TODO: Rework this hit check
				let diffX = x - (bounds.X + bounds.W / 2);
				let diffY = y - (bounds.Y + bounds.H / 2);
				
				let curDiff = diffX * diffX + diffY * diffY;
				if (null === diff || diff > curDiff)
				{
					diff         = curDiff;
					targetPos    = pos;
					targetBounds = bounds;
				}
			}
		}
		
		if (-1 === targetPos)
			return;
		
		this.curX = targetBounds.X;
		this.curY = targetBounds.Y;
		
		base.Content[targetPos].getParagraphContentPosByXY(this);
		this.reset();
	};
	ParagraphSearchPositionXY.prototype.handleRunElement = function(element, run, inRunPos)
	{
		if (!this.complexFields.checkRunElement(element))
			return;
		
		if (this.emptyRun)
		{
			this.emptyRunHandler[element] = this.emptyRun;
			this.emptyRun = null;
		}
		
		this.bidiFlow.add([element, run, inRunPos], element.getBidiType());
	};
	ParagraphSearchPositionXY.prototype.handleBidiFlow = function(data, direction)
	{
		let item     = data[0];
		let run      = data[1];
		let inRunPos = data[2];
		
		if (this.emptyRunHandler[item] && direction === AscBidi.DIRECTION.L)
		{
			this.handleEmptyRun(this.emptyRunHandler[item]);
			delete this.emptyRunHandler[item];
		}
		
		let w = 0;
		if (!item.IsDrawing() || item.IsInline())
			w = item.GetWidthVisible();
		
		if (run.IsMathRun())
		{
			let posLine = run.ParaMath.GetLinePosition(this.line, this.range);
			let loc     = item.GetLocationOfLetter();
			this.curX   = posLine.x + loc.x;
		}
		
		let diffL = this.x - this.curX;
		let diffR = this.x - this.curX - w + (item.RGap ? item.RGap : 0);
		
		if (-EPSILON <= diffL && diffL <= w + EPSILON)
		{
			this.inTextX = true;
			this.inTextPosInfo.run = run;
			this.inTextPosInfo.pos = inRunPos;
		}
		
		if (direction === AscBidi.DIRECTION.R)
		{
			let tmp = diffR;
			diffR = diffL;
			diffL = tmp;
		}
		
		if (this.checkPosition(diffL))
		{
			this.setDiff(diffL);
			this.posInfo.run = run;
			this.posInfo.pos = inRunPos;
		}
		
		if (!item.IsBreak() && this.checkPosition(diffR))
		{
			if (item.IsParaEnd())
				this.paraEnd = true;
			
			if (!item.IsParaEnd() || this.stepEnd)
			{
				if (item.RGap)
					diffR = Math.min(diffR, diffR - item.RGap);
				
				this.setDiff(diffR);
				this.posInfo.run = run;
				this.posInfo.pos = inRunPos + 1;
			}
		}
		
		this.curX += w;
		
		if (this.emptyRunHandler[item] && direction === AscBidi.DIRECTION.R)
		{
			this.handleEmptyRun(this.emptyRunHandler[item]);
			delete this.emptyRunHandler[item];
		}
	};
	ParagraphSearchPositionXY.prototype.getPos = function()
	{
		if (this.pos)
			return this.pos;
		
		this.pos = this.getPosByPosInfo(this.posInfo)
		return this.pos;
	};
	ParagraphSearchPositionXY.prototype.getInTextPos = function()
	{
		if (this.inTextPos)
			return this.inTextPos;
		
		this.inTextPos = this.getPosByPosInfo(this.inTextPosInfo);
		return this.inTextPos;
	};
	ParagraphSearchPositionXY.prototype.getLine = function()
	{
		return this.line;
	};
	ParagraphSearchPositionXY.prototype.getRange = function()
	{
		return this.range;
	};
	ParagraphSearchPositionXY.prototype.isNumbering = function()
	{
		return this.numbering;
	};
	ParagraphSearchPositionXY.prototype.isBeyondEnd = function()
	{
		return this.paraEnd;
	};
	ParagraphSearchPositionXY.prototype.isInText = function()
	{
		return this.inText;
	};
	ParagraphSearchPositionXY.prototype.isInTextByX = function()
	{
		return this.inTextX;
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	ParagraphSearchPositionXY.prototype.correctPageAndLineNumber = function(page)
	{
		this.page = (-1 === page || undefined === page || null === page ? 0 : page);
		
		let pageCount = this.paragraph.getPageCount();
		if (this.page >= pageCount)
		{
			this.page = pageCount - 1;
			this.line = this.paragraph.getLineCount() - 1;
			return false;
		}
		else if (this.page < 0)
		{
			this.page = 0;
			this.line = 0;
			return false;
		}
		
		return true;
	};
	ParagraphSearchPositionXY.prototype.calculateLineNumber = function(y, page)
	{
		let p = this.paragraph;
		
		let line     = p.Pages[page].FirstLine;
		let lastLine = page >= p.getPageCount() - 1 ? p.getLineCount() - 1 : p.Pages[page + 1].FirstLine - 1;
		
		for (; line < lastLine; ++line)
		{
			let lineY = p.Pages[page].Y + p.Lines[line].Y + p.Lines[line].Metrics.Descent + p.Lines[line].Metrics.LineGap;
			if (y < lineY)
				break;
		}
		
		return line;
	};
	ParagraphSearchPositionXY.prototype.calculateRangeNumber = function(x)
	{
		let p = this.paragraph;
		
		let rangeCount = p.Lines[this.line].Ranges.length;
		if (rangeCount <= 0)
			return -1;
		else if (1 === rangeCount)
			return 0;
		
		let range = 0;
		if (p.isRtlDirection())
		{
			for (; range < rangeCount - 1; ++range)
			{
				let currRange = p.Lines[this.line].Ranges[range];
				let nextRange = p.Lines[this.line].Ranges[range + 1];
				if (x > (currRange.X + nextRange.XEnd) / 2 || currRange.WEnd > 0.001)
					break;
			}
		}
		else
		{
			for (; range < rangeCount - 1; ++range)
			{
				let currRange = p.Lines[this.line].Ranges[range];
				let nextRange = p.Lines[this.line].Ranges[range + 1];
				if (x < (currRange.XEnd + nextRange.X) / 2 || currRange.WEnd > 0.001)
					break;
			}
		}
		
		return Math.max(0, Math.min(range, rangeCount - 1));
	};
	ParagraphSearchPositionXY.prototype.checkNumbering = function(isRtl)
	{
		let p = this.paragraph;
		let numPr = p.GetNumPr();
		let prevNumPr = p.GetPrChangeNumPr();
		let numLvl = null;
		if (numPr && numPr.IsValid())
			numLvl = p.Parent.GetNumbering().GetNum(numPr.NumId).GetLvl(numPr.Lvl);
		else if (prevNumPr && prevNumPr.IsValid())
			numLvl = p.Parent.GetNumbering().GetNum(prevNumPr.NumId).GetLvl(undefined !== prevNumPr.Lvl && null !== prevNumPr.Lvl ? prevNumPr.Lvl : 0);
		
		let numWidthVisible = p.Numbering.WidthVisible;
		if (para_Numbering === p.Numbering.Type && numLvl)
		{
			let numJc = numLvl.GetJc();
			
			let numX0 = this.curX;
			let numX1 = this.curX;
			
			let numWidth = p.Numbering.WidthNum;
			if (isRtl)
			{
				numX0 += numWidthVisible;
				numX1 += numWidthVisible;
				
				if (AscCommon.align_Right === numJc)
				{
					numX1 += numWidth;
				}
				else if (AscCommon.align_Center === numJc)
				{
					numX0 -= numWidth / 2;
					numX1 += numWidth / 2;
				}
				else// if (AscCommon.align_Left === numJc)
				{
					numX0 -= numWidth;
				}
			}
			else
			{
				if (AscCommon.align_Right === numJc)
				{
					numX0 -= numWidth;
				}
				else if (AscCommon.align_Center === numJc)
				{
					numX0 -= numWidth / 2;
					numX1 += numWidth / 2;
				}
				else// if (AscCommon.align_Left === numJc)
				{
					numX1 += numWidth;
				}
			}
			
			if (numX0 <= this.x && this.x <= numX1)
				this.numbering = true;
		}
		
		this.curX += numWidthVisible;
	};
	ParagraphSearchPositionXY.prototype.checkPosition = function(diff)
	{
		return (((diff <= 0 && Math.abs(diff) < this.diffX - EPSILON) || (diff > 0 && diff < this.diffX + EPSILON))
			&& (this.centerMode || this.x > this.curX));
	}
	ParagraphSearchPositionXY.prototype.checkRangeBounds = function(x, range)
	{
		if (this.stepEnd)
			return;
		
		let para = this.paragraph;
		if (para.isRtlDirection())
		{
			if (x < range.XVisible)
			{
				this.setDiff(range.XVisible - x);
				let pos = para.Get_EndRangePos2(this.line, this.range, false);
				
				this.pos       = para.GetCursorPlaceablePos(pos);
				this.inTextPos = this.pos.Copy();
				this.inTextX   = false;
			}
			else if (x > range.XEndVisible)
			{
				this.setDiff(range.XEndVisible - x);
				let pos = para.Get_StartRangePos2(this.line, this.range);
				
				this.pos       = para.GetCursorPlaceablePos(pos);
				this.inTextPos = this.pos.Copy();
				this.inTextX   = false;
			}
		}
		else
		{
			if (x < range.XVisible)
			{
				this.setDiff(range.XVisible - x);
				let pos = para.Get_StartRangePos2(this.line, this.range);
				
				this.pos       = para.GetCursorPlaceablePos(pos);
				this.inTextPos = this.pos.Copy();
				this.inTextX   = false;
			}
			else if (x > range.XEndVisible)
			{
				this.setDiff(range.XEndVisible - x);
				let pos = para.Get_EndRangePos2(this.line, this.range, false);
				
				this.pos       = para.GetCursorPlaceablePos(pos);
				this.inTextPos = this.pos.Copy();
				this.inTextX   = false;
			}
		}
	};
	ParagraphSearchPositionXY.prototype.checkInText = function()
	{
		this.inText = false;
		if (!this.inTextX || undefined === this.y)
			return;
		
		let p = this.paragraph;
		
		let lineTop    = p.Pages[this.page].Y + p.Lines[this.line].Y - p.Lines[this.line].Metrics.Ascent - EPSILON;
		let lineBottom = p.Pages[this.page].Y + p.Lines[this.line].Y + p.Lines[this.line].Metrics.Descent + p.Lines[this.line].Metrics.LineGap + EPSILON;
		
		this.inText = lineTop <= this.y && this.y <= lineBottom;
	};
	ParagraphSearchPositionXY.prototype.getPosByPosInfo = function(posInfo)
	{
		let paraPos = this.paragraph.GetPosByElement(posInfo.run);
		paraPos.Update(posInfo.pos, paraPos.GetDepth() + 1);
		return this.paragraph.private_GetClosestPosInCombiningMark(paraPos, this.diffAbs);
	};
	ParagraphSearchPositionXY.prototype.getStartPosOfElement = function(element)
	{
		let paraPos = this.paragraph.GetPosByElement(element);
		element.Get_StartPos(paraPos, paraPos.GetDepth() + 1);
		return paraPos;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.ParagraphSearchPositionXY = ParagraphSearchPositionXY;
	
})(window);


