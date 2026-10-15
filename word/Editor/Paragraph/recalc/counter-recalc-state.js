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
	function CounterRecalcState(wrapState)
	{
		AscWord.ParagraphBidiFlow.call(this);
		this.wrapState   = wrapState;
		this.Paragraph   = undefined;
		this.Range       = undefined;
		this.Rtl         = false;
		this.Word        = false;
		this.HasContent  = false; // Whether a non-space element has been encountered (including a tab)
		this.SpaceLen    = 0;
		this.SpacesCount = 0;

		this.SpacesLeft  = 0; // Spaces before the first non-space element
		this.SpacesLeftW = 0; // Width of the spaces before the first non-space element
		this.Spaces      = 0; // Spaces between the first and last non-space elements
		
		// Spaces after the last non-space element are stored in SpacesCount until the end of the range
		// SpacesRight  = SpacesCount
		// SpacesRightW = SpaceLen

		this.HasTab      = false;
		this.WordsFirst  = 0; // Words before the first tab
		this.SpacesFirst = 0; // Spaces before the first tab (SpacesLeft is not counted here)
		this.WordsLast   = 0; // Words after the last tab
		this.SpacesLast  = 0; // Spaces after the last tab (SpacesRight, i.e. SpacesCount, is not counted here)

		this.Letters     = 0; // Letters after the last tab
		this.LettersSkip = 0; // Letters before the last tab

		this.CurRange = 0;
		this.CurLine  = 0;
		
		this.ComplexFields = new AscWord.ParagraphComplexFieldStack();
	}
	CounterRecalcState.prototype = Object.create(AscWord.ParagraphBidiFlow.prototype);
	CounterRecalcState.prototype.constructor = CounterRecalcState;
	
	CounterRecalcState.prototype.beginRange = function(paragraph, range, lineNum, rangeNum)
	{
		this.Paragraph   = paragraph;
		this.Range       = range;
		this.Rtl         = paragraph.isRtlDirection();
		this.Word        = false;
		this.HasContent  = false;
		this.SpaceLen    = 0;
		this.SpacesCount = 0;

		this.CurRange = rangeNum;
		this.CurLine  = lineNum;

		this.SpacesLeft  = 0;
		this.SpacesLeftW = 0;
		this.Spaces      = 0;

		this.HasTab      = false;
		this.WordsFirst  = 0;
		this.SpacesFirst = 0;
		this.WordsLast   = 0;
		this.SpacesLast  = 0;

		this.Letters     = 0;
		this.LettersSkip = 0;

		this.ParaEnd   = false;
		this.LineBreak = false;
		
		this.Range.W      = 0;
		this.Range.WEnd   = 0;
		this.Range.WBreak = 0;
		
		if (this.Paragraph.Numbering.checkRange(this.CurRange, this.CurLine))
			this.Range.W += this.Paragraph.Numbering.WidthVisible;
		
		this.bidiFlow.begin(this.Rtl);
	};
	CounterRecalcState.prototype.endRange = function()
	{
		this.bidiFlow.end();

		if (!this.HasTab)
		{
			this.WordsFirst  = this.WordsLast;
			this.SpacesFirst = this.SpacesLast;
		}

		this.Range.Spaces      = this.Spaces;
		this.Range.SpacesLeft  = this.SpacesLeft;
		this.Range.SpacesRight = this.SpacesCount;

		if (this.Rtl)
			this.Range.W += this.SpaceLen - this.SpacesLeftW;
	};
	CounterRecalcState.prototype.handleRunElement = function(element, run)
	{
		let type = element.Type;
		if (this.ComplexFields.isHiddenFieldContent() && para_End !== type && para_FieldChar !== type)
			return;
		
		let isHiddenCFPart = this.ComplexFields.isHiddenComplexFieldPart();
		if (isHiddenCFPart && para_End !== type && para_InstrText !== type && para_FieldChar !== type)
			return;
		
		if (para_FieldChar === type)
		{
			if (this.isFastRecalculation())
				this.ComplexFields.processFieldChar(element);
			else
				this.ComplexFields.processFieldCharAndCollectComplexField(element);
		}
		
		this.bidiFlow.add([para_RunBase, element, run], element.getBidiType());
	};
	CounterRecalcState.prototype.handleParaMath = function(paraMath)
	{
		this.bidiFlow.add([para_Math, paraMath], paraMath.getBidiType());
	};
	CounterRecalcState.prototype.handleMathBase = function(mathBase)
	{
		this.bidiFlow.add([para_Math_Composition, mathBase], mathBase.getBidiType());
	};
	/**
	 * @param {AscWord.CRunElementBase} element
	 * @param {AscWord.CRun} run
	 */
	CounterRecalcState.prototype.handleBidiFlowRunElement = function(data, direction)
	{
		let element = data[1];
		let run     = data[2];
		
		let type = element.Type;
		let isHiddenCFPart = this.ComplexFields.isHiddenComplexFieldPart();
		if (!isHiddenCFPart && para_InstrText === type)
			type = para_Text;
		
		let textPr = run.Get_CompiledPr(false);
		switch (type)
		{
			case para_Sym:
			case para_Text:
			case para_FootnoteReference:
			case para_FootnoteRef:
			case para_EndnoteReference:
			case para_EndnoteRef:
			case para_Separator:
			case para_ContinuationSeparator:
			{
				this.Letters++;

				this.flushSpaces();

				if (!this.Word)
				{
					this.Word = true;
					this.addWord();
				}

				this.Range.W += element.GetWidth(textPr);
				if (element.IsSpaceAfter())
					this.Word = false;

				break;
			}
			case para_Math_Text:
			case para_Math_Placeholder:
			case para_Math_Ampersand:
			case para_Math_BreakOperator:
			{
				this.Letters++;
				this.HasContent = true;
				this.Range.W += element.GetWidth() / AscWord.TEXTWIDTH_DIVIDER;
				break;
			}
			case para_Space:
			{
				if (true === this.Word)
				{
					this.Word        = false;
					this.SpacesCount = 1;
					this.SpaceLen    = element.GetWidth();
				}
				else
				{
					this.SpacesCount++;
					this.SpaceLen += element.GetWidth();
				}
				break;
			}
			case para_Drawing:
			{
				if (!element.IsInline() && !this.Paragraph.Parent.Is_DrawingShape())
					break;
				
				this.flushSpaces();
				this.addWord();
				this.Word = false;
				this.Range.W += element.GetWidth();
				break;
			}
			case para_PageNum:
			case para_PageCount:
			{
				this.flushSpaces();
				this.addWord();
				this.Word = false;
				this.Range.W += element.GetWidth();
				break;
			}
			case para_Tab:
			{
				this.flushSpaces();
				
				this.Range.W += element.GetWidth();
				this.Word        = false;
				this.HasContent  = true;

				if (!this.HasTab)
				{
					this.HasTab      = true;
					this.WordsFirst  = this.WordsLast;
					this.SpacesFirst = this.SpacesLast;
				}

				this.LettersSkip += this.Letters;

				this.WordsLast  = 0;
				this.SpacesLast = 0;
				this.Letters    = 0;
				break;
			}
			case para_NewLine:
			{
				if (true === this.Word && this.WordsLast > 1)
					this.Spaces += this.SpacesCount;

				this.SpacesCount = 0;
				this.Word        = false;
				this.LineBreak   = true;
				break;
			}
			case para_End:
			{
				if (true === this.Word)
					this.Spaces += this.SpacesCount;

				this.ParaEnd = true;
				break;
			}
			case para_InstrText:
			{
				if (this.isFastRecalculation() || reviewtype_Remove === run.GetReviewType())
					break;

				this.ComplexFields.processInstruction(element);
				break;
			}
			case para_FieldChar:
			{
				if (!element.IsVisual())
					break;

				this.flushSpaces();
				this.addWord();
				this.Word = false;
				this.Range.W += element.GetWidth();

				break;
			}
		}
	};
	CounterRecalcState.prototype.handleBidiFlowParaMath = function(data, direction)
	{
		let paraMath = data[1];
		
		let spaceLen    = this.SpaceLen;
		let breakBefore = paraMath.Is_BrkBinBefore();
		
		paraMath.Root.UpdateOperators(this.CurLine, this.CurRange, !!breakBefore, !breakBefore);
		paraMath.Root.recalculateRangeWidth(this);

		this.Range.W        += this.SpaceLen - spaceLen;
		this.Range.SpaceLen = spaceLen;

		this.addWord();
	};
	CounterRecalcState.prototype.handleBidiFlowMathBase = function(data, direction)
	{
		let mathBase = data[1];
		mathBase._recalculateRangeWidth(this);
	};
	CounterRecalcState.prototype.isFastRecalculation = function()
	{
		return this.wrapState.isFastRecalculation();
	};
	CounterRecalcState.prototype.addWord = function()
	{
		this.HasContent = true;
		this.WordsLast++;
	};
	CounterRecalcState.prototype.flushSpaces = function()
	{
		if (!this.HasContent)
		{
			this.SpacesLeft  += this.SpacesCount;
			this.SpacesLeftW += this.SpaceLen;
		}
		else
		{
			this.Spaces     += this.SpacesCount;
			this.SpacesLast += this.SpacesCount;
		}

		this.Range.W += this.SpaceLen;

		this.SpaceLen    = 0;
		this.SpacesCount = 0;
	};
	/**
	 * Number of words in the part of the range which is stretched when justifying
	 * @returns {number}
	 */
	CounterRecalcState.prototype.getJustifyWords = function()
	{
		return this.Rtl ? this.WordsFirst : this.WordsLast;
	};
	/**
	 * Number of spaces whose width is increased when justifying
	 * @returns {number}
	 */
	CounterRecalcState.prototype.getJustifySpaces = function()
	{
		return this.Rtl ? this.SpacesFirst : this.SpacesLast;
	};
	/**
	 * Number of spaces at the start of the range whose width is not changed
	 * @returns {number}
	 */
	CounterRecalcState.prototype.getJustifySpacesSkip = function()
	{
		return this.Rtl ? this.SpacesLeft : this.SpacesLeft + this.Spaces - this.SpacesLast;
	};
	/**
	 * Number of letters at the start of the range whose width is not changed
	 * @returns {number}
	 */
	CounterRecalcState.prototype.getJustifyLettersSkip = function()
	{
		return this.Rtl ? 0 : this.LettersSkip;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.Paragraph.CounterRecalcState = CounterRecalcState;
	
})(window);
