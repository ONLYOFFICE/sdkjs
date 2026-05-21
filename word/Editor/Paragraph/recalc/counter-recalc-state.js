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
		this.wrapState   = wrapState;
		this.Paragraph   = undefined;
		this.Range       = undefined;
		this.Word        = false;
		this.SpaceLen    = 0;
		this.SpacesCount = 0;

		this.Words       = 0;
		this.Spaces      = 0;
		this.Letters     = 0;
		this.SpacesSkip  = 0;
		this.LettersSkip = 0;

		this.ComplexFields = new AscWord.ParagraphComplexFieldStack();
	}
	CounterRecalcState.prototype.beginRange = function(paragraph, range)
	{
		this.Paragraph   = paragraph;
		this.Range       = range;
		this.Word        = false;
		this.SpaceLen    = 0;
		this.SpacesCount = 0;

		this.Words       = 0;
		this.Spaces      = 0;
		this.Letters     = 0;
		this.SpacesSkip  = 0;
		this.LettersSkip = 0;

		this.ParaEnd   = false;
		this.LineBreak = false;
	};
	CounterRecalcState.prototype.endRange = function()
	{
	};
	/**
	 * @param {AscWord.CRunElementBase} element
	 * @param {AscWord.CRun} run
	 */
	CounterRecalcState.prototype.handleRunElement = function(element, run)
	{
		let type = element.Type;

		if (para_FieldChar === type)
		{
			if (this.isFastRecalculation())
				this.ComplexFields.processFieldChar(element);
			else
				this.ComplexFields.processFieldCharAndCollectComplexField(element);

			if (element.IsVisual())
			{
				this.Words++;
				this.Range.W += this.SpaceLen;

				if (this.Words > 1)
					this.Spaces += this.SpacesCount;
				else
					this.SpacesSkip += this.SpacesCount;

				this.Word        = false;
				this.SpacesCount = 0;
				this.SpaceLen    = 0;
				this.Range.W += element.GetWidth();
			}
			return;
		}

		if (this.ComplexFields.isHiddenFieldContent() && para_End !== type)
			return;
		
		let isHiddenCFPart = this.ComplexFields.isHiddenComplexFieldPart();
		if (isHiddenCFPart && para_End !== type && para_InstrText !== type)
			return;
		
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

				if (true !== this.Word)
				{
					this.Word = true;
					this.Words++;
				}

				this.Range.W += element.GetWidth(textPr);
				this.Range.W += this.SpaceLen;
				this.SpaceLen = 0;

				if (this.Words > 1)
					this.Spaces += this.SpacesCount;
				else
					this.SpacesSkip += this.SpacesCount;

				this.SpacesCount = 0;

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

				this.Words++;
				this.Range.W += this.SpaceLen;

				if (this.Words > 1)
					this.Spaces += this.SpacesCount;
				else
					this.SpacesSkip += this.SpacesCount;

				this.Word        = false;
				this.SpacesCount = 0;
				this.SpaceLen    = 0;
				this.Range.W += element.GetWidth();
				break;
			}
			case para_PageNum:
			case para_PageCount:
			{
				this.Words++;
				this.Range.W += this.SpaceLen;

				if (this.Words > 1)
					this.Spaces += this.SpacesCount;
				else
					this.SpacesSkip += this.SpacesCount;

				this.Word        = false;
				this.SpacesCount = 0;
				this.SpaceLen    = 0;
				this.Range.W += element.GetWidth();
				break;
			}
			case para_Tab:
			{
				this.Range.W += element.GetWidth();
				this.Range.W += this.SpaceLen;

				this.LettersSkip += this.Letters;
				this.SpacesSkip  += this.Spaces;

				this.Words   = 0;
				this.Spaces  = 0;
				this.Letters = 0;

				this.SpaceLen    = 0;
				this.SpacesCount = 0;
				this.Word        = false;
				break;
			}
			case para_NewLine:
			{
				if (true === this.Word && this.Words > 1)
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
		}
	};
	CounterRecalcState.prototype.isFastRecalculation = function()
	{
		return this.wrapState.isFastRecalculation();
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.Paragraph.CounterRecalcState = CounterRecalcState;
	
})(window);
