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
	 * Class for storing the current draw state of various lines in the paragraph (underline/spelling/etc.)
	 * @param {AscWord.ParagraphDrawState} drawState
	 * @constructor
	 */
	function ParagraphLineDrawState(drawState)
	{
		this.drawState = drawState;
		
		this.Paragraph = undefined;
		this.Graphics  = undefined;
		
		this.logicDocument = null;
		
		this.CurPos   = new AscWord.CParagraphContentPos();
		this.CurDepth = 0;
		
		this.VisitedHyperlink = false;
		this.Hyperlink = false;
		
		this.ulTrailSpace = false;
		
		this.Strikeout  = new CParaDrawingRangeHorizontalLines();
		this.DStrikeout = new CParaDrawingRangeHorizontalLines();
		this.Underline  = new CParaDrawingRangeHorizontalLines();
		this.Spelling   = new CParaDrawingRangeHorizontalLines();
		this.DUnderline = new CParaDrawingRangeHorizontalLines();
		this.RunReview  = new CParaDrawingRangeLines();
		this.CollChange = new CParaDrawingRangeLines();
		this.FormBorder = new CParaDrawingRangeLines();
		
		this.Page  = 0;
		this.Line  = 0;
		this.Range = 0;
		
		this.X               = 0;
		this.BaseLine        = 0;
		this.UnderlineOffset = 0;
		this.Spaces          = 0;
		
		this.complexFields = new AscWord.ParagraphComplexFieldStack();
		this.contentControls = [];
		
		this.run = null;
		
		this.isFormPlaceholder = false;
		
		this.yOffset = 0;
		this.color   = null;
		
		this.paraLineRange = null;
		
		this.isUnderline  = false;
		this.isStrikeout  = false;
		this.isDStrikeout = false;
		
		this.form            = null;
		this.formBorderColor = null;
		this.formBorderWidth = 0;
		this.combMax         = -1;
		
		this.reviewAdd    = false;
		this.reviewRem    = false;
		this.reviewRemAdd = false;
		this.reviewMove   = false;
		
		this.reviewColor       = REVIEW_COLOR;
		this.reviewRemAddColor = REVIEW_COLOR;
		this.reviewPrColor     = null;
		
		this.collPrChangeColor = null;
		
		this.bidiFlow = new AscWord.BidiFlow(this);
	}
	ParagraphLineDrawState.prototype.init = function()
	{
		this.Paragraph = this.drawState.getParagraph();
		this.Graphics  = this.drawState.getGraphics();
		
		this.logicDocument = this.GetLogicDocument();
		if (this.logicDocument && this.logicDocument.IsDocumentEditor())
			this.ulTrailSpace = this.logicDocument.IsUnderlineTrailSpace();
	};
	ParagraphLineDrawState.prototype.resetPage = function(page)
	{
		this.Page = page;
		
		this.VisitedHyperlink = false;
		this.Hyperlink        = false;
		
		this.CurPos   = new AscWord.CParagraphContentPos();
		this.CurDepth = 0;
		
		this.complexFields.resetPage(this.Paragraph, page);
	};
	ParagraphLineDrawState.prototype.resetLine = function(Line, Baseline, UnderlineOffset)
	{
		this.Line = Line;
		
		this.Baseline        = Baseline;
		this.UnderlineOffset = UnderlineOffset;
		
		this.Strikeout.Clear();
		this.DStrikeout.Clear();
		this.Underline.Clear();
		this.Spelling.Clear();
		this.RunReview.Clear();
		this.CollChange.Clear();
		this.DUnderline.Clear();
		this.FormBorder.Clear();
	};
	ParagraphLineDrawState.prototype.beginRange = function(range, x, spaces)
	{
		this.run = null;
		
		this.Range  = range;
		this.X      = x;
		this.Spaces = spaces;
		
		this.bidiFlow.begin(this.Paragraph.isRtlDirection());
		
		this.paraLineRange = this.Paragraph.Lines[this.Line].Ranges[this.Range];
	};
	ParagraphLineDrawState.prototype.endRange = function()
	{
		this.bidiFlow.end();
	};
	/**
	 * @param element {AscWord.CRunElementBase}
	 * @param run {AscWord.CRun}
	 * @param inRunPos {number}
	 * @param misspell {boolean}
	 */
	ParagraphLineDrawState.prototype.handleRunElement = function(element, run, inRunPos, misspell)
	{
		if (!this.complexFields.checkRunElement(element))
			return;
		
		if (para_Drawing === element.Type && !element.IsInline())
			return;
		
		this.bidiFlow.add([element, run, inRunPos, misspell], element.getBidiType());
	};
	ParagraphLineDrawState.prototype.handleBidiFlow = function(data)
	{
		let element  = data[0];
		let run      = data[1];
		let inRunPos = data[2];
		let misspell = data[3];
		
		this.handleRun(run);
		
		this.handleFormBorder(element, run, inRunPos);
		
		if (this.isFormPlaceholder)
		{
			this.X += element.GetWidthVisible();
			return;
		}
		
		this.addCompositeInputLine(element, run, inRunPos);
		let oTextDrawerInfo;
		if (this.Graphics.m_bIsTextDrawer) {
			oTextDrawerInfo = {TextDrawer: {SplitType: this.Graphics.m_nCurrentSplitOptions, TextElement: element}};
		}
		let startX = this.X;
		let endX   = this.X + element.GetWidthVisible();
		switch (element.Type)
		{
			case para_PageNum:
			case para_PageCount:
			case para_Tab:
			case para_Text:
			case para_Sym:
			case para_FootnoteReference:
			case para_FootnoteRef:
			case para_EndnoteReference:
			case para_EndnoteRef:
			case para_Separator:
			case para_ContinuationSeparator:
			case para_Math_Text:
			case para_Math_BreakOperator:
			case para_Math_Ampersand:
			case para_Math_Placeholder:
				this.addLines(startX, endX, undefined, oTextDrawerInfo);
				break;
			case para_Space:
				if (this.paraLineRange)
				{
					startX = this.paraLineRange.CorrectX(startX);
					endX   = this.paraLineRange.CorrectX(endX);
				}
				
				if (this.Spaces > 0 || this.ulTrailSpace)
				{
					--this.Spaces;
					this.addLines(startX, endX, undefined, oTextDrawerInfo);
				}
				break;
			case para_Drawing:
				if (element.IsInline())
					this.addLines(startX, endX, false, oTextDrawerInfo);
				break;
			case para_End:
				this.isUnderline  = false;
				this.isStrikeout  = false;
				this.isDStrikeout = false;
				this.addLines(startX, endX, undefined, oTextDrawerInfo);
				break;
			case para_FieldChar:
				if (element.IsVisual())
					this.addLines(startX, endX, undefined, oTextDrawerInfo);
				break;
		}
		
		if (misspell)
			this.Spelling.Add(startX, endX, AscWord.BLACK_COLOR);
		
		if (this.reviewPrColor)
			this.RunReview.Add(0, 0, startX, endX, 0, this.reviewPrColor.r, this.reviewPrColor.g, this.reviewPrColor.b, {RunPr: this.textPr});
		
		if (this.collPrChangeColor)
			this.CollChange.Add(0, 0, startX, endX, 0, this.collPrChangeColor.r, this.collPrChangeColor.g, this.collPrChangeColor.b, {RunPr : this.textPr});
		
		this.X = endX;
	};
	/**
	 * Получаем количество орфографических ошибок в данном месте
	 * @returns {number}
	 */
	ParagraphLineDrawState.prototype.GetSpellingErrorsCounter = function()
	{
		var nCounter = 0;
		var oSpellChecker = this.Paragraph.GetSpellChecker();
		for (var nIndex = 0, nCount = oSpellChecker.GetElementsCount(); nIndex < nCount; ++nIndex)
		{
			var oSpellElement = oSpellChecker.GetElement(nIndex);
			
			if (false !== oSpellElement.Checked || oSpellElement.CurPos)
				continue;
			
			var oStartPos = oSpellElement.GetStartPos();
			var oEndPos   = oSpellElement.GetEndPos();
			
			if (this.CurPos.Compare(oStartPos) > 0 && this.CurPos.Compare(oEndPos) < 0)
				nCounter++;
		}
		
		return nCounter;
	};
	ParagraphLineDrawState.prototype.GetLogicDocument = function()
	{
		return this.Paragraph.GetLogicDocument();
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	/**
	 * @param run {AscWord.CRun}
	 */
	ParagraphLineDrawState.prototype.handleRun = function(run)
	{
		if (run === this.run)
			return;
		
		this.run = run;
		
		this.isFormPlaceholder = false;
		
		let form       = run.GetParentForm();
		let formBorder = null;
		let combMax    = -1;
		if (form)
		{
			if (form.IsTextForm() && form.GetTextFormPr().IsComb())
				combMax = form.GetTextFormPr().GetMaxCharacters();

			if (!form.IsMainForm())
				form = form.GetMainForm();
			
			if (form.IsFormRequired() && form.CheckOFormUserMaster() && this.logicDocument.IsHighlightRequiredFields() && !this.Graphics.isPrintMode)
				formBorder = this.logicDocument.GetRequiredFieldsBorder();
			else if (form.GetFormPr().GetBorder())
				formBorder = form.GetFormPr().GetBorder();
			
			this.isFormPlaceholder = (form.IsPlaceHolder() && this.Graphics.isPrintMode);
			
			let isForceDrawPlaceHolders = this.logicDocument.ForceDrawPlaceHolders;
			if (true === isForceDrawPlaceHolders)
				this.isFormPlaceholder = false;
			else if (false === isForceDrawPlaceHolders && form.IsPlaceHolder())
				this.isFormPlaceholder = true;
		}
		
		this.form    = form;
		this.combMax = combMax;
		
		this.formBorderWidth = 0;
		this.formBorderColor = null;
		
		if (this.form && formBorder)
		{
			this.formBorderWidth = formBorder.GetWidth();
			this.formBorderColor = AscWord.CDocumentColorA.fromObjectRgb(formBorder.GetColor());
		}
		else
		{
			// TODO: It's better to handle CC stack on the fly instead of getting it for each run
			let contentControls = run.GetParentContentControls();
			for (let i = contentControls.length - 1; i >= 0; --i)
			{
				let cc = contentControls[i];
				if (cc instanceof AscWord.CInlineLevelSdt && cc.getBorderColor())
				{
					this.form            = cc;
					this.formBorderWidth = 0;
					this.formBorderColor = cc.getBorderColor();
					break;
				}
			}
		}
		
		this.yOffset = run.getYOffset();
		
		let textPr = run.getCompiledPr();
		
		this.updateStrikeoutUnderlinePos(run, textPr.FontSize, textPr.VertAlign);
		this.updateColor(textPr);
		this.updateReviewState(run);
		
		this.isUnderline  = textPr.Underline;
		this.isStrikeout  = textPr.Strikeout;
		this.isDStrikeout = textPr.DStrikeout;
		
		if (run.IsMathRun() && run.IsPlaceholder())
		{
			this.isUnderline = false;
			
			let ctrPrp = run.Parent.GetCtrPrp();
			this.isStrikeout  = ctrPrp.Strikeout;
			this.isDStrikeout = ctrPrp.DStrikeout;
		}
		
		this.collPrChangeColor = run.getCollPrChangeColor();
		
		this.textPr = textPr;
	};
	ParagraphLineDrawState.prototype.handleFormBorder = function(item, run, inRunPos)
	{
		let itemWidth = item.GetWidthVisible();
		if (!this.form || itemWidth <= 0.001 || !this.formBorderColor)
			return;
		
		let borderW     = this.formBorderWidth;
		let borderColor = this.formBorderColor;
		
		let Y = this.Baseline;
		let X = this.X;
		
		let formBounds = this.form.GetRangeBounds(this.Line, this.Range);
		let additional = {
			Form    : this.form,
			Comb    : this.combMax,
			Y       : formBounds.Y,
			H       : formBounds.H,
			BorderL : 0 === inRunPos
				|| item.IsSpace()
				|| (item.IsText() && !item.IsCombiningMark()),
			BorderR : run.Content.length - 1 === inRunPos
				|| item.IsSpace()
				|| (item.IsText()
					&& inRunPos < run.Content.length - 1
					&& (run.Content[inRunPos + 1].IsSpace()
						|| (run.Content[inRunPos + 1].IsText() && !run.Content[inRunPos + 1].IsCombiningMark())))
		};
		
		if (item.RGapCount)
		{
			var nGapEnd = X + itemWidth;
			this.FormBorder.addWithAlpha(Y, Y, X, nGapEnd - item.RGapCount * item.RGapShift,
				borderW,
				borderColor.r,
				borderColor.g,
				borderColor.b,
				borderColor.a,
				additional
			);
			
			for (var nGapIndex = 0; nGapIndex < item.RGapCount; ++nGapIndex)
			{
				this.FormBorder.addWithAlpha(Y, Y, nGapEnd - (item.RGapCount - nGapIndex) * item.RGapShift, nGapEnd - (item.RGapCount - nGapIndex - 1) * item.RGapShift,
					borderW,
					borderColor.r,
					borderColor.g,
					borderColor.b,
					borderColor.a,
					additional
				);
			}
		}
		else
		{
			this.FormBorder.addWithAlpha(Y, Y, X, X + itemWidth,
				borderW,
				borderColor.r,
				borderColor.g,
				borderColor.b,
				borderColor.a,
				additional
			);
		}
	};
	ParagraphLineDrawState.prototype.updateStrikeoutUnderlinePos = function(run, fontSize, vertAlign)
	{
		let fontSizeMM = fontSize * g_dKoef_pt_to_mm;
		if (run.IsMathRun())
			fontSizeMM *= MatGetKoeffArgSize(fontSize, run.Parent.Compiled_ArgSz.value);
		
		let strikeoutShift = 0.27;
		if (AscCommon.vertalign_SubScript === vertAlign)
			strikeoutShift = AscCommon.vaKSize * 0.27 + AscCommon.vaKSub;
		else if (AscCommon.vertalign_SuperScript === vertAlign)
			strikeoutShift = AscCommon.vaKSize * 0.27 + AscCommon.vaKSuper;
		
		let strikeoutY = this.Baseline - this.yOffset - fontSizeMM * strikeoutShift;
		let underlineY = this.Baseline - this.yOffset + this.UnderlineOffset;
		
		if (AscCommon.vertalign_SubScript === vertAlign)
			underlineY -= AscCommon.vaKSub * fontSizeMM;
		
		let lineW = (fontSize / 18) * g_dKoef_pt_to_mm;
		
		this.Strikeout.set(strikeoutY, lineW);
		this.DStrikeout.set(strikeoutY, lineW);
		this.Underline.set(underlineY, lineW);
		this.DUnderline.set(underlineY, lineW);
		this.Spelling.set(underlineY, lineW);
	};
	ParagraphLineDrawState.prototype.updateColor = function(textPr)
	{
		if (this.VisitedHyperlink)
		{
			AscFormat.G_O_VISITED_HLINK_COLOR.check(this.Paragraph.getTheme(), this.Paragraph.getColorMap());
			let RGBA = AscFormat.G_O_VISITED_HLINK_COLOR.getRGBAColor();
			this.color = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B, RGBA.A);
		}
		else if (textPr.Color.IsAuto() && !textPr.Unifill)
		{
			if (textPr.FontRef && textPr.FontRef.Color)
			{
				textPr.FontRef.Color.check(this.Paragraph.getTheme(), this.Paragraph.getColorMap());
				let RGBA   = textPr.FontRef.Color.RGBA;
				this.color = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B, RGBA.A);
			}
			else
			{
				let bgColor = this.drawState.getBgColor();
				if (textPr.Shd && !textPr.Shd.IsNil())
					bgColor = textPr.Shd.GetSimpleColor(this.Paragraph.getTheme(), this.Paragraph.getColorMap());
				
				this.color = bgColor && !bgColor.isBlackAutoColor() ? AscWord.WHITE_COLOR : AscWord.BLACK_COLOR;
			}
		}
		else
		{
			if (this.isSlideEditor() && this.Hyperlink)
			{
				AscFormat.G_O_HLINK_COLOR.check(this.Paragraph.getTheme(), this.Paragraph.getColorMap());
				let RGBA   = AscFormat.G_O_HLINK_COLOR.getRGBAColor();
				this.color = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B, RGBA.A);
			}
			else if (textPr.Unifill)
			{
				textPr.Unifill.check(this.Paragraph.getTheme(), this.Paragraph.getColorMap());
				let RGBA   = textPr.Unifill.getRGBAColor();
				this.color = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B);
			}
			else
			{
				this.color = textPr.Color;
			}
		}
	};
	ParagraphLineDrawState.prototype.updateReviewState = function(run)
	{
		this.reviewAdd    = false;
		this.reviewRem    = false;
		this.reviewRemAdd = false;
		this.reviewMove   = false;
		
		this.reviewColor       = REVIEW_COLOR;
		this.reviewRemAddColor = REVIEW_COLOR;
		this.reviewPrColor     = run.Pr.HavePrChange() && !run.IsMathRun() ? run.GetPrReviewColor() : null;
		
		let reviewType = run.GetReviewType();
		if (reviewType !== reviewtype_Common)
		{
			this.reviewAdd   = reviewtype_Add === reviewType;
			this.reviewRem   = reviewtype_Remove === reviewType;
			this.reviewColor = run.GetReviewColor();
			this.reviewMove  = reviewtype_Add === reviewType ? run.GetReviewInfo().IsMovedTo() : run.GetReviewInfo().IsMovedFrom();
			
			let prevInfo = run.GetReviewInfo().GetPrevAdded();
			if (prevInfo)
			{
				this.reviewRemAdd = true;
				this.reviewRemAddColor = prevInfo.GetColor();
			}
		}
	};
	ParagraphLineDrawState.prototype.isSlideEditor = function()
	{
		return this.Paragraph && !this.Paragraph.bFromDocument;
	};
	/**
	 * @param startX {number}
	 * @param endX {number}
	 * @param drawStrikeout {boolean}
	 */
	ParagraphLineDrawState.prototype.addLines = function(startX, endX, drawStrikeout, oTextDrawerInfo)
	{
		if (endX - startX < 0.001)
			return;
		
		if (false !== drawStrikeout)
		{
			if (this.reviewRem)
			{
				if (this.reviewMove)
					this.DStrikeout.Add(startX, endX, this.reviewColor, oTextDrawerInfo);
				else
					this.Strikeout.Add(startX, endX, this.reviewColor, oTextDrawerInfo);

				if (this.reviewRemAdd)
					this.Underline.Add(startX, endX, this.reviewRemAddColor, oTextDrawerInfo);
			}
			else if (this.isDStrikeout)
			{
				this.DStrikeout.Add(startX, endX, this.color, oTextDrawerInfo, this.textPr);
			}
			else if (this.isStrikeout)
			{
				this.Strikeout.Add(startX, endX, this.color, oTextDrawerInfo, this.textPr);
			}
		}

		if (this.reviewAdd)
		{
			if (this.reviewMove)
				this.DUnderline.Add(startX, endX, this.reviewColor, oTextDrawerInfo);
			else
				this.Underline.Add(startX, endX, this.reviewColor, oTextDrawerInfo);
		}
		else if (this.isUnderline)
		{
			this.Underline.Add(startX, endX, this.color, oTextDrawerInfo, this.textPr);
		}
	};
	ParagraphLineDrawState.prototype.addCompositeInputLine = function(element, run, inRunPos)
	{
		if (para_Text !== element.Type || !run.CompositeInput || !run.CompositeInput.isInside(inRunPos))
			return;
		
		this.Underline.Add(this.X, this.X + element.GetWidthVisible(), this.color, undefined, this.textPr);
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.ParagraphLineDrawState = ParagraphLineDrawState;
	
})(window);
