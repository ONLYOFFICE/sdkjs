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
	 * Class for storing the current draw state of various lines in the paragraph (underline/spelling/etc.)
	 * @param {AscWord.ParagraphDrawState} drawState
	 * @constructor
	 */
	function ParagraphLineDrawState(drawState)
	{
		AscWord.ParagraphBidiFlow.call(this);
		this.drawState = drawState;
		
		this.Paragraph = undefined;
		this.Graphics  = undefined;
		this.BgColor   = undefined;
		
		this.logicDocument = null;
		
		this.CurPos   = new AscWord.CParagraphContentPos();
		this.CurDepth = 0;
		
		this.VisitedHyperlink = false;
		this.Hyperlink = false;
		
		this.ulTrailSpace     = false;
		this.activePermRanges = false;
		
		this.Strikeout   = new CParaDrawingRangeHorizontalLines();
		this.DStrikeout  = new CParaDrawingRangeHorizontalLines();
		this.Underline   = new CParaDrawingRangeHorizontalLines();
		this.Spelling    = new CParaDrawingRangeHorizontalLines();
		this.DUnderline  = new CParaDrawingRangeHorizontalLines();
		this.RunReview   = new CParaDrawingRangeLines();
		this.CollChange  = new CParaDrawingRangeLines();
		this.FormBorder  = new CParaDrawingRangeLines();
		this.CustomLines = {};
		
		this.Page  = 0;
		this.Line  = 0;
		this.Range = 0;
		
		this.X               = 0;
		this.BaseLine        = 0;
		this.UnderlineOffset = 0;
		this.LineY0          = 0;
		this.LineY1          = 0;
		this.Spaces          = 0;
		this.SpacesSkip      = 0;
		
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
		
		this.annotationMarks = [];
		
		this.customMarks        = null;
		this.currentCustomMarks = [];
		this.customMarkInRun    = null;
		this.customMarkPos      = 0;
	}
	ParagraphLineDrawState.prototype = Object.create(AscWord.ParagraphBidiFlow.prototype);
	ParagraphLineDrawState.prototype.constructor = ParagraphLineDrawState;
	
	ParagraphLineDrawState.prototype.init = function()
	{
		this.Paragraph = this.drawState.getParagraph();
		this.Graphics  = this.drawState.getGraphics();
		this.BgColor   = this.drawState.getBgColor();
		
		this.logicDocument = this.GetLogicDocument();
		if (this.logicDocument && this.logicDocument.IsDocumentEditor())
		{
			this.ulTrailSpace     = this.logicDocument.IsUnderlineTrailSpace();
			this.activePermRanges = this.logicDocument.IsEditCommentsMode() || this.logicDocument.IsViewModeInEditor();

			let commentManager        = this.logicDocument.GetCommentsManager();
			this.isDrawComments       = commentManager && commentManager.isUse();
			this.isDrawSolvedComments = commentManager && commentManager.isUseSolved();
		}
		else
		{
			this.isDrawComments       = false;
			this.isDrawSolvedComments = false;
		}

		if (this.logicDocument && this.logicDocument.GetCustomMarks)
			this.customMarks = this.logicDocument.GetCustomMarks();
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
	ParagraphLineDrawState.prototype.resetLine = function(Line, Baseline, UnderlineOffset, lineY0, lineY1)
	{
		this.Line = Line;
		
		this.Baseline        = Baseline;
		this.UnderlineOffset = UnderlineOffset;
		this.LineY0          = lineY0;
		this.LineY1          = lineY1;
		
		this.Strikeout.Clear();
		this.DStrikeout.Clear();
		this.Underline.Clear();
		this.Spelling.Clear();
		this.RunReview.Clear();
		this.CollChange.Clear();
		this.DUnderline.Clear();
		this.FormBorder.Clear();
		
		this.CustomLines = {};
	};
	ParagraphLineDrawState.prototype.beginRange = function(rangeNum, x)
	{
		this.run = null;
		
		this.Range  = rangeNum;
		this.X      = x;
		
		let range = this.Paragraph.Lines[this.Line].Ranges[this.Range];
		this.paraLineRange = range;
		
		let isRtl = this.Paragraph.isRtlDirection();
		
		this.SpacesSkip = isRtl && !this.ulTrailSpace ? range.SpacesLeft : 0;
		this.Spaces     = range.Spaces;
		
		if (this.ulTrailSpace)
			this.Spaces += range.SpacesLeft + range.SpacesRight;
		else if (isRtl)
			this.Spaces += range.SpacesRight;
		else
			this.Spaces += range.SpacesLeft;
		
		// TODO: Numering strikeout and underline are rendering in Draw_Element, we should move it here
		if (this.Paragraph.Numbering.checkRange(this.Range, this.Line) && !isRtl)
			this.X += this.Paragraph.Numbering.WidthVisible;

		this.bidiFlow.begin(isRtl);
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
		
		this.checkCustomMarks(inRunPos);
		
		this.bidiFlow.add([para_RunBase, element, run, inRunPos, misspell, this.annotationMarks, this.currentCustomMarks.length ? this.currentCustomMarks.slice() : null], element.getBidiType());
		
		this.annotationMarks = [];
	};
	ParagraphLineDrawState.prototype.handleParaMath = function(paraMath)
	{
		if (paraMath.Root.IsEmptyRange(this.Line, this.Range))
			return;
		
		this.bidiFlow.add([para_Math, paraMath], paraMath.getBidiType());
	};
	ParagraphLineDrawState.prototype.handleMathBase = function(mathBase)
	{
		this.bidiFlow.add([para_Math_Composition, mathBase], mathBase.getBidiType());
	};
	ParagraphLineDrawState.prototype.handleBidiFlowRunElement = function(data, direction)
	{
		let element     = data[1];
		let run         = data[2];
		let inRunPos    = data[3];
		let misspell    = data[4];
		let marks       = data[5];
		let customMarks = data[6];
		
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
				if (this.SpacesSkip)
				{
					--this.SpacesSkip;
				}
				else if (this.Spaces)
				{
					if (this.paraLineRange)
					{
						startX = this.paraLineRange.CorrectX(startX);
						endX   = this.paraLineRange.CorrectX(endX);
					}					
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
		
		if (customMarks)
			this.addCustomMarkLines(startX, endX, customMarks);
		
		if (this.reviewPrColor)
			this.RunReview.Add(0, 0, startX, endX, 0, this.reviewPrColor.r, this.reviewPrColor.g, this.reviewPrColor.b, {RunPr: this.textPr});
		
		if (this.collPrChangeColor)
			this.CollChange.Add(0, 0, startX, endX, 0, this.collPrChangeColor.r, this.collPrChangeColor.g, this.collPrChangeColor.b, {RunPr : this.textPr});
		
		if (marks)
		{
			let markX = direction === AscBidi.DIRECTION.L ? startX : endX;
			for (let i = 0; i < marks.length; ++i)
			{
				marks[i].drawMark(markX, this.LineY0, this.LineY1 - this.LineY0, this.Graphics, direction === AscBidi.DIRECTION.R, this);
			}
		}
		
		this.X = endX;
	};
	ParagraphLineDrawState.prototype.handleBidiFlowParaMath = function(data, direction)
	{
		let paraMath = data[1];
		let FirstRPrp = paraMath.GetFirstRPrp();

		let Para = this.Paragraph;

		let aUnderline  = this.Underline;

		// TODO: Check bgColor
		
		let X          = this.X;
		let Y          = this.Baseline;
		let UndOff     = this.UnderlineOffset;
		let UnderlineY = Y + UndOff;
		let LineW      = (FirstRPrp.FontSize / 18) * g_dKoef_pt_to_mm;
		let BgColor    = this.BgColor;
		if ( undefined !== FirstRPrp.Shd && Asc.c_oAscShdNil !== FirstRPrp.Shd.Value )
			BgColor = FirstRPrp.Shd.Get_Color( Para );
		let AutoColor = ( undefined != BgColor && false === BgColor.Check_BlackAutoColor() ? new CDocumentColor( 255, 255, 255, false ) : new CDocumentColor( 0, 0, 0, false ) );
		let CurColor, RGBA, Theme = this.Paragraph.Get_Theme(), ColorMap = this.Paragraph.Get_ColorMap();

		if ( true === this.VisitedHyperlink && ( undefined === FirstRPrp.Color && undefined === FirstRPrp.Unifill ) )
			CurColor = new CDocumentColor( 128, 0, 151 );
		else if ( true === FirstRPrp.Color.Auto && !FirstRPrp.Unifill)
			CurColor = new CDocumentColor( AutoColor.r, AutoColor.g, AutoColor.b );
		else
		{
			if(FirstRPrp.Unifill)
			{
				FirstRPrp.Unifill.check(Theme, ColorMap);
				RGBA = FirstRPrp.Unifill.getRGBAColor();
				CurColor = new CDocumentColor( RGBA.R, RGBA.G, RGBA.B );
			}
			else
			{
				CurColor = new CDocumentColor( FirstRPrp.Color.r, FirstRPrp.Color.g, FirstRPrp.Color.b );
			}
		}

		let Bound = paraMath.Root.Get_LineBound(this.Line, this.Range),
			Width = Bound.W;

		if ( true === FirstRPrp.Underline )
		{
			aUnderline.set(UnderlineY, LineW);
			aUnderline.Add(X, X + Width, CurColor);
		}
		paraMath.Root.drawLines(this);

		this.X = Bound.X + Width;
	};
	ParagraphLineDrawState.prototype.handleBidiFlowMathBase = function(data, direction)
	{
		let mathBase = data[1];

		let CtrPrp		= mathBase.Get_CompiledCtrPrp(false);

		let aStrikeout	= this.Strikeout;
		let aDStrikeout	= this.DStrikeout;

		let ReviewType	= mathBase.GetReviewType();
		let bAddReview	= reviewtype_Add === ReviewType ? true : false;
		let bRemReview	= reviewtype_Remove === ReviewType ? true : false;
		let ReviewColor	= null;

		if (bAddReview || bRemReview)
			ReviewColor	= mathBase.ReviewInfo.Get_Color();

		let ArgSize		= mathBase.Get_CompiledArgSize();
		let fontCoeff	= MatGetKoeffArgSize(CtrPrp.FontSize, ArgSize.value);

		// calculate y-coordinate and LineW like we do it in AscWord.Run
		let X     = this.X;
		let Y     = this.Baseline - CtrPrp.FontSize * fontCoeff * g_dKoef_pt_to_mm * 0.27;
		let LineW = (CtrPrp.FontSize / 18) * g_dKoef_pt_to_mm;
		let Para  = this.Paragraph;

		// set aStrikeout && aDStrikeout
		if (true === bRemReview || true === CtrPrp.Strikeout)
			aStrikeout.set(Y, LineW);
		else if (true === CtrPrp.DStrikeout)
			aDStrikeout.set(Y, LineW);

		let BgColor = this.BgColor;
		if (CtrPrp.Shd && !CtrPrp.Shd.IsNil())
			BgColor = CtrPrp.Shd.GetSimpleColor(Para.GetTheme(), Para.GetColorMap());

		let AutoColor = ( undefined != BgColor && false === BgColor.Check_BlackAutoColor() )
			? new CDocumentColor( 255, 255, 255, false)
			: new CDocumentColor( 0, 0, 0, false );

		let CurColor,
			RGBA,
			Theme		= this.Paragraph.Get_Theme(),
			ColorMap	= this.Paragraph.Get_ColorMap();

		// stroke color
		if ( true === this.VisitedHyperlink && ( undefined === mathBase.Pr.Color && undefined === mathBase.Pr.Unifill ) )
			CurColor = new CDocumentColor( 128, 0, 151 );
		else if ( true === CtrPrp.Color.Auto && !CtrPrp.Unifill)
			CurColor = new CDocumentColor( AutoColor.r, AutoColor.g, AutoColor.b );
		else
		{
			if(CtrPrp.Unifill)
			{
				CtrPrp.Unifill.check(Theme, ColorMap);
				RGBA = CtrPrp.Unifill.getRGBAColor();
				CurColor = new CDocumentColor( RGBA.R, RGBA.G, RGBA.B );
			}
			else
			{
				CurColor = new CDocumentColor( CtrPrp.Color.r, CtrPrp.Color.g, CtrPrp.Color.b );
			}
		}

		let CurLine		= this.Line - mathBase.StartLine;
		let CurRange	= ( 0 === CurLine ? this.Range - mathBase.StartRange : this.Range );
		let Bound		= mathBase.Bounds.Get_LineBound(CurLine, CurRange);

		if ( true === bRemReview )
			aStrikeout.Add( X, X + Bound.W, CurColor );
		else if ( true === CtrPrp.DStrikeout )
			aDStrikeout.Add( X, X + Bound.W, CurColor );
		else if ( true === CtrPrp.Strikeout )
			aStrikeout.Add( X, X + Bound.W, CurColor );

		mathBase.drawLinesForContent(this);

		this.X = Bound.X + Bound.W;
	};
	ParagraphLineDrawState.prototype.handleAnnotationMark = function(mark)
	{
		this.annotationMarks.push(mark);
	};
	ParagraphLineDrawState.prototype.handleCommentMark = function(paraComment)
	{
		if (!this.isDrawComments)
			return;
		
		let commentId = paraComment.GetCommentId();
		let docComments = this.logicDocument.GetComments();
		let comment = docComments.Get_ById(commentId);
		if (!comment
			|| (!this.isDrawSolvedComments && comment.IsSolved())
			|| !AscCommon.UserInfoParser.canViewComment(comment.GetUserName()))
			return;
		
		this.handleAnnotationMark(paraComment);
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
	ParagraphLineDrawState.prototype.getStartedCustomMarks = function()
	{
		if (!this.customMarks)
			return [];
		
		return this.customMarks.getStartedMarks(this.Paragraph, this.CurPos);
	};
	ParagraphLineDrawState.prototype.initCustomMarks = function(run, startPos)
	{
		if (!this.customMarks)
			return;
		
		this.currentCustomMarks = this.customMarks.getStartedMarks(this.Paragraph, this.CurPos);
		this.customMarkInRun = this.customMarks.flatRunMarks(run.GetId());
		this.customMarkPos = 0;
		
		if (this.customMarkInRun)
		{
			while (this.customMarkPos < this.customMarkInRun.length && startPos > this.customMarkInRun[this.customMarkPos].getPos())
			{
				++this.customMarkPos;
			}
		}
	};
	ParagraphLineDrawState.prototype.checkCustomMarks = function(pos)
	{
		if (!this.customMarkInRun)
			return this.currentCustomMarks;
		
		while (this.customMarkPos < this.customMarkInRun.length && pos >= this.customMarkInRun[this.customMarkPos].getPos())
		{
			let mark = this.customMarkInRun[this.customMarkPos];
			if (mark.isStart())
			{
				this.currentCustomMarks.push(mark);
			}
			else
			{
				let handlerId = mark.getHandlerId();
				let rangeId   = mark.getRangeId();
				for (let i = 0; i < this.currentCustomMarks.length; ++i)
				{
					let _mark = this.currentCustomMarks[i];
					if (handlerId === _mark.getHandlerId() && rangeId === _mark.getRangeId())
					{
						if (i === this.currentCustomMarks)
							this.currentCustomMarks.length--;
						else
							this.currentCustomMarks.splice(i, 1);
						break;
					}
				}
			}
			++this.customMarkPos;
		}
	};
	ParagraphLineDrawState.prototype.GetLogicDocument = function()
	{
		return this.Paragraph.GetLogicDocument();
	};
	ParagraphLineDrawState.prototype.isActivePermRanges = function()
	{
		return this.activePermRanges;
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

			if (!form.IsMainForm() && form.GetMainForm() && !form.GetMainForm().IsLabeledCheckBox())
				form = form.GetMainForm();
			
			if (form.IsFormRequired() && !form.IsLabeledCheckBox() && form.CheckOFormUserMaster() && this.logicDocument.IsHighlightRequiredFields() && !this.Graphics.isPrintMode)
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
	ParagraphLineDrawState.prototype.addCustomMarkLines = function(startX, endX, customMarks)
	{
		for (let i = 0; i < customMarks.length; ++i)
		{
			let mark      = customMarks[i];
			let handlerId = mark.getHandlerId();
			let rangeId   = mark.getRangeId();
			
			if (!this.CustomLines[handlerId])
				this.CustomLines[handlerId] = {};
			
			if (!this.CustomLines[handlerId][rangeId])
				this.CustomLines[handlerId][rangeId] = new CParaDrawingRangeHorizontalLines();
			
			this.CustomLines[handlerId][rangeId].Add(startX, endX, {r : 0, g : 0, b : 0});
		}
	};
	ParagraphLineDrawState.prototype.drawCustomMarks = function(graphics, y, h, baseLineY)
	{
		for (let handlerId in this.CustomLines)
		{
			for (let rangeId in this.CustomLines[handlerId])
			{
				let lines   = this.CustomLines[handlerId][rangeId];
				let element = lines.getNext();
				while (element)
				{
					graphics.drawCustomRange(handlerId, rangeId, element.x0, y, element.x1 - element.x0, h, baseLineY);
					element = lines.getNext();
				}
			}
		}
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.ParagraphLineDrawState = ParagraphLineDrawState;
	
})(window);
