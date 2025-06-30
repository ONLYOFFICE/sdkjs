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
	 * Class for storing the current draw state of paragraph content (text, images)
	 * @param {AscWord.ParagraphDrawState} drawState - reference to the main state
	 * @constructor
	 */
	function ParagraphContentDrawState(drawState)
	{
		this.drawState = drawState;
		
		this.Paragraph = undefined;
		this.Graphics  = undefined;
		this.BgColor   = undefined;
		
		this.Theme     = undefined;
		this.ColorMap  = undefined;
		
		this.CurPos = new AscWord.CParagraphContentPos();
		
		this.VisitedHyperlink = false;
		this.Hyperlink = false;
		
		this.Page   = 0;
		this.Line   = 0;
		this.Range  = 0;
		
		this.X = 0;
		this.Y = 0;
		
		this.LineTop    = 0;
		this.LineBottom = 0;
		this.BaseLine   = 0;
		
		this.complexFields = new AscWord.ParagraphComplexFieldStack();
		
		this.run            = null;
		this.yOffset        = 0;
		this.textPr         = null;
		this.reviewColor    = null;
		this.themeColor     = null;
		this.autoColor      = null;
		this.calcY          = 0; // calculated vertical position taking into account sub/super script
		
		// for Math
		this.mathTextInfo = null;
		this.paraMath     = null;
		
		this.bidiFlow = new AscWord.BidiFlow(this);
	}
	ParagraphContentDrawState.prototype.init = function()
	{
		this.Paragraph = this.drawState.getParagraph();
		this.Graphics  = this.drawState.getGraphics();
		
		this.BgColor  = this.drawState.getBgColor();
		this.Theme    = this.drawState.getTheme();
		this.ColorMap = this.drawState.getColorMap();
		
		this.VisitedHyperlink = false;
		this.Hyperlink        = false;
		
		this.CurPos = new AscWord.CParagraphContentPos();
	};
	ParagraphContentDrawState.prototype.resetPage = function(page)
	{
		this.Page = page;
		this.complexFields.resetPage(this.Paragraph, page);
	};
	ParagraphContentDrawState.prototype.resetLine = function(line, baseLine, lineTop, lineBottom)
	{
		this.Line = line;
		
		this.LineTop    = lineTop;
		this.LineBottom = lineBottom;
		this.BaseLine   = baseLine;
		
		this.Y = baseLine;
	};
	ParagraphContentDrawState.prototype.beginRange = function(range, x)
	{
		this.run = null;
		
		this.Range = range;
		
		this.X = x;
		
		let isRtl = this.Paragraph.isRtlDirection();
		if (this.Paragraph.Numbering.checkRange(this.Range, this.Line) && !isRtl)
			this.handleNumbering();
		
		this.bidiFlow.begin(this.Paragraph.isRtlDirection());
	};
	ParagraphContentDrawState.prototype.endRange = function()
	{
		this.bidiFlow.end();
		
		let isRtl = this.Paragraph.isRtlDirection();
		if (this.Paragraph.Numbering.checkRange(this.Range, this.Line) && isRtl)
			this.handleNumbering();
	};
	/**
	 * @param element {AscWord.CRunElementBase}
	 * @param run {AscWord.CRun}
	 */
	ParagraphContentDrawState.prototype.handleRunElement = function(element, run)
	{
		if (!this.complexFields.checkRunElement(element))
			return;
		
		this.bidiFlow.add([element, run], element.getBidiType());
	};
	ParagraphContentDrawState.prototype.handleBidiFlow = function(data, direction)
	{
		let element = data[0];
		this.handleRun(data[1]);
		
		switch (element.Type)
		{
			case para_Text:
				this.handleText(element, direction);
				break;
			case para_Drawing:
				this.handleDrawing(element);
				break;
			case para_End:
				this.handleParagraphMark(element);
				break;
			case para_Math_Ampersand:
			case para_Math_Text:
			case para_Math_BreakOperator:
			case para_Math_Placeholder:
				this.handleMathElement(element);
				break;
			case para_FieldChar:
				this.handleFieldChar(element);
				break;
			default:
				this.handleRegularElement(element, direction);
				break;
		}
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	/**
	 * @param run {AscWord.CRun}
	 */
	ParagraphContentDrawState.prototype.handleRun = function(run)
	{
		if (this.run === run)
			return;
		
		this.run = run;
		
		this.yOffset = run.getYOffset();
		
		let textPr = run.getCompiledPr();
		
		if (run.IsMathRun())
		{
			this.mathTextInfo = new CMathInfoTextPr({
				TextPr      : textPr,
				ArgSize     : run.Parent.Compiled_ArgSz.value,
				bNormalText : run.IsNormalText(),
				bEqArray    : run.bEqArray
			});
			this.paraMath = run.ParaMath;
		}
		else
		{
			this.mathTextInfo = null;
			this.paraMath     = null;
		}
		
		this.reviewColor = null;
		this.themeColor  = null;
		this.autoColor   = null;
		
		if (reviewtype_Common !== run.GetReviewType())
		{
			this.reviewColor = run.GetReviewColor();
		}
		else if (this.VisitedHyperlink)
		{
			AscFormat.G_O_VISITED_HLINK_COLOR.check(this.Theme, this.ColorMap);
			let RGBA = AscFormat.G_O_VISITED_HLINK_COLOR.getRGBAColor();
			this.themeColor = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B, false);
		}
		else if (textPr.Unifill)
		{
			let RGBA = null;
			if (this.isSlideEditor() && this.Hyperlink)
			{
				AscFormat.G_O_HLINK_COLOR.check(this.Theme, this.ColorMap);
				RGBA = AscFormat.G_O_HLINK_COLOR.getRGBAColor();
			}
			else if (!this.isTextArtDraw())
			{
				textPr.Unifill.check(this.Theme, this.ColorMap);
				RGBA = textPr.Unifill.getRGBAColor();
			}
			
			if (RGBA)
				this.themeColor = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B, false);
		}
		
		if (textPr.FontRef && textPr.FontRef.Color)
		{
			textPr.FontRef.Color.check(this.Theme, this.ColorMap);
			let RGBA = textPr.FontRef.Color.RGBA;
			this.autoColor = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B, false);
		}
		else
		{
			let bgColor = this.BgColor;
			if (textPr.Shd && !textPr.Shd.IsNil())
				bgColor = textPr.Shd.GetSimpleColor(this.Theme, this.ColorMap);
			
			this.autoColor = (bgColor && !bgColor.Check_BlackAutoColor() ? new CDocumentColor(255, 255, 255, false) : new CDocumentColor(0, 0, 0, false));
		}
		
		this.updateGraphicsState(textPr, run.IsUseAscFont(textPr));
		
		this.calcY = this.calculateY(textPr);
	};
	/**
	 * @param textPr {AscWord.CTextPr}
	 * @param [useAscFont=false] {boolean}
	 */
	ParagraphContentDrawState.prototype.updateGraphicsState = function(textPr, useAscFont)
	{
		if (textPr)
		{
			this.textPr = textPr;
			
			if (useAscFont)
			{
				textPr = textPr.Copy();
				textPr.RFonts.SetAll("ASCW3", -1);
			}
			
			this.Graphics.SetTextPr(textPr, this.Theme);
		}
		
		let color = this.getTextColor();
		if (color)
		{
			this.Graphics.b_color1(color.r, color.g, color.b, 255);
			this.Graphics.p_color(color.r, color.g, color.b, 255);
		}
	};
	ParagraphContentDrawState.prototype.getTextColor = function()
	{
		if (this.reviewColor)
			return this.reviewColor;
		if (this.themeColor)
			return this.themeColor;
		
		if (this.isTextArtDraw())
		{
			if (this.textPr.Color.IsAuto() && !this.textPr.TextFill)
				return this.autoColor;
			return null;
		}
		
		if (this.textPr.Color.IsAuto())
			return this.autoColor;
		
		return this.textPr.Color;
	};
	/**
	 * @param text {AscWord.CRunText}
	 * @param direction {AscBidi.DIRECTION}
	 */
	ParagraphContentDrawState.prototype.handleText = function(text, direction)
	{
		if (AscBidi.DIRECTION.R === direction && AscBidi.isPairedBracket(text.GetCodePoint()))
			text.Draw(this.X, this.calcY - this.yOffset, this.Graphics, this, this.textPr, AscBidi.getPairedBracketGrapheme(text.Grapheme));
		else
			text.Draw(this.X, this.calcY - this.yOffset, this.Graphics, this, this.textPr);
		
		this.X += text.GetWidthVisible();
	};
	/**
	 * @param drawing {ParaDrawing}
	 */
	ParagraphContentDrawState.prototype.handleDrawing = function(drawing)
	{
		if (!drawing.IsInline())
			return;
		
		drawing.Draw(this.X, this.calcY - this.yOffset, this.Graphics, this, this.textPr);
		this.X += drawing.GetWidthVisible();
		
		// Внутри отрисовки инлайн-автофигур могут изменится цвета и шрифт, поэтому восстанавливаем настройки
		this.updateGraphicsState();
	};
	/**
	 * @param element {AscWord.CRunElementBase}
	 */
	ParagraphContentDrawState.prototype.handleRegularElement = function(element)
	{
		element.Draw(this.X, this.calcY - this.yOffset, this.Graphics, this, this.textPr);
		this.X += element.GetWidthVisible();
		
		if (element.IsTab())
			this.updateGraphicsState();
	};
	/**
	 * @param paraMark {AscWord.CRunParagraphMark}
	 */
	ParagraphContentDrawState.prototype.handleParagraphMark = function(paraMark)
	{
		if (this.Paragraph.IsInFixedForm())
			return;
		
		let sectPr = this.Paragraph.Get_SectionPr();
		let logicDocument = this.Paragraph.GetLogicDocument();
		if (!logicDocument)
			return;
		
		if (logicDocument !== this.Paragraph.GetParent())
			sectPr = undefined;
		
		let editor = logicDocument.GetApi();
		if ((!editor || !editor.ShowParaMarks) && (sectPr || !this.reviewColor))
		{
			this.X += paraMark.GetWidthVisible();
			return;
		}
		
		let y = this.Y;
		if (!sectPr)
		{
			let endTextPr = true !== this.Graphics.m_bIsTextDrawer ? this.Paragraph.GetParaEndCompiledPr() : this.textPr;
			this.updateGraphicsState(endTextPr);
			y = this.calculateY(endTextPr);
		}
		
		paraMark.Draw(this.X, y - this.yOffset, this.Graphics);
		this.X += paraMark.GetWidthVisible();
	};
	/**
	 * @param element {CMathText | CMathAmp}
	 */
	ParagraphContentDrawState.prototype.handleMathElement = function(element)
	{
		if (para_Math_Placeholder === element.Type && this.Graphics.isPdf())
			return;
		
		let linePos = this.paraMath.GetLinePosition(this.Line, this.Range);
		element.Draw(linePos.x, linePos.y, this.Graphics, this.mathTextInfo);
		this.X += element.GetWidthVisible();
	};
	/**
	 * @param fieldChar {ParaFieldChar}
	 */
	ParagraphContentDrawState.prototype.handleFieldChar = function(fieldChar)
	{
		if (!fieldChar.IsVisual())
			return;
		
		// Draw the auto-calculated pageNum in the header/footer
		let complexField = fieldChar.GetComplexField();
		
		// We can't use textPr of the current run, since the properties of the actual text
		// may be different
		let textPr = complexField.GetFieldValueTextPr(true);
		this.updateGraphicsState(textPr);
		
		this.handleRegularElement(fieldChar);
	};
	ParagraphContentDrawState.prototype.handleNumbering = function()
	{
		let graphics = this.Graphics;
		let BgColor  = this.drawState.getBgColor();
		let Pr       = this.drawState.getParagraphCompiledPr();
		let isRtl    = Pr.ParaPr.Bidi;
		
		let X = this.X;
		let Y = this.Y;
		let CurLine = this.Line;
		let CurPage = this.Page;
		
		let para = this.Paragraph;
		
		var nReviewType  = para.GetReviewType();
		var oReviewColor = para.GetReviewColor();
		
		let numItem = para.Numbering;
		if (para_Numbering === numItem.Type)
		{
			var isHavePrChange = para.HavePrChange();
			
			let prevNumPr = para.GetPrChangeNumPr();
			let numPr     = Pr.ParaPr.NumPr;
			
			if (prevNumPr && prevNumPr.IsValid())
			{
				prevNumPr = prevNumPr.Copy();
				if (undefined === prevNumPr.Lvl)
					prevNumPr.Lvl = 0;
			}
			else
			{
				prevNumPr = null;
			}
			
			var isHaveNumbering = false;
			if ((undefined === para.Get_SectionPr() || true !== para.IsEmpty())
				&& (!para.Parent || !para.Parent.IsEmptyParagraphAfterTableInTableCell(para.GetIndex()))
				&& ((numPr && numPr.IsValid()) || (prevNumPr && prevNumPr.IsValid())))
			{
				isHaveNumbering = true;
			}
			
			if (!isHaveNumbering || (!numPr && !prevNumPr))
			{
				// Ничего не делаем
			}
			else
			{
				var oNumbering = para.Parent.GetNumbering();
				
				let oNumLvl  = null;
				let nNumSuff = Asc.c_oAscNumberingSuff.None;
				if (numPr)
				{
					oNumLvl  = oNumbering.GetNum(numPr.NumId).GetLvl(numPr.Lvl);
					nNumSuff = oNumLvl.GetSuff();
				}
				else if (prevNumPr)
				{
					// MSWord uses tab instead of suff from PrevNum (74525)
					oNumLvl  = oNumbering.GetNum(prevNumPr.NumId).GetLvl(prevNumPr.Lvl);
					nNumSuff = Asc.c_oAscNumberingSuff.Tab;
				}
				
				var nNumJc     = oNumLvl.GetJc();
				var oNumTextPr = para.GetNumberingTextPr();
				
				var oPrevNumTextPr = prevNumPr ? para.Get_CompiledPr2(false).TextPr.Copy() : null;
				if (oPrevNumTextPr && (prevNumPr && prevNumPr.IsValid()))
				{
					var oPrevNumLvl = oNumbering.GetNum(prevNumPr.NumId).GetLvl(prevNumPr.Lvl);
					oPrevNumTextPr.Merge(para.TextPr.Value.Copy());
					oPrevNumTextPr.Merge(oPrevNumLvl.GetTextPr());
				}
				
				var X_start = this.X;
				
				let numWidth  = numItem.getNumWidth();
				let suffWidth = numItem.getSuffWidth();
				if (isRtl)
				{
					X_start = this.X + numItem.getVisibleWidth();
					if (align_Left === nNumJc)
						X_start -= numWidth;
					else if (align_Center === nNumJc)
						X_start -= numWidth / 2;
				}
				else
				{
					if (align_Right === nNumJc)
						X_start -= numWidth;
					else if (align_Center === nNumJc)
						X_start -= numWidth / 2;
				}
				
				var AutoColor = ( undefined !== BgColor && false === BgColor.Check_BlackAutoColor() ? new CDocumentColor(255, 255, 255, false) : new CDocumentColor(0, 0, 0, false) );
				
				var RGBA;
				if (oNumTextPr.Unifill)
				{
					oNumTextPr.Unifill.check(this.Theme, this.ColorMap);
					RGBA = oNumTextPr.Unifill.getRGBAColor();
					graphics.b_color1(RGBA.R, RGBA.G, RGBA.B, 255);
				}
				else
				{
					if (true === oNumTextPr.Color.Auto)
					{
						if(oNumTextPr.FontRef && oNumTextPr.FontRef.Color)
						{
							oNumTextPr.FontRef.Color.check(this.Theme, this.ColorMap);
							RGBA = oNumTextPr.FontRef.Color.RGBA;
							graphics.b_color1(RGBA.R, RGBA.G, RGBA.B, 255);
						}
						else
						{
							graphics.b_color1(AutoColor.r, AutoColor.g, AutoColor.b, 255);
						}
					}
					else
					{
						graphics.b_color1(oNumTextPr.Color.r, oNumTextPr.Color.g, oNumTextPr.Color.b, 255);
					}
				}
				
				if (numItem.havePrevNumbering() || reviewtype_Common !== nReviewType)
				{
					if (reviewtype_Common === nReviewType)
						graphics.b_color1(REVIEW_NUMBERING_COLOR.r, REVIEW_NUMBERING_COLOR.g, REVIEW_NUMBERING_COLOR.b, 255);
					else
						graphics.b_color1(oReviewColor.r, oReviewColor.g, oReviewColor.b, 255);
				}
				else if (isHavePrChange && numPr && !prevNumPr)
				{
					var oPrReviewColor = para.GetPrReviewColor();
					graphics.b_color1(oPrReviewColor.r, oPrReviewColor.g, oPrReviewColor.b, 255);
				}
				
				let baseY = Y;
				switch (oNumTextPr.VertAlign)
				{
					case AscCommon.vertalign_SubScript:
					{
						baseY -= AscCommon.vaKSub * oNumTextPr.FontSize * g_dKoef_pt_to_mm;
						break;
					}
					case AscCommon.vertalign_SuperScript:
					{
						baseY -= AscCommon.vaKSuper * oNumTextPr.FontSize * g_dKoef_pt_to_mm;
						break;
					}
				}
				
				// Рисуется только сам символ нумерации
				switch (nNumJc)
				{
					case align_Right:
						numItem.Draw(X_start, baseY, graphics, oNumbering, oNumTextPr, this.Theme, oPrevNumTextPr, isRtl);
						break;
					
					case align_Center:
						numItem.Draw(X_start, baseY, graphics, oNumbering, oNumTextPr, this.Theme, oPrevNumTextPr, isRtl);
						break;
					
					case align_Left:
					default:
						numItem.Draw(X_start, baseY, graphics, oNumbering, oNumTextPr, this.Theme, oPrevNumTextPr, isRtl);
						break;
				}
				
				if (true === oNumTextPr.Strikeout || true === oNumTextPr.Underline)
				{
					if (oNumTextPr.Unifill)
					{
						graphics.p_color(RGBA.R, RGBA.G, RGBA.B, 255);
					}
					else
					{
						if (true === oNumTextPr.Color.Auto)
							graphics.p_color(AutoColor.r, AutoColor.g, AutoColor.b, 255);
						else
							graphics.p_color(oNumTextPr.Color.r, oNumTextPr.Color.g, oNumTextPr.Color.b, 255);
					}
				}
				
				let lineW      = (oNumTextPr.FontSize / 18) * g_dKoef_pt_to_mm;
				let strikeoutY = (baseY - oNumTextPr.FontSize * g_dKoef_pt_to_mm * 0.27);
				let underlineY = (baseY + para.Lines[CurLine].Metrics.TextDescent * 0.4);
				if (numItem.havePrevNumbering() || reviewtype_Common !== nReviewType)
				{
					let prevNumWidth = numItem.getPrevNumWidth();
					if (reviewtype_Common === nReviewType)
						graphics.p_color(REVIEW_NUMBERING_COLOR.r, REVIEW_NUMBERING_COLOR.g, REVIEW_NUMBERING_COLOR.b, 255);
					else
						graphics.p_color(oReviewColor.r, oReviewColor.g, oReviewColor.b, 255);
					
					// Либо у нас есть удаленная часть, либо у нас одновременно добавлен и удален параграф, тогда мы зачеркиваем суффикс
					if (numItem.havePrevNumbering() || (!numItem.havePrevNumbering() && !numItem.haveFinalNumbering()))
					{
						if (numItem.haveFinalNumbering() || nReviewType !== reviewtype_Remove)
						{
							let x0 = X_start;
							let x1 = X_start + prevNumWidth;
							graphics.drawHorLine(0, strikeoutY, x0, x1, lineW);
						}
						else
						{
							let x0 = isRtl ? X_start - suffWidth : X_start;
							let x1 = x0 + numWidth + suffWidth;
							graphics.drawHorLine(0, strikeoutY, x0, x1, lineW);
						}
					}
					
					if (numItem.haveFinalNumbering())
					{
						if (isRtl)
						{
							let x0 = X_start - suffWidth;
							let x1 = X_start;
							graphics.drawHorLine(0, underlineY, x0, x1, lineW);
							
							x0 = X_start + prevNumWidth;
							x1 = X_start + numWidth;
							graphics.drawHorLine(0, underlineY, x0, x1, lineW);
						}
						else
						{
							let x0 = X_start + prevNumWidth;
							let x1 = X_start + numWidth + suffWidth;
							graphics.drawHorLine(0, underlineY, x0, x1, lineW);
						}
					}
				}
				else if (isHavePrChange && numPr && !prevNumPr)
				{
					let prevColor = para.GetPrReviewColor();
					graphics.p_color(prevColor.r, prevColor.g, prevColor.b, 255);
					
					let x0 = isRtl ? X_start - suffWidth : X_start;
					let x1 = x0 + numWidth + suffWidth;
					graphics.drawHorLine(0, underlineY, x0, x1, lineW);
				}
				else
				{
					let x0 = X_start;
					let x1 = x0 + numWidth;
					
					if (oNumTextPr.Strikeout)
						graphics.drawHorLine(0, strikeoutY, x0, x1, lineW);
					
					if (oNumTextPr.Underline)
						graphics.drawHorLine(0, underlineY, x0, x1, lineW);
				}
				
				if (true === editor.ShowParaMarks && (Asc.c_oAscNumberingSuff.Tab === nNumSuff || oNumLvl.IsLegacy()))
				{
					let tabSymbolWidth = 3.143; // ширина символа "стрелка влево" в шрифте Wingding3,10
					let tabX = this.X;
					
					if (!isRtl)
					{
						if (AscCommon.align_Left === nNumJc)
							tabX += numWidth;
						else if (AscCommon.align_Center === nNumJc)
							tabX += numWidth / 2;
					}
					
					graphics.SetFont({
						FontFamily : {Name : "ASCW3", Index : -1},
						FontSize   : 10,
						Italic     : false,
						Bold       : false
					});
					
					let tabCode = isRtl ? tab_Symbol_Rtl : tab_Symbol;
					if (suffWidth > tabSymbolWidth)
						graphics.FillText2(tabX + suffWidth / 2 - tabSymbolWidth / 2, Y, String.fromCharCode(tabCode), 0, suffWidth);
					else if (isRtl)
						graphics.FillText2(tabX, Y, String.fromCharCode(tabCode), 0, suffWidth);
					else
						graphics.FillText2(tabX, Y, String.fromCharCode(tabCode), tabSymbolWidth - suffWidth, suffWidth);
				}
			}
		}
		else if (para_PresentationNumbering === numItem.Type)
		{
			var bIsEmpty = para.IsEmpty();
			if (!bIsEmpty ||
				para.IsThisElementCurrent() ||
				para.Parent.IsSelectionUse() && para.Parent.IsSelectionEmpty() && para.Parent.Selection.StartPos === para.GetIndex())
			{
				if(!isRtl)
				{
					if (Pr.ParaPr.Ind.FirstLine < 0)
						numItem.Draw(X, Y, graphics, this);
					else
						numItem.Draw(para.Pages[CurPage].X + Pr.ParaPr.Ind.Left, Y, graphics, this);
				}
				else
				{
					if (Pr.ParaPr.Ind.FirstLine < 0)
						numItem.Draw(X + numItem.getVisibleWidth() / 2, Y, graphics, this);
					else
						numItem.Draw(para.Pages[CurPage].X + para.Pages[CurPage].XLimit - Pr.ParaPr.Ind.Right, Y, graphics, this);
				}
			}
		}
		
		this.X += numItem.getVisibleWidth();
	};
	ParagraphContentDrawState.prototype.calculateY = function(textPr)
	{
		let y = this.Y;
		
		if (AscCommon.vertalign_SubScript === textPr.VertAlign)
			y -= AscCommon.vaKSub * textPr.FontSize * g_dKoef_pt_to_mm;
		else if (AscCommon.vertalign_SuperScript === textPr.VertAlign)
			y -= AscCommon.vaKSuper * textPr.FontSize * g_dKoef_pt_to_mm;
		
		return y;
	};
	ParagraphContentDrawState.prototype.isSlideEditor = function()
	{
		return this.Paragraph && !this.Paragraph.bFromDocument;
	};
	ParagraphContentDrawState.prototype.isTextArtDraw = function()
	{
		return this.Graphics && this.Graphics.m_bIsTextDrawer;
	};
	ParagraphContentDrawState.prototype.isRtlMainDirection = function()
	{
		return this.Paragraph.isRtlDirection();
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.ParagraphContentDrawState = ParagraphContentDrawState;
	
})(window);

