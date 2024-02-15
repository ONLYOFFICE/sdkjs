/*
 * (c) Copyright Ascensio System SIA 2010-2023
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
		
		this.rtl = false;
		
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
		
		this.rtl      = false;
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
		this.bidiFlow.begin(this.rtl);
	};
	ParagraphContentDrawState.prototype.endRange = function()
	{
		this.bidiFlow.end();
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
	ParagraphContentDrawState.prototype.handleBidiFlow = function(data)
	{
		let element = data[0];
		this.handleRun(data[1]);
		
		switch (element.Type)
		{
			case para_Text:
				this.handleText(element);
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
				this.handleRegularElement(element);
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
	 */
	ParagraphContentDrawState.prototype.handleText = function(text)
	{
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
			return;
		
		let y = this.Y;
		if (!sectPr)
		{
			let endTextPr = true !== this.Graphics.m_bIsTextDrawer ? this.Paragraph.GetParaEndCompiledPr() : this.textPr;
			this.updateGraphicsState(endTextPr);
			y = this.calculateY(endTextPr);
		}
		
		paraMark.Draw(this.X, y - this.yOffset, this.Graphics);
		this.X += paraMark.GetWidth();
	};
	/**
	 * @param element {CMathText | CMathAmp}
	 */
	ParagraphContentDrawState.prototype.handleMathElement = function(element)
	{
		if (para_Math_Placeholder === element.Type && this.Graphics.RENDERER_PDF_FLAG)
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
		if (!fieldChar.IsNumValue())
			return;
		
		// Draw the auto-calculated pageNum in the header/footer
		let complexField = fieldChar.GetComplexField();
		
		// We can't use textPr of the current run, since the properties of the actual text
		// may be different
		let textPr = complexField.GetFieldValueTextPr(true);
		this.updateGraphicsState(textPr);
		
		this.handleRegularElement(fieldChar);
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
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.ParagraphContentDrawState = ParagraphContentDrawState;
	
})(window);

