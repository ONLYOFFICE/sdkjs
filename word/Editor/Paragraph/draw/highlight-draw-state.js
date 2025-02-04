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
	const FLAG_HIGHLIGHT     = 0x0001;
	const FLAG_SEARCH        = 0x0004;
	const FLAG_COMMENT       = 0x0008;
	const FLAG_COMPLEX_FIELD = 0x0010;
	const FLAG_COLLABORATION = 0x0020;
	const FLAG_SHD           = 0x0040;
	const FLAG_PERM_RANGE    = 0x0080;
	
	/**
	 * Class for storing the current draw state of paragraph highlight (text/paragraph/field/etc. background)
	 * @param {AscWord.ParagraphDrawState} drawState
	 * @constructor
	 */
	function ParagraphHighlightDrawState(drawState)
	{
		this.drawState = drawState;
		
		this.Page   = 0;
		this.Line   = 0;
		this.Range  = 0;
		
		this.CurPos = new AscWord.CParagraphContentPos();
		
		this.DrawColl     = false;
		this.DrawMMFields = false;
		
		this.High     = new CParaDrawingRangeLines();
		this.Coll     = new CParaDrawingRangeLines();
		this.Find     = new CParaDrawingRangeLines();
		this.Comm     = new CParaDrawingRangeLines();
		this.Shd      = new CParaDrawingRangeLines();
		this.MMFields = new CParaDrawingRangeLines();
		this.CFields  = new CParaDrawingRangeLines();
		this.HyperCF  = new CParaDrawingRangeLines();
		this.Perm     = new CParaDrawingRangeLines();
		
		this.DrawComments       = true;
		this.DrawSolvedComments = true;
		this.haveCurrentComment = false;
		this.currentCommentId   = null;
		this.comments           = []; // current list of comments
		this.permRanges         = {};
		
		this.hyperlinks = [];
		
		this.searchCounter = 0;
		
		this.Paragraph = undefined;
		this.Graphics  = undefined;
		
		this.X  = 0;
		this.Y0 = 0;
		this.Y1 = 0;
		
		this.spaces = 0;
		
		this.InlineSdt = [];
		this.CollectFixedForms = false;
		
		this.complexFields = new AscWord.ParagraphComplexFieldStack();
		
		this.bidiFlow = new AscWord.BidiFlow(this);
		
		this.run       = null;
		this.highlight = highlight_None;
		this.shdColor  = null;
		this.shd       = null;
		
		this.permColor = null;
	}
	ParagraphHighlightDrawState.prototype.init = function(paragraph, graphics)
	{
		this.Paragraph = paragraph;
		this.Graphics  = graphics;
		
		let logicDocument = paragraph.GetLogicDocument();
		let commentManager = logicDocument && logicDocument.IsDocumentEditor() ? logicDocument.GetCommentsManager() : null;
		
		this.DrawColl           = !graphics.isPdf();
		this.DrawSearch         = logicDocument && logicDocument.IsDocumentEditor() && logicDocument.SearchEngine.Selection;
		this.DrawComments       = commentManager && commentManager.isUse();
		this.DrawSolvedComments = commentManager && commentManager.isUseSolved();
		this.DrawMMFields       = logicDocument && logicDocument.IsDocumentEditor() && logicDocument.isHighlightMailMergeFields();
		this.currentCommentId   = commentManager ? commentManager.getCurrentCommentId() : -1;
		
		this.permColor = new AscWord.CDocumentColor(233, 233, 233, 255);
		if (logicDocument && logicDocument.IsDocumentEditor())
		{
			let docApi = logicDocument.GetApi();
			if (docApi.isRestrictionView() || docApi.isRestrictionComments())
				this.permColor = new AscWord.CDocumentColor(255, 254, 213, 255);
		}
	};
	ParagraphHighlightDrawState.prototype.resetPage = function(page)
	{
		this.Page = page;
		
		this.CurPos = new AscWord.CParagraphContentPos();
		
		this.searchCounter = 0;
		
		this.comments           = [];
		this.haveCurrentComment = false;
		this.permRanges         = [];
		
		let pageEndInfo = this.Paragraph.GetEndInfoByPage(page - 1);
		if (pageEndInfo)
		{
			for (let index = 0, count = pageEndInfo.Comments.length; index < count; ++index)
			{
				this.addComment(pageEndInfo.Comments[index]);
			}
			
			for (let index = 0, count = pageEndInfo.PermRanges.length; index < count; ++index)
			{
				this.addPermRange(pageEndInfo.PermRanges[index]);
			}
		}
		this.complexFields.resetPage(this.Paragraph, page);
	};
	ParagraphHighlightDrawState.prototype.resetLine = function(line, top, bottom)
	{
		this.Line = line;
		this.Y0   = top;
		this.Y1   = bottom;
	};
	ParagraphHighlightDrawState.prototype.beginRange = function(range, X, spaceCount)
	{
		this.Range = range;
		this.X = X;
		this.checkNumbering();
		
		this.spaces = spaceCount;
		this.bidiFlow.begin(this.Paragraph.isRtlDirection());
		
		this.InlineSdt = [];
		
		this.High.Clear();
		this.Coll.Clear();
		this.Find.Clear();
		this.Comm.Clear();
		this.Shd.Clear();
		this.MMFields.Clear();
		this.CFields.Clear();
		this.HyperCF.Clear();
		this.Perm.Clear();
		
		this.run       = null;
		this.highlight = highlight_None;
		this.shdColor  = null;
		this.shd       = null;
	};
	ParagraphHighlightDrawState.prototype.endRange = function()
	{
		this.bidiFlow.end();
	};
	ParagraphHighlightDrawState.prototype.AddInlineSdt = function(oSdt)
	{
		this.InlineSdt.push(oSdt);
	};
	ParagraphHighlightDrawState.prototype.addComment = function(commentId)
	{
		let comment = this.getComment(commentId);
		if (!comment)
			return;
		
		if (commentId === this.currentCommentId)
			this.haveCurrentComment = true;
		
		this.comments.push(commentId);
	};
	ParagraphHighlightDrawState.prototype.removeComment = function(commentId)
	{
		let comment = this.getComment(commentId);
		if (!comment)
			return;
		
		if (commentId === this.currentCommentId)
			this.haveCurrentComment = false;
		
		let index = this.comments.indexOf(commentId);
		if (-1 !== index)
			this.comments.splice(index, 1);
	};
	ParagraphHighlightDrawState.prototype.addPermRange = function(rangeId)
	{
		this.permRanges.push(rangeId);
	};
	ParagraphHighlightDrawState.prototype.removePermRange = function(rangeId)
	{
		if (!this.permRanges.length)
			return;
		
		if (this.permRanges[this.permRanges.length - 1] === rangeId)
		{
			--this.permRanges.length;
		}
		else
		{
			let pos = this.permRanges.indexOf(rangeId);
			if (-1 !== pos)
				this.permRanges.splice(pos, 1);
		}
	};
	ParagraphHighlightDrawState.prototype.increaseSearchCounter = function()
	{
		++this.searchCounter;
	};
	ParagraphHighlightDrawState.prototype.decreaseSearchCounter = function()
	{
		--this.searchCounter;
	};
	ParagraphHighlightDrawState.prototype.IsCollectFixedForms = function()
	{
		return this.CollectFixedForms;
	};
	ParagraphHighlightDrawState.prototype.SetCollectFixedForms = function(isCollect)
	{
		this.CollectFixedForms = isCollect;
	};
	ParagraphHighlightDrawState.prototype.handleRunElement = function(element, run, collaborationColor)
	{
		if (!this.complexFields.checkRunElement(element))
			return;
		
		if (para_Drawing === element.Type && !element.IsInline())
			return;
		
		let flags = this.getFlags(element, !!collaborationColor);
		let hyperlink = this.getHyperlinkObject();
		this.bidiFlow.add([element, run, flags, hyperlink, collaborationColor, this.comments.slice(), this.haveCurrentComment], element.getBidiType());
	};
	ParagraphHighlightDrawState.prototype.handleBidiFlow = function(data)
	{
		let element    = data[0];
		let run        = data[1];
		let flags      = data[2];
		let hyperlink  = data[3];
		let collColor  = data[4];
		let comments   = data[5];
		let curComment = data[6];
		
		let w = element.GetWidthVisible();
		
		this.handleRun(run);
		this.addHighlight(this.X, this.X + w, flags, hyperlink, collColor, comments, curComment);
		
		this.X += w;
	};
	ParagraphHighlightDrawState.prototype.handleParaMath = function(math)
	{
		if (!math || math.Root.IsEmptyRange(this.Line, this.Range))
			return;
		
		this.bidiFlow.end();
		
		let y0 = this.Y0;
		let y1 = this.Y1;
		
		let coll = this.Coll;
		let comm = this.Comm;
		
		this.Coll = new CParaDrawingRangeLines();
		this.Comm = new CParaDrawingRangeLines();
		
		math.Root.Draw_HighLights(this);
		
		let mathComments = this.Comm.getNext();
		if (mathComments)
		{
			let bounds = math.Root.Get_LineBound(this.Line, this.Range);
			comm.Add(bounds.Y, bounds.Y + bounds.H, bounds.X, bounds.X + bounds.W, 0, 0, 0, 0, mathComments.Additional);
		}
		
		let mathColl = this.Coll.getNext();
		if (mathColl)
		{
			let bounds = math.Root.Get_LineBound(this.Line, this.Range);
			coll.Add(bounds.Y, bounds.Y + bounds.H, bounds.X, bounds.X + bounds.W, 0, mathColl.r, mathColl.g, mathColl.b);
		}
		
		this.Coll = coll;
		this.Comm = comm;
		
		this.Y0 = y0;
		this.Y1 = y1;
	};
	ParagraphHighlightDrawState.prototype.handleMathBase = function(element)
	{
		this.bidiFlow.end();
		
		let textPr = element.Get_CompiledCtrPrp();
		
		let shdColor = (textPr.Shd && !textPr.Shd.IsNil() ? textPr.Shd.GetSimpleColor(this.drawState.getTheme(), this.drawState.getColorMap()) : null);
		
		let w = element.Get_LineBound(this.Line, this.Range).W;
		
		let x  = this.X;
		let y0 = this.Y0;
		let y1 = this.Y1;
		
		let startPos = 0;
		let endPos   = element.Content.length - 1;
		if (!element.bOneLine)
		{
			let rangeInfo = element.getRangePos(this.Line, this.Range);
			
			startPos = rangeInfo[0];
			endPos   = rangeInfo[1];
		}
		
		for (let pos = startPos; pos <= endPos; ++pos)
		{
			element.Content[pos].Draw_HighLights(this);
		}
		
		// Add after because we are rendering in reverse direction
		if (shdColor)
			this.Shd.Add(y0, y1, x, x + w, 0, shdColor.r, shdColor.g, shdColor.b);
		
		if (highlight_None !== textPr.HighLight)
			this.High.Add(y0, y1, x, x + w, 0, textPr.HighLight.r, textPr.HighLight.g, textPr.HighLight.b);
		
		this.X = x + w;
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	ParagraphHighlightDrawState.prototype.checkNumbering = function()
	{
		let paraNumbering = this.Paragraph.Numbering;
		if (!paraNumbering.checkRange(this.Range, this.Line))
			return;
		
		let x = this.X;
		this.X += paraNumbering.WidthVisible;
		
		let numPr = this.drawState.getParagraphCompiledPr().ParaPr.NumPr;
		let paraParent = this.Paragraph.GetParent();

		if (para_Numbering !== paraNumbering.Type
			|| !numPr
			|| !numPr.IsValid()
			|| !paraParent
			|| !paraParent.IsEmptyParagraphAfterTableInTableCell(this.Paragraph.GetIndex()))
			return;
		
		let numManager = paraParent.GetNumbering();
		let numLvl     = numManager.GetNum(numPr.NumId).GetLvl(numPr.Lvl);
		let numJc      = numLvl.GetJc();
		let numTextPr  = this.Paragraph.GetNumberingTextPr();
		
		if (AscCommon.align_Right === numJc)
			x -= paraNumbering.WidthNum;
		else if (AscCommon.align_Center === numJc)
			x -= paraNumbering.WidthNum / 2;
		
		if (highlight_None !== numTextPr.HighLight)
			this.High.Add(this.Y0, this.Y1, x, x + paraNumbering.WidthNum + paraNumbering.WidthSuff, 0, numTextPr.HighLight.r, numTextPr.HighLight.g, numTextPr.HighLight.b, undefined, numTextPr);
	};
	/**
	 * @param {string} commentId
	 * @returns {CComment}
	 */
	ParagraphHighlightDrawState.prototype.getComment = function(commentId)
	{
		if (!this.DrawComments)
			return null;
		
		let comment = AscCommon.g_oTableId.GetById(commentId);
		if (comment
			&& (this.DrawSolvedComments || !comment.IsSolved())
			&& AscCommon.UserInfoParser.canViewComment(comment.GetUserName()))
			return comment;
		
		return null;
	};
	/**
	 * @param {AscWord.CRun} run
	 */
	ParagraphHighlightDrawState.prototype.handleRun = function(run)
	{
		if (this.run === run)
			return;
		
		this.run = run;
		
		let textPr = run.getCompiledPr();
		let shd    = textPr.Shd;
		
		this.shd = textPr.Shd;
		this.shdColor = shd && !shd.IsNil() ? shd.GetSimpleColor(this.drawState.getTheme(), this.drawState.getColorMap()) : null;
		if (!this.shdColor || this.shdColor.IsAuto() || (run.IsMathRun() && run.IsPlaceholder()))
			this.shdColor = null;
		
		this.highlight = textPr.HighLight;
		if (textPr.HighlightColor)
		{
			textPr.HighlightColor.check(this.drawState.getTheme(), this.drawState.getColorMap());
			let RGBA = textPr.HighlightColor.RGBA;
			this.highlight = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B, RGBA.A);
		}
	};
	/**
	 *
	 */
	ParagraphHighlightDrawState.prototype.addHighlight = function(startX, endX, flags, hyperlink, collColor, comments, curComment)
	{
		let startY = this.Y0;
		let endY   = this.Y1;
		
		if ((flags & FLAG_SHD) && this.shdColor)
			this.Shd.Add(startY, endY, startX, endX, 0, this.shdColor.r, this.shdColor.g, this.shdColor.b, undefined, this.shd);
		
		if (hyperlink)
			this.HyperCF.Add(startY, endY, startX, endX, 0, 0, 0, 0, {HyperlinkCF : hyperlink});
		
		if (flags & FLAG_COMPLEX_FIELD)
			this.CFields.Add(startY, endY, startX, endX, 0, 0, 0, 0);
		
		if (flags & FLAG_COMMENT && comments.length)
			this.Comm.Add(startY, endY, startX, endX, 0, 0, 0, 0, {Active : curComment, CommentId : comments});
		
		if ((flags & FLAG_HIGHLIGHT) && (this.highlight && highlight_None !== this.highlight))
			this.High.Add(startY, endY, startX, endX, 0, this.highlight.r, this.highlight.g, this.highlight.b, undefined, this.highlight);
		
		if (flags & FLAG_SEARCH)
			this.Find.Add(startY, endY, startX, endX, 0, 0, 0, 0);
		
		if ((flags & FLAG_COLLABORATION) && collColor)
			this.Coll.Add(startY, endY, startX, endX, 0, collColor.r, collColor.g, collColor.b);
		
		if (flags & FLAG_PERM_RANGE && this.permColor)
			this.Perm.Add(startY, endY, startX, endX, 0, this.permColor.r, this.permColor.g, this.permColor.b);
	};
	ParagraphHighlightDrawState.prototype.pushHyperlink = function(hyperlink)
	{
		this.hyperlinks.push(hyperlink)
	};
	ParagraphHighlightDrawState.prototype.popHyperlink = function()
	{
		if (!this.hyperlinks.length)
			return;
		
		--this.hyperlinks.length;
	};
	ParagraphHighlightDrawState.prototype.getHyperlinkObject = function()
	{
		if (this.hyperlinks.length)
			return this.hyperlinks[this.hyperlinks.length - 1];
		
		let complexField = null;
		if (this.complexFields.isComplexField() && (complexField = this.complexFields.getReferenceToHyperlink()))
			return complexField.GetInstruction();
		
		return null;
	};
	ParagraphHighlightDrawState.prototype.isComplexFieldHighlight = function()
	{
		return (this.complexFields.isComplexField()
			&& !this.complexFields.isHiddenComplexFieldPart()
			&& this.complexFields.isCurrentComplexField()
			&& !this.complexFields.isHyperlinkField());
	};
	ParagraphHighlightDrawState.prototype.getFlags = function(element, isCollaboration)
	{
		let flags = 0;
		if (this.DrawSearch && this.searchCounter > 0)
			flags |= FLAG_SEARCH;
		if (this.isComplexFieldHighlight())
			flags |= FLAG_COMPLEX_FIELD;
		if (this.permRanges.length > 0)
			flags |= FLAG_PERM_RANGE;
		
		if (element.Type !== para_End)
			flags |= FLAG_SHD;
		
		switch (element.Type)
		{
			case para_PageNum:
			case para_PageCount:
			case para_Drawing:
			case para_Tab:
			case para_Text:
			case para_Math_Text:
			case para_Math_Placeholder:
			case para_Math_BreakOperator:
			case para_Math_Ampersand:
			case para_Sym:
			case para_FootnoteReference:
			case para_FootnoteRef:
			case para_Separator:
			case para_ContinuationSeparator:
			case para_EndnoteReference:
			case para_EndnoteRef:
			{
				if (this.DrawComments && this.comments.length > 0)
					flags |= FLAG_COMMENT;
				else
					flags |= FLAG_HIGHLIGHT;
				
				if (this.DrawSearch && this.searchCounter > 0)
					flags |= FLAG_SEARCH;
				else if (this.DrawColl && isCollaboration)
					flags |= FLAG_COLLABORATION;
				
				break;
			}
			case para_Space:
			{
				if (this.spaces > 0)
				{
					if (this.DrawComments && this.comments.length > 0)
						flags |= FLAG_COMMENT;
					else
						flags |= FLAG_HIGHLIGHT;
					
					--this.spaces;
				}
				
				if (this.DrawSearch && this.searchCounter > 0)
					flags |= FLAG_SEARCH;
				else if (this.DrawColl && isCollaboration)
					flags |= FLAG_COLLABORATION;
				
				break;
			}
			case para_End:
			{
				if (this.DrawColl && isCollaboration)
					flags |= FLAG_COLLABORATION;
				
				break;
			}
			case para_FieldChar:
			{
				if (element.IsVisual())
				{
					if (this.DrawComments && this.comments.length > 0)
						flags |= FLAG_COMMENT;
					else
						flags |= FLAG_HIGHLIGHT;
					
					if (this.DrawSearch && this.searchCounter > 0)
						flags |= FLAG_SEARCH;
					else if (this.DrawColl && isCollaboration)
						flags |= FLAG_COLLABORATION;
					
					if (element.IsFormField())
						flags |= FLAG_COMPLEX_FIELD;
				}
				break;
			}
		}
		
		return flags;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.ParagraphHighlightDrawState = ParagraphHighlightDrawState;
	
})(window);

