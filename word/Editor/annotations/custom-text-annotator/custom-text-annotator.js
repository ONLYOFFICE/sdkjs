/*
 * (c) Copyright Ascensio System SIA 2010-2025
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
	const MAX_ACTION_TIME = 20;
	
	/**
	 * CustomTextAnnotator manages external text annotation workflows for document paragraphs.
	 *
	 * This class handles the process of sending paragraph text to external applications/plugins
	 * for analysis and receives back highlight positions that can be rendered and made interactive.
	 * It maintains state for paragraphs at different stages of the annotation pipeline.
	 
	 * @param {AscWord.Document} logicDocument
	 * @constructor
	 */
	function CustomTextAnnotator(logicDocument)
	{
		this.logicDocument = logicDocument;
		
		this.waitingParagraphs  = {};
		this.paragraphs         = {};
		this.checkingParagraphs = {};
		
		this.eventManager = this.logicDocument.GetApi().getTextAnnotatorEventManager();
		this.marks        = new AscWord.CustomMarks(logicDocument);
		
		this.textGetter = new TextGetter(this.marks);
		this.markSetter = new MarkSetter(this.marks);
	}
	
	CustomTextAnnotator.prototype.getMarks = function()
	{
		return this.marks;
	};
	CustomTextAnnotator.prototype.isActive = function()
	{
		return true;
	};
	CustomTextAnnotator.prototype.addParagraphToCheck = function(para)
	{
		this.checkingParagraphs[para.GetId()] = para;
	};
	CustomTextAnnotator.prototype.continueProcessing = function()
	{
		if (!this.isActive())
			return;
		
		let startTime = performance.now();
		while (true)
		{
			if (performance.now() - startTime > MAX_ACTION_TIME)
				break;
			
			let paragraph = this.popNextParagraph();
			if (!paragraph)
				break;
			
			this.handleParagraph(paragraph);
		}
	};
	CustomTextAnnotator.prototype.onCurrentParagraph = function(paragraph)
	{
		this.eventManager.onCurrentRanges(paragraph, this.getCurrentRanges(paragraph));
	};
	CustomTextAnnotator.prototype.onClick = function(x, y, page, e)
	{
		if (!e || AscCommon.g_mouse_button_left !== e.Button)
			return;
		
		let anchorPos = null;
		let drawingObjects = this.logicDocument.GetDrawingObjects();
		if (-1 !== drawingObjects.IsInDrawingObject(x, y, page, this.logicDocument))
		{
			let paragraph = drawingObjects.getParagraphByXY(x, y, page);
			if (paragraph)
			{
				let transform = paragraph.Get_ParentTextTransform();
				if (transform)
				{
					let inverted = transform.CreateDublicate().Invert();
					let _x = inverted.TransformPointX(x, y);
					let _y = inverted.TransformPointY(x, y);
					x = _x;
					y = _y;
				}
				anchorPos = paragraph.Get_NearestPos(0, x, y, false);
			}
		}
		else
		{
			anchorPos = this.logicDocument.Get_NearestPos(page, x, y);
		}
		
		if (!anchorPos || !anchorPos.Paragraph || !anchorPos.ContentPos)
			return;
		
		let ranges = this.getRangesByParaContentPos(anchorPos.Paragraph, anchorPos.ContentPos);
		this.eventManager.onClick(anchorPos.Paragraph, ranges);
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CustomTextAnnotator.prototype.popNextParagraph = function()
	{
		for (let paraId in this.checkingParagraphs)
		{
			let paragraph = this.checkingParagraphs[paraId];
			delete this.checkingParagraphs[paraId];
			
			if (!paragraph.IsUseInDocument())
				continue;
			
			return paragraph;
		}
		
		return null;
	};
	CustomTextAnnotator.prototype.handleParagraph = function(paragraph)
	{
		this.eventManager.onChangeParagraph(paragraph);
	};
	CustomTextAnnotator.prototype.getEventObject = function(paragraph)
	{
		this.textGetter.check(paragraph);
		return {
			text   : this.textGetter.text,
			ranges : this.textGetter.ranges
		};
	};
	CustomTextAnnotator.prototype.highlightTextResponse = function(handlerId, paraId, ranges)
	{
		ranges.sort(function(a, b){
			return a.start - b.start;
		});
		let paragraph = AscCommon.g_oTableId.GetById(paraId);
		this.markSetter.placeMarks(paragraph, ranges, handlerId);
		
		let _ranges = [];
		ranges.forEach(r => _ranges.push([r.start, r.length, r.id]))
		//console.log(`Response from handlerId=${handlerId} ParaId=${paraId}; Ranges=${_ranges}`);
		
		// TODO: Надо проверить ситуацию, когда меток не было, и они не появились
		paragraph.ReDraw();
	};
	CustomTextAnnotator.prototype.getCurrentRanges = function(paragraph)
	{
		if (!paragraph)
			return {};
		
		let paraPos = paragraph.GetParaContentPos(false, false);
		return this.getRangesByParaContentPos(paragraph, paraPos);
	};
	CustomTextAnnotator.prototype.getRangesByParaContentPos = function(paragraph, paraPos)
	{
		if (!paraPos || !paragraph)
			return {};
		
		let marks = this.marks.getMarksByPos(paragraph, paraPos);
		let ranges = {};
		for (let i = 0; i < marks.length; ++i)
		{
			let mark      = marks[i];
			let handlerId = mark.getHandlerId();
			let rangeId   = mark.getRangeId();
			if (!ranges[handlerId])
				ranges[handlerId] = {};
			
			if (!ranges[handlerId][rangeId])
				ranges[handlerId][rangeId] = mark;
		}
		return ranges;
	};
	/**
	 * @param customMarks {AscWord.CustomMarks}
	 * @constructor
	 */
	function TextGetter(customMarks)
	{
		AscWord.DocumentVisitor.call(this);
		this.marks  = customMarks;
		this.text   = "";
		this.pos    = 0;
		this.ranges = {};
	}
	TextGetter.prototype = Object.create(AscWord.DocumentVisitor.prototype);
	TextGetter.prototype.constructor = TextGetter;
	TextGetter.prototype.check = function(paragraph)
	{
		this.pos    = 0
		this.text   = "";
		this.ranges = {};
		this.traverseParagraph(paragraph);
	};
	TextGetter.prototype.run = function(run)
	{
		this.handleMarks(run);
		for (let pos = 0, len = run.GetElementsCount(); pos < len; ++pos)
		{
			let item = run.GetElement(pos);
			if (item.IsSpace() || (item.IsText() && item.IsNBSP()))
				this.text += String.fromCodePoint(0x20);
			else if (item.IsText())
				this.text += String.fromCodePoint(item.GetCodePoint());
			else if (item.IsTab())
				this.text += String.fromCodePoint(0x09);
			else if (item.IsBreak())
				this.text += "\n";
			else
				continue;
			
			++this.pos;
		}
		return true;
	};
	TextGetter.prototype.handleMarks = function(run)
	{
		let _t = this;
		this.marks.forEachInRun(run.GetId(), function(mark){
			let rangeId = mark.getRangeId();
			let handlerId = mark.getHandlerId();
			let pos = mark.getPos() + _t.pos;
			
			if (!_t.ranges[handlerId])
				_t.ranges[handlerId] = {};
			
			if (!_t.ranges[handlerId][rangeId])
				_t.ranges[handlerId][rangeId] = {};
			
			if (mark.isStart())
				_t.ranges[handlerId][rangeId].start = pos;
			else
				_t.ranges[handlerId][rangeId].end = pos;
		});
	};
	/**
	 * @param customMarks {AscWord.CustomMarks}
	 * @constructor
	 */
	function MarkSetter(customMarks)
	{
		AscWord.DocumentVisitor.call(this);
		
		this.para      = null;
		this.pos       = -1;
		this.firstPos  = true;
		this.ranges    = [];
		this.rangePos  = 0;
		this.endMarks  = {};
		this.handlerId = null;
		
		this.customMarks = customMarks;
	}
	MarkSetter.prototype = Object.create(AscWord.DocumentVisitor.prototype);
	MarkSetter.prototype.constructor = MarkSetter;
	MarkSetter.prototype.placeMarks = function(paragraph, ranges, handlerId)
	{
		this.para      = paragraph;
		this.pos       = -1;
		this.firstPos  = true;
		this.ranges    = ranges;
		this.rangePos  = 0;
		this.endMarks  = {};
		this.handlerId = handlerId;
		
		this.customMarks.clear(paragraph, handlerId);
		
		this.traverseParagraph(paragraph);
	};
	MarkSetter.prototype.run = function(run)
	{
		this.checkStartPosition(run);
		
		for (let pos = 0, len = run.GetElementsCount(); pos < len; ++pos)
		{
			let item = run.GetElement(pos);
			if (item.IsText() || item.IsSpace() || item.IsTab() || item.IsBreak())
				this.handleNextPosition(run, pos + 1);
		}
		return true;
	};
	MarkSetter.prototype.handleNextPosition = function(run, pos)
	{
		++this.pos;
		
		while (this.rangePos < this.ranges.length && this.ranges[this.rangePos].start <= this.pos)
		{
			let range = this.ranges[this.rangePos];
			let endPos = range.start + range.length;
			let markId = range.id;
			
			let mark = new AscWord.CustomMarkStart(run, pos, this.para, this.handlerId, markId);
			this.customMarks.add(mark);
			
			if (!this.endMarks[endPos])
				this.endMarks[endPos] = [];
			
			this.endMarks[endPos].push(markId);
			
			++this.rangePos;
		}
		
		if (this.endMarks[this.pos])
		{
			for (let i = 0, count = this.endMarks[this.pos].length; i < count; ++i)
			{
				let markId = this.endMarks[this.pos][i];
				let mark = new AscWord.CustomMarkEnd(run, pos, this.para, this.handlerId, markId);
				this.customMarks.add(mark);
			}
		}
	};
	MarkSetter.prototype.checkStartPosition = function(run)
	{
		if (!this.firstPos)
			return;
		
		this.firstPos = false;
		this.handleNextPosition(run, 0);
	};
	//-------------------------------------------------------------export-----------------------------------------------
	AscWord.CustomTextAnnotator = CustomTextAnnotator;
})(window);
