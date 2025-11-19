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
	/**
	 * Class for handling a collection of custom annotation range marks
	 * @param {AscWord.Document} logicDocument
	 * @constructor
	 */
	function CustomMarks(logicDocument)
	{
		this.logicDocument = logicDocument;
		this.runs = {}; // runId -> markId
		this.paragraphs = {}; // paraId -> handlerId -> rangeId -> {start, end}
	}
	
	CustomMarks.prototype.add = function(mark)
	{
		let rangeId   = mark.getRangeId();
		let runId     = mark.getRun().GetId();
		let paraId    = mark.getParagraph().GetId();
		let handlerId = mark.getHandlerId();
		
		if (!this.runs[runId])
			this.runs[runId] = {};
		
		this.runs[runId][mark.getId()] = mark;
		
		if (!this.paragraphs[paraId])
			this.paragraphs[paraId] = {};
		
		if (!this.paragraphs[paraId][handlerId])
			this.paragraphs[paraId][handlerId] = {};
		
		if (!this.paragraphs[paraId][handlerId][rangeId])
			this.paragraphs[paraId][handlerId][rangeId] = {start : null, end : null};
		
		if (mark.isStart())
			this.paragraphs[paraId][handlerId][rangeId].start = mark;
		else
			this.paragraphs[paraId][handlerId][rangeId].end = mark;
	};
	CustomMarks.prototype.clear = function(paragraph, handlerId)
	{
		let paraId = paragraph.GetId();
		if (!this.paragraphs[paraId] || !this.paragraphs[paraId][handlerId])
			return;
		
		for (let rangeId in this.paragraphs[paraId][handlerId])
		{
			this.removeMarkFromRun(this.paragraphs[paraId][handlerId][rangeId].start);
			this.removeMarkFromRun(this.paragraphs[paraId][handlerId][rangeId].end);
		}
		
		delete this.paragraphs[paraId][handlerId];
	};
	CustomMarks.prototype.removeMarkFromRun = function(mark)
	{
		if (!mark)
			return;
		
		let runId = mark.getRun().GetId();
		if (!this.runs[runId])
			return;
		
		if (this.runs[runId] && this.runs[runId][mark.getId()])
			delete this.runs[runId][mark.getId()];
	};
	/**
	 * @param {AscWord.Paragraph} paragraph
	 * @param {AscWord.CParagraphContentPos} paraContentPos
	 * @returns {[]}
	 */
	CustomMarks.prototype.getStartedMarks = function(paragraph, paraContentPos)
	{
		let paraId = paragraph.GetId();
		if (!this.paragraphs[paraId])
			return [];
		
		let marks = this.paragraphs[paraId];
		
		let result = [];
		for (let handlerId in marks)
		{
			for (let rangeId in marks[handlerId])
			{
				let startMark = marks[handlerId][rangeId].start;
				let endMark   = marks[handlerId][rangeId].end;
				if (!startMark || !endMark)
					continue;
				
				let startPos = startMark.getParaPos();
				let endPos   = endMark.getParaPos();
				if (!startPos || !endPos)
					continue;
				
				if (paraContentPos.Compare(startPos) > 0 && paraContentPos.Compare(endPos) < 0)
					result.push(startMark);
			}
		}
		return result;
	};
	CustomMarks.prototype.selectRange = function(paraId, handlerId, rangeId)
	{
		let paragraph = AscCommon.g_oTableId.GetById(paraId);
		if (!this.paragraphs[paraId]
			|| !this.paragraphs[paraId][handlerId]
			|| !this.paragraphs[paraId][handlerId][rangeId])
			return;
		
		let start = this.paragraphs[paraId][handlerId][rangeId].start;
		let end   = this.paragraphs[paraId][handlerId][rangeId].end;
		
		let startPos = start.getParaPos();
		let endPos   = end.getParaPos();
		
		this.logicDocument.RemoveSelection();
		paragraph.SetSelectionUse(true);
		paragraph.SetSelectionContentPos(startPos, endPos, false);
		paragraph.SetThisElementCurrent();
		this.logicDocument.UpdateSelection();
		this.logicDocument.UpdateInterface();
		this.logicDocument.UpdateTracks();
	};
	CustomMarks.prototype.removeRange = function(paraId, handlerId, rangeId)
	{
		let paragraph = AscCommon.g_oTableId.GetById(paraId);
		if (!this.paragraphs[paraId]
			|| !this.paragraphs[paraId][handlerId]
			|| !this.paragraphs[paraId][handlerId][rangeId])
			return;
		
		this.removeMarkFromRun(this.paragraphs[paraId][handlerId][rangeId].start);
		this.removeMarkFromRun(this.paragraphs[paraId][handlerId][rangeId].end);
		delete this.paragraphs[paraId][handlerId][rangeId];
		
		this._checkEmptyParaHandler(paraId, handlerId);
		this._checkEmptyPara(paraId);
		
		if (paragraph)
			paragraph.ReDraw();
	};
	CustomMarks.prototype.removeAllRanges = function(paraId, handlerId)
	{
		let paragraph = AscCommon.g_oTableId.GetById(paraId);
		if (!this.paragraphs[paraId]
			|| !this.paragraphs[paraId][handlerId]
			|| !this.paragraphs[paraId][handlerId])
			return;
		
		for (let rangeId in this.paragraphs[paraId][handlerId])
		{
			this.removeMarkFromRun(this.paragraphs[paraId][handlerId][rangeId].start);
			this.removeMarkFromRun(this.paragraphs[paraId][handlerId][rangeId].end);
		}
		
		delete this.paragraphs[paraId][handlerId];
		this._checkEmptyPara(paraId);
		
		if (paragraph)
			paragraph.ReDraw();
	};
	CustomMarks.prototype.flatRunMarks = function(runId)
	{
		if (!this.runs[runId])
			return null;
		
		let result = [];
		this.forEachInRun(runId, function(mark){
			result.push(mark);
		});
		result.sort(function(a, b){
			let aPos = a.getPos();
			let bPos = b.getPos();
			if (aPos === bPos && a.isStart() !== b.isStart())
				return a.isStart() ? -1 : 1;
			
			return aPos - bPos;
		});
		return result;
	};
	CustomMarks.prototype.onAddToRun = function(runId, pos, count)
	{
		this.forEachInRun(runId, function(mark){
			mark.onAdd(pos, count);
		});
	};
	CustomMarks.prototype.onRemoveFromRun = function(runId, pos, count)
	{
		this.forEachInRun(runId, function(mark){
			mark.onRemove(pos, count);
		});
	};
	CustomMarks.prototype.onSplitRun = function(runId, pos, nextRunId)
	{
		let _t = this;
		let nextRun = AscCommon.g_oTableId.GetById(nextRunId);
		this.forEachInRun(runId, function(mark){
			mark.onSplit(pos, nextRun);
			
			if (mark.getRun().GetId() === nextRunId)
			{
				if (_t.runs[runId])
					delete _t.runs[runId][mark.getId()];
				
				if (!_t.runs[nextRunId])
					_t.runs[nextRunId] = {};
				
				_t.runs[nextRunId][mark.getId()] = mark;
			}
		});
	};
	CustomMarks.prototype.onConcatRun = function(runId, runLen, nextRunId)
	{
		let _t = this;
		let prevRun = AscCommon.g_oTableId.GetById(runId);
		this.forEachInRun(nextRunId, function(mark){
			mark.onConcat(runLen, prevRun);
			
			if (_t.runs[nextRunId])
				delete _t.runs[nextRunId][mark.getId()];
			
			if (!_t.runs[runId])
				_t.runs[runId] = {};
				
			_t.runs[runId][mark.getId()] = mark;
		});
		
		delete this.runs[nextRunId];
	};
	CustomMarks.prototype.forEachInRun = function(runId, f)
	{
		for (let id in this.runs[runId])
		{
			f.call(this, this.runs[runId][id]);
		}
	};
	CustomMarks.prototype._checkEmptyParaHandler = function(paraId, handlerId)
	{
		if (!this.paragraphs[paraId] || !this.paragraphs[paraId][handlerId])
			return;
		
		for (let rangeId in this.paragraphs[paraId][handlerId])
		{
			return;
		}
		
		delete this.paragraphs[paraId][handlerId];
	};
	CustomMarks.prototype._checkEmptyPara = function(paraId)
	{
		if (!this.paragraphs[paraId])
			return;
		
		for (let handlerId in this.paragraphs[paraId])
		{
			return;
		}
		
		delete this.paragraphs[paraId];
	};
	
	let idCounter = 0;
	function getNextId()
	{
		return ++idCounter;
	}
	
	/**
	 * @param run {AscWord.Run}
	 * @param pos {number}
	 * @param paragraph {AscWord.Paragraph}
	 * @param handlerId {string}
	 * @param rangeId {string}
	 * @constructor
	 */
	function CustomMark(run, pos, paragraph, handlerId, rangeId)
	{
		this.id        = getNextId();
		this.run       = run;
		this.paragraph = paragraph;
		this.pos       = pos;
		this.handlerId = handlerId;
		this.rangeId   = rangeId;
	}
	CustomMark.prototype.getId = function()
	{
		return this.id;
	};
	CustomMark.prototype.getRangeId = function()
	{
		return this.rangeId;
	};
	CustomMark.prototype.getHandlerId = function()
	{
		return this.handlerId;
	};
	CustomMark.prototype.isStart = function()
	{
		return true;
	};
	CustomMark.prototype.onAdd = function(pos, count)
	{
		if (this.pos >= pos)
			this.pos += count;
	};
	CustomMark.prototype.onRemove = function(pos, count)
	{
		if (this.pos > pos + count)
			this.pos -= count;
		else if (this.pos > pos)
			this.pos = Math.max(0, pos);
	};
	CustomMark.prototype.onSplit = function(pos, nextRun)
	{
		if (this.pos < pos)
			return;
		
		this.pos -= pos;
		this.run = nextRun;
	};
	CustomMark.prototype.onConcat = function(pos, prevRun)
	{
		this.pos += pos;
		this.run = prevRun;
	};
	CustomMark.prototype.getPos = function()
	{
		return this.pos;
	};
	CustomMark.prototype.getRun = function()
	{
		return this.run;
	};
	CustomMark.prototype.getParagraph = function()
	{
		return this.run ? this.run.GetParagraph() : null;
	};
	CustomMark.prototype.getParaPos = function()
	{
		let paragraph = this.getParagraph();
		let paraPos   = paragraph ? paragraph.GetPosByElement(this.run) : null;
		if (!paraPos)
			return new AscWord.CParagraphContentPos();
		
		paraPos.Update(this.pos, paraPos.GetDepth() + 1);
		return paraPos;
	};
	
	/**
	 * Метка начала промежутка
	 * @constructor
	 */
	function CustomMarkStart()
	{
		CustomMark.apply(this, arguments);
	}
	CustomMarkStart.prototype = Object.create(CustomMark.prototype);
	CustomMarkStart.prototype.constructor = CustomMarkStart;
	CustomMarkStart.prototype.isStart = function()
	{
		return true;
	};
	
	/**
	 * Метка окончания промежутка
	 * @constructor
	 */
	/**
	 * Метка начала промежутка
	 * @constructor
	 */
	function CustomMarkEnd()
	{
		CustomMark.apply(this, arguments);
	}
	CustomMarkEnd.prototype = Object.create(CustomMark.prototype);
	CustomMarkEnd.prototype.constructor = CustomMarkEnd;
	CustomMarkEnd.prototype.isStart = function()
	{
		return false;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.CustomMarks       = CustomMarks;
	AscWord.CustomMarkStart   = CustomMarkStart;
	AscWord.CustomMarkEnd     = CustomMarkEnd;
	
})(window);
