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
	 * Class for traversing the entire document structure
	 * @constructor
	 */
	function DocumentVisitor()
	{
		this.stopped = false;
	}
	
	/**
	 * @param {AscWord.CSelectedContent} selectedContent
	 */
	DocumentVisitor.prototype.traverseSelectedContent = function(selectedContent)
	{
		let content = selectedContent.GetContentArray();
		for (let i = 0, count = content.length; i < count; ++i)
		{
			content[i].visit(this);
			if (this.isStopped())
				return;
		}
	};
	/**
	 * @param {AscWord.Document} logicDocument
	 * @param {boolean} isSelection
	 */
	DocumentVisitor.prototype.traverseDocument = function(logicDocument, isSelection)
	{
		logicDocument.visit(this, isSelection);
	};
	/**
	 * @param {AscWord.Document | AscWord.DocumentContent} docContent
	 * @param {boolean} isSelection
	 */
	DocumentVisitor.prototype.traverseDocContent = function(docContent, isSelection)
	{
		docContent.visit(this, isSelection);
	};
	DocumentVisitor.prototype.traverseParagraph = function(paragraph)
	{
		paragraph.visit(this);
	};
	DocumentVisitor.prototype.stop = function()
	{
		this.stopped = true;
	};
	DocumentVisitor.prototype.isStopped = function()
	{
		return this.stopped;
	};
	DocumentVisitor.prototype.paragraph = function(paragraph, isStart, start, end)
	{
		return false;
	};
	DocumentVisitor.prototype.table = function(table, isStart)
	{
		return false;
	};
	DocumentVisitor.prototype.tableRow = function(tableRow, isStart)
	{
		return false;
	};
	DocumentVisitor.prototype.tableCell = function(tableCell, isStart, start, end)
	{
		return false;
	};
	DocumentVisitor.prototype.fldSimple = function(field, isStart, start, end)
	{
		return false;
	};
	DocumentVisitor.prototype.blockLevelSdt = function(sdt, isStart, start, end)
	{
		return false;
	};
	DocumentVisitor.prototype.inlineLevelSdt = function(sdt, isStart, start, end)
	{
		return false;
	};
	DocumentVisitor.prototype.oform = function(form, isStart, start, end)
	{
		return false;
	};
	DocumentVisitor.prototype.run = function(run, isStart, start, end)
	{
		return false;
	};
	DocumentVisitor.prototype.oMath = function(oMath, isStart, start, end)
	{
		return false;
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	DocumentVisitor.prototype.visitDocContent = function(docContent, isSelection)
	{
		if (isSelection && !docContent.IsSelectionUse())
			return;
		
		let start = isSelection ? Math.min(docContent.Selection.StartPos, docContent.Selection.EndPos) : 0;
		let end   = isSelection ? Math.max(docContent.Selection.StartPos, docContent.Selection.EndPos) : docContent.Content.length - 1;
		
		for (let i = start; i <= end; ++i)
		{
			docContent.Content[i].visit(this, isSelection);
			if (this.isStopped())
				return;
		}
	};
	DocumentVisitor.prototype.visitDrawings = function(drawingObjects, isSelection)
	{
		let docContent = drawingObjects.getTargetDocContent();
		if (docContent)
			return docContent.visit(this, isSelection);
		
		// TODO: We need to visit the place where the main drawing lies
	};
	DocumentVisitor.prototype.visitNotes = function(notesController, isSelection)
	{
		if (isSelection && !notesController.IsSelectionUse())
			return;
		
		let notes = notesController.private_GetSelectionArray();
		for (let i = 0; i < notes.length; ++i)
		{
			if (notes[i])
				notes[i].visit(this, isSelection);
			
			if (this.isStopped())
				return;
		}
	};
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	AscWord.Document.prototype.visit = function(visitor, isSelection)
	{
		if (!isSelection)
		{
			visitor.visitDocContent(this, false);
			
			this.SectionsInfo.forEachHdrFtr(function(hdrFtr){
				if (visitor.isStopped())
					return;
				
				hdrFtr.GetContent().visit(visitor, false);
			});
		}
		else
		{
			switch (this.GetDocPosType())
			{
				case docpostype_Content:
					visitor.visitDocContent(this, true);
					break;
				case docpostype_HdrFtr:
				{
					let hdrFtr = this.HdrFtr.GetCurHdrFtr();
					if (hdrFtr)
						hdrFtr.GetContent().visit(visitor, true);
					break;
				}
				case docpostype_DrawingObjects:
					visitor.visitDrawings(this.DrawingObjects, true);
					break;
				case docpostype_Footnotes:
					visitor.visitNotes(this.Footnotes, true);
					break;
				case docpostype_Endnotes:
					visitor.visitNotes(this.Endnotes, true);
					break;
			}
		}
	};
	AscWord.Paragraph.prototype.visit = function(visitor, isSelection)
	{
		let isPartial = isSelection && !this.IsSelectedAll();
		if (isPartial && (!this.Selection.Use || this.IsSelectionEmpty(true)))
			return;
		
		let start = isPartial ? Math.min(this.Selection.StartPos, this.Selection.EndPos) : undefined;
		let end   = isPartial ? Math.max(this.Selection.StartPos, this.Selection.EndPos) : undefined;
		if (visitor.paragraph(this, true, start, end) || visitor.isStopped())
			return;
		
		let from = isPartial ? start : 0;
		let to   = isPartial ? end : this.Content.length - 1;
		for (let i = from; i <= to; ++i)
		{
			this.Content[i].visit(visitor, isPartial);
			if (visitor.isStopped())
				return;
		}
		
		visitor.paragraph(this, false, start, end);
	};
	AscWord.Table.prototype.visit = function(visitor, isSelection)
	{
		let isPartial = isSelection && !this.IsSelectedAll();
		if (isPartial && !this.Selection.Use)
			return;
		
		let cells = isPartial ? this.GetSelectionArray(false) : null;
		if (cells && !cells.length)
			return;
		
		if (visitor.table(this, true) || visitor.isStopped())
			return;
		
		let startRow = cells ? cells[0].Row : 0;
		let endRow   = cells ? cells[cells.length - 1].Row : this.GetRowsCount() - 1;
		for (let i = startRow; i <= endRow; ++i)
		{
			this.GetRow(i).visit(visitor, isPartial);
			if (visitor.isStopped())
				return;
		}
		
		visitor.table(this, false);
	};
	AscWord.TableRow.prototype.visit = function(visitor, isSelection)
	{
		let cells = null;
		if (isSelection)
		{
			cells = [];
			
			let rowIndex  = this.GetIndex();
			let selection = this.GetTable().GetSelectionArray(false);
			for (let i = 0; i < selection.length; ++i)
			{
				if (rowIndex === selection[i].Row)
					cells.push(selection[i].Cell);
			}
			
			if (!cells.length)
				return;
		}
		
		if (visitor.tableRow(this, true) || visitor.isStopped())
			return;
		
		let isTextSelection = !!isSelection && !this.GetTable().IsCellSelection();
		for (let i = 0, count = cells ? cells.length : this.GetCellsCount(); i < count; ++i)
		{
			this.GetCell(cells ? cells[i] : i).visit(visitor, isTextSelection);
			if (visitor.isStopped())
				return;
		}
		
		visitor.tableRow(this, false);
	};
	AscWord.TableCell.prototype.visit = function(visitor, isSelection)
	{
		let content   = this.GetContent();
		let isPartial = isSelection && !content.IsSelectedAll();
		if (isPartial && !content.Selection.Use)
			return;
		
		let start = isPartial ? Math.min(content.Selection.StartPos, content.Selection.EndPos) : undefined;
		let end   = isPartial ? Math.max(content.Selection.StartPos, content.Selection.EndPos) : undefined;
		if (visitor.tableCell(this, true, start, end) || visitor.isStopped())
			return;
		
		content.visit(visitor, isPartial);
		if (visitor.isStopped())
			return;
		
		visitor.tableCell(this, false, start, end);
	};
	AscWord.BlockLevelSdt.prototype.visit = function(visitor, isSelection)
	{
		let content   = this.Content;
		let isPartial = isSelection && !content.IsSelectedAll() && !this.IsPlaceHolder();
		if (isPartial && !content.Selection.Use)
			return;
		
		let start = isPartial ? Math.min(content.Selection.StartPos, content.Selection.EndPos) : undefined;
		let end   = isPartial ? Math.max(content.Selection.StartPos, content.Selection.EndPos) : undefined;
		if (visitor.blockLevelSdt(this, true, start, end) || visitor.isStopped())
			return;
		
		content.visit(visitor, isPartial);
		if (visitor.isStopped())
			return;
		
		visitor.blockLevelSdt(this, false, start, end);
	};
	AscWord.DocumentContent.prototype.visit = function(visitor, isSelection)
	{
		if (isSelection && docpostype_DrawingObjects === this.GetDocPosType())
			visitor.visitDrawings(this.DrawingObjects, true);
		else
			visitor.visitDocContent(this, isSelection);
	};
	AscWord.ParagraphContentBase.prototype.visit = function(visitor, isSelection)
	{
	};
	AscWord.ParagraphContentWithParagraphLikeContent.prototype.visit = function(visitor, isSelection)
	{
		if (isSelection && !this.Selection.Use)
			return;
		
		let start = isSelection ? Math.min(this.Selection.StartPos, this.Selection.EndPos) : 0;
		let end   = isSelection ? Math.max(this.Selection.StartPos, this.Selection.EndPos) : this.Content.length - 1;
		for (let i = start; i <= end; ++i)
		{
			this.Content[i].visit(visitor, isSelection);
			if (visitor.isStopped())
				return;
		}
	};
	AscWord.InlineLevelSdt.prototype.visit = function(visitor, isSelection)
	{
		if (isSelection && !this.Selection.Use)
			return;
		
		let isPartial = isSelection && !this.IsPlaceHolder() && !this.IsSelectedAll();
		if (isPartial && this.IsSelectionEmpty(true))
			return;
		
		let start = isPartial ? Math.min(this.Selection.StartPos, this.Selection.EndPos) : undefined;
		let end   = isPartial ? Math.max(this.Selection.StartPos, this.Selection.EndPos) : undefined;
		
		let isForm = this.IsForm();
		if (isForm && visitor.oform(this, true, start, end))
			return;
		
		if (!isForm && visitor.inlineLevelSdt(this, true, start, end))
			return;
		
		if (visitor.isStopped())
			return;
		
		AscWord.ParagraphContentWithParagraphLikeContent.prototype.visit.call(this, visitor, isPartial);
		if (visitor.isStopped())
			return;
		
		if (isForm)
			visitor.oform(this, false, start, end);
		else
			visitor.inlineLevelSdt(this, false, start, end);
	};
	AscWord.Run.prototype.visit = function(visitor, isSelection)
	{
		let start = undefined;
		let end   = undefined;
		if (isSelection)
		{
			if (!this.Selection.Use)
				return;
			
			start = Math.min(this.Selection.StartPos, this.Selection.EndPos);
			end   = Math.max(this.Selection.StartPos, this.Selection.EndPos);
			if (start >= end)
				return;
			
			if (0 === start && this.Content.length === end)
			{
				start = undefined;
				end   = undefined;
			}
		}
		
		if (visitor.run(this, true, start, end) || visitor.isStopped())
			return;
		
		// TODO: inner elements
		
		visitor.run(this, false, start, end);
	};
	AscWord.FldSimple.prototype.visit = function(visitor, isSelection)
	{
		if (isSelection && !this.Selection.Use)
			return;
		
		let isPartial = isSelection && !this.IsSelectedAll();
		if (isPartial && this.IsSelectionEmpty(true))
			return;
		
		let start = isPartial ? Math.min(this.Selection.StartPos, this.Selection.EndPos) : undefined;
		let end   = isPartial ? Math.max(this.Selection.StartPos, this.Selection.EndPos) : undefined;
		if (visitor.fldSimple(this, true, start, end) || visitor.isStopped())
			return;
		
		AscWord.ParagraphContentWithParagraphLikeContent.prototype.visit.call(this, visitor, isPartial);
		if (visitor.isStopped())
			return;
		
		visitor.fldSimple(this, false, start, end);
	};
	AscWord.ParaMath.prototype.visit = function(visitor, isSelection)
	{
		if (isSelection && !this.IsSelectionUse())
			return;
		
		let root      = this.Root;
		let isPartial = isSelection && !this.IsSelectedAll();
		if (isPartial && this.IsSelectionEmpty())
			return;
		
		let start = isPartial ? Math.min(root.Selection.StartPos, root.Selection.EndPos) : undefined;
		let end   = isPartial ? Math.max(root.Selection.StartPos, root.Selection.EndPos) : undefined;
		if (visitor.oMath(this, true, start, end) || visitor.isStopped())
			return;
		
		root.visit(visitor, isPartial);
		if (visitor.isStopped())
			return;
		
		visitor.oMath(this, false, start, end);
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.DocumentVisitor = DocumentVisitor;
	
	
	// Example of usage
	//
	// function DocumentText()
	// {
	// 	DocumentVisitor.call(this);
	// 	this.text = "";
	// }
	// DocumentText.prototype = Object.create(DocumentVisitor.prototype);
	// DocumentText.prototype.constructor = DocumentText;
	// DocumentText.prototype.run = function(run)
	// {
	// 	this.text += run.GetText();
	// 	return true;
	// };
	// DocumentText.prototype.oform = function(form)
	// {
	// 	if (form.IsCheckBox())
	// 		this.text += "<form key=" + form.GetFormKey() + " checked=" + form.IsCheckBoxChecked() + ">";
	// 	else
	// 		this.text += "<form key=" + form.GetFormKey() + ">";
	// 	return true;
	// };
	//
	// AscWord.DocumentText = DocumentText;
	
	
	// TODO: Move to the proper place
	/**
	 * Check if header/footer have fields than need to be recalculated across different pages
	 * @constructor
	 */
	function HdrFtrFieldChecker()
	{
		DocumentVisitor.call(this);
		
		this.haveField = false;
	}
	HdrFtrFieldChecker.check = function(docContent)
	{
		let checker = new HdrFtrFieldChecker();
		checker.traverseDocContent(docContent);
		return checker.haveField;
	};
	HdrFtrFieldChecker.prototype = Object.create(DocumentVisitor.prototype);
	HdrFtrFieldChecker.prototype.constructor = HdrFtrFieldChecker;
	HdrFtrFieldChecker.prototype.run = function(run)
	{
		for (let i = 0; i < run.GetElementsCount(); ++i)
		{
			let item = run.GetElement(i);
			if (item.IsFieldChar() && item.IsBegin())
			{
				let complexField = item.GetComplexField();
				let instruction  = complexField ? complexField.GetInstruction() : null;
				if (instruction)
					this.checkFieldType(instruction.GetType());
			}
			else if (para_PageCount === item.Type || para_PageNum === item.Type)
			{
				this.stop();
				this.haveField = true;
				return true;
			}
		}
		return false;
	};
	HdrFtrFieldChecker.prototype.fldSimple = function(field, isStart)
	{
		if (!isStart)
			return false;
		
		this.checkFieldType(field.GetFieldType());
		return false;
	};
	HdrFtrFieldChecker.prototype.checkFieldType = function(type)
	{
		if (AscWord.fieldtype_PAGE === type || AscWord.fieldtype_NUMPAGES === type)
		{
			this.haveField = true;
			this.stop();
		}
	};
	AscWord.HdrFtrFieldChecker = HdrFtrFieldChecker;
	
})(window);

