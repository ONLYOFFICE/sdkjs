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
		this.visitDocContent(content);
	};
	/**
	 * @param {AscWord.Document} logicDocument
	 * @param {boolean} isSelection
	 */
	DocumentVisitor.prototype.traverseDocument = function(logicDocument, isSelection)
	{
		if (isSelection)
		{
			let selectedContent = logicDocument.GetSelectedContent(false);
			this.traverseSelectedContent(selectedContent)
		}
		else
		{
			this.visitDocContent(logicDocument.Content);
		}
	};
	DocumentVisitor.prototype.traverseDocContent = function(docContent)
	{
		this.visitDocContent(docContent.Content);
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
	DocumentVisitor.prototype.paragraph = function(paragraph, isStart)
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
	DocumentVisitor.prototype.tableCell = function(tableCell, isStart)
	{
		return false;
	};
	DocumentVisitor.prototype.fldSimple = function(field, isStart)
	{
		return false;
	};
	DocumentVisitor.prototype.blockLevelSdt = function(sdt, isStart)
	{
		return false;
	};
	DocumentVisitor.prototype.inlineLevelSdt = function(sdt, isStart)
	{
		return false;
	};
	DocumentVisitor.prototype.oform = function(form, isStart)
	{
		return false;
	};
	DocumentVisitor.prototype.run = function(run, isStart)
	{
		return false;
	};
	DocumentVisitor.prototype.oMath = function(oMath, isStart)
	{
		return false;
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	DocumentVisitor.prototype.visitDocContent = function(content)
	{
		for (let i = 0, count = content.length; i < count; ++i)
		{
			content[i].visit(this);
			if (this.isStopped())
				return;
		}
	};
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	AscWord.Paragraph.prototype.visit = function(visitor)
	{
		if (visitor.paragraph(this, true))
			return;

		for (let i = 0; i < this.Content.length; ++i)
		{
			this.Content[i].visit(visitor);
			if (visitor.isStopped())
				return;
		}
		
		visitor.paragraph(this, false);
	};
	AscWord.Table.prototype.visit = function(visitor)
	{
		if (visitor.table(this, true))
			return;

		for (let i = 0, rowCount = this.GetRowsCount(); i < rowCount; ++i)
		{
			this.GetRow(i).visit(visitor);
			if (visitor.isStopped())
				return;
		}
		
		visitor.table(this, false);
	};
	AscWord.TableRow.prototype.visit = function(visitor)
	{
		if (visitor.tableRow(this, true))
			return;
		
		for (let i = 0, cellCount = this.GetCellsCount(); i < cellCount; ++i)
		{
			let cell = this.GetCell(i);
			cell.visit(visitor);
			if (visitor.isStopped())
				return;
		}
		
		visitor.tableRow(this, false);
	};
	AscWord.TableCell.prototype.visit = function(visitor)
	{
		if (visitor.tableCell(this, true))
			return;
		
		this.GetContent().visit(visitor);
		
		visitor.tableCell(this, false);
	};
	AscWord.BlockLevelSdt.prototype.visit = function(visitor)
	{
		if (visitor.blockLevelSdt(this, true))
			return;
		
		this.Content.visit(visitor);
		
		visitor.blockLevelSdt(this, false);
	};
	AscWord.DocumentContent.prototype.visit = function(visitor)
	{
		visitor.visitDocContent(this.Content);
	};
	AscWord.ParagraphContentBase.prototype.visit = function(visitor)
	{
	};
	AscWord.ParagraphContentWithParagraphLikeContent.prototype.visit = function(visitor)
	{
		for (let i = 0; i < this.Content.length; ++i)
		{
			this.Content[i].visit(visitor);
			if (visitor.isStopped())
				return;
		}
	};
	AscWord.InlineLevelSdt.prototype.visit = function(visitor)
	{
		let isForm = this.IsForm();
		if (isForm && visitor.oform(this, true))
			return;
		
		if (!isForm && visitor.inlineLevelSdt(this, true))
			return;
		
		AscWord.ParagraphContentWithParagraphLikeContent.prototype.visit.apply(this, arguments);
		
		if (isForm)
			visitor.oform(this, false);
		else
			visitor.inlineLevelSdt(this, false);
	};
	AscWord.Run.prototype.visit = function(visitor)
	{
		if (visitor.run(this, true))
			return;
		
		// TODO: inner elements
		
		visitor.run(this, false);
	};
	AscWord.FldSimple.prototype.visit = function(visitor)
	{
		if (visitor.fldSimple(this, true))
			return;
		
		AscWord.ParagraphContentWithParagraphLikeContent.prototype.visit.apply(this, arguments);
		
		visitor.fldSimple(this, false);
	};
	AscWord.ParaMath.prototype.visit = function(visitor)
	{
		if (visitor.oMath(this, true))
			return;
		
		AscWord.ParagraphContentWithParagraphLikeContent.prototype.visit.apply(this.Root, arguments);
		
		visitor.oMath(this, false);
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

