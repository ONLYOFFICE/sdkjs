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
	 * Class for traversing the presentation structure: slides, layouts, masters, notes, drawings and their text
	 * @constructor
	 */
	function PresentationVisitor()
	{
		AscWord.DocumentVisitor.call(this);
	}
	PresentationVisitor.prototype = Object.create(AscWord.DocumentVisitor.prototype);
	PresentationVisitor.prototype.constructor = PresentationVisitor;

	PresentationVisitor.prototype.slide = function(slide, isStart)
	{
		return false;
	};
	PresentationVisitor.prototype.layout = function(layout, isStart)
	{
		return false;
	};
	PresentationVisitor.prototype.master = function(master, isStart)
	{
		return false;
	};
	PresentationVisitor.prototype.notes = function(notes, isStart)
	{
		return false;
	};
	PresentationVisitor.prototype.drawing = function(drawing, isStart)
	{
		return false;
	};

	PresentationVisitor.prototype.traversePresentation = function(presentation, options)
	{
		options = options || {};

		let slides = presentation.Slides;
		for (let i = 0; i < slides.length; ++i)
		{
			this.traverseSlide(slides[i], options.notes);
			if (this.isStopped())
				return;
		}

		let masters = presentation.slideMasters;
		for (let i = 0; i < masters.length; ++i)
		{
			if (options.masters)
				this.traverseMaster(masters[i], options.layouts);
			else if (options.layouts)
				this.visitLayouts(masters[i]);

			if (this.isStopped())
				return;
		}
	};
	PresentationVisitor.prototype.traverseSlide = function(slide, withNotes)
	{
		if (this.slide(slide, true))
			return;

		this.visitSpTree(slide.cSld.spTree);
		if (this.isStopped())
			return;

		if (withNotes && slide.notes)
		{
			this.traverseNotes(slide.notes);
			if (this.isStopped())
				return;
		}

		this.slide(slide, false);
	};
	PresentationVisitor.prototype.traverseMaster = function(master, withLayouts)
	{
		if (this.master(master, true))
			return;

		this.visitSpTree(master.cSld.spTree);
		if (this.isStopped())
			return;

		if (withLayouts)
		{
			this.visitLayouts(master);
			if (this.isStopped())
				return;
		}

		this.master(master, false);
	};
	PresentationVisitor.prototype.traverseLayout = function(layout)
	{
		if (this.layout(layout, true))
			return;

		this.visitSpTree(layout.cSld.spTree);
		if (this.isStopped())
			return;

		this.layout(layout, false);
	};
	PresentationVisitor.prototype.traverseNotes = function(notes)
	{
		if (this.notes(notes, true))
			return;

		let bodyShape = notes.getBodyShape();
		if (bodyShape)
		{
			bodyShape.visit(this);
			if (this.isStopped())
				return;
		}

		this.notes(notes, false);
	};
	PresentationVisitor.prototype.traverseDrawing = function(drawing)
	{
		drawing.visit(this);
	};
	PresentationVisitor.prototype.traverseSelection = function(presentation, withNotes)
	{
		let controller = presentation.GetCurrentController();
		let selected   = controller ? controller.selectedObjects : [];
		if (controller && controller.selection.groupSelection)
			selected = controller.selection.groupSelection.selectedObjects;

		if (!selected.length)
		{
			let slideObjects = presentation.GetSelectedSlideObjects();
			for (let i = 0; i < slideObjects.length; ++i)
			{
				this.traverseSlideObject(slideObjects[i], withNotes);
				if (this.isStopped())
					return;
			}
			return;
		}

		let textObject = AscFormat.getTargetTextObject(controller);
		if (!textObject)
		{
			for (let i = 0; i < selected.length; ++i)
			{
				selected[i].visit(this);
				if (this.isStopped())
					return;
			}
			return;
		}

		if (AscDFH.historyitem_type_GraphicFrame === textObject.getObjectType() && textObject.graphicObject.IsCellSelection())
		{
			if (this.drawing(textObject, true))
				return;

			let table = textObject.graphicObject;
			let cells = table.GetSelectionArray();
			for (let i = 0; i < cells.length; ++i)
			{
				table.GetRow(cells[i].Row).GetCell(cells[i].Cell).visit(this);
				if (this.isStopped())
					return;
			}

			this.drawing(textObject, false);
			return;
		}

		let docContent = controller.getTargetDocContent();
		if (!docContent)
			return;

		let isDrawing = textObject instanceof AscFormat.CGraphicObjectBase;
		if (isDrawing && this.drawing(textObject, true))
			return;

		this.visitSelectedParagraphs(docContent);
		if (this.isStopped())
			return;

		if (isDrawing)
			this.drawing(textObject, false);
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	PresentationVisitor.prototype.traverseSlideObject = function(slideObject, withNotes)
	{
		switch (slideObject.getObjectType())
		{
			case AscDFH.historyitem_type_Slide:
				return this.traverseSlide(slideObject, withNotes);
			case AscDFH.historyitem_type_SlideMaster:
				return this.traverseMaster(slideObject, false);
			case AscDFH.historyitem_type_SlideLayout:
				return this.traverseLayout(slideObject);
		}
	};
	PresentationVisitor.prototype.visitLayouts = function(master)
	{
		let layouts = master.sldLayoutLst;
		for (let i = 0; i < layouts.length; ++i)
		{
			this.traverseLayout(layouts[i]);
			if (this.isStopped())
				return;
		}
	};
	PresentationVisitor.prototype.visitSpTree = function(spTree)
	{
		for (let i = 0; i < spTree.length; ++i)
		{
			spTree[i].visit(this);
			if (this.isStopped())
				return;
		}
	};
	PresentationVisitor.prototype.visitSelectedParagraphs = function(docContent)
	{
		if (!docContent.IsSelectionUse())
		{
			docContent.visit(this);
			return;
		}

		let start = Math.min(docContent.Selection.StartPos, docContent.Selection.EndPos);
		let end   = Math.max(docContent.Selection.StartPos, docContent.Selection.EndPos);
		for (let i = start; i <= end; ++i)
		{
			docContent.Content[i].visit(this);
			if (this.isStopped())
				return;
		}
	};
	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	AscFormat.CGraphicObjectBase.prototype.visit = function(visitor)
	{
		if (visitor.drawing(this, true))
			return;

		this.visitContent(visitor);
		if (visitor.isStopped())
			return;

		visitor.drawing(this, false);
	};
	AscFormat.CGraphicObjectBase.prototype.visitContent = function(visitor)
	{
	};
	AscFormat.CShape.prototype.visitContent = function(visitor)
	{
		let docContent = this.getDocContent();
		if (docContent)
			docContent.visit(visitor);
	};
	AscFormat.CGroupShape.prototype.visitContent = function(visitor)
	{
		visitor.visitSpTree(this.spTree);
	};
	AscFormat.CGraphicFrame.prototype.visitContent = function(visitor)
	{
		if (this.graphicObject)
			this.graphicObject.visit(visitor);
	};
	AscFormat.CChartSpace.prototype.visitContent = function(visitor)
	{
		let titles = this.getAllTitles();
		for (let i = 0; i < titles.length; ++i)
		{
			let title = titles[i];
			if (!title.tx || !title.tx.rich)
				continue;

			let docContent = title.getDocContent();
			if (docContent)
				docContent.visit(visitor);

			if (visitor.isStopped())
				return;
		}
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommonSlide'] = window['AscCommonSlide'] || {};
	window['AscCommonSlide'].PresentationVisitor = PresentationVisitor;
})(window);
