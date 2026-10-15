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

(function () {
	function SlidesCollector(presentation, sourceFormattingContent, endFormattingContent, imagesSelectedContent) {
		this.presentation = presentation;
		this.sourceFormattingContent = sourceFormattingContent;
		this.endFormattingContent = endFormattingContent;
		this.imagesSelectedContent = imagesSelectedContent;

		this.mastersMap = {};
	}

	SlidesCollector.prototype.addMaster = function (master) {
		this.sourceFormattingContent.Masters.push(master);
		const layouts = master.sldLayoutLst;
		for (var i = 0; i < layouts.length; ++i) {
			this.sourceFormattingContent.Layouts.push(layouts[i]);
		}
		this.sourceFormattingContent.Themes.push(master.Theme);
	};
	SlidesCollector.prototype.fillEndFormatting = function () {
		const sourceFormattingContent = this.sourceFormattingContent;
		const endFormattingContent = this.endFormattingContent;
		endFormattingContent.Layouts = sourceFormattingContent.Layouts;
		endFormattingContent.LayoutsIndexes = sourceFormattingContent.LayoutsIndexes;
		endFormattingContent.Masters = sourceFormattingContent.Masters;
		endFormattingContent.MastersIndexes = sourceFormattingContent.MastersIndexes;
		endFormattingContent.Themes = sourceFormattingContent.Themes;
		endFormattingContent.ThemesIndexes = sourceFormattingContent.ThemesIndexes;
		endFormattingContent.Notes = sourceFormattingContent.Notes;
		endFormattingContent.NotesMasters = sourceFormattingContent.NotesMasters;
		endFormattingContent.NotesMastersIndexes = sourceFormattingContent.NotesMastersIndexes;
		endFormattingContent.NotesThemes = sourceFormattingContent.NotesThemes;
	};
	SlidesCollector.prototype.fillIndexes = function () {
		const sourceFormattingContent = this.sourceFormattingContent;
		const slides = sourceFormattingContent.SlideObjects;
		const layouts = sourceFormattingContent.Layouts;
		const masters = sourceFormattingContent.Masters;
		const notes = sourceFormattingContent.Notes;
		const notesMasters = sourceFormattingContent.NotesMasters;

		const layoutIndexesMap = {};
		for (let i = 0; i < layouts.length; i += 1) {
			const layout = layouts[i];
			layoutIndexesMap[layout.GetId()] = i;
		}
		for (let i = 0; i < slides.length; i += 1) {
			const layout = slides[i].Layout;
			sourceFormattingContent.LayoutsIndexes.push(layoutIndexesMap[layout.GetId()]);
		}

		const masterIndexesMap = {};
		for (let i = 0; i < masters.length; i += 1) {
			const master = masters[i];
			masterIndexesMap[master.GetId()] = i;
		}
		for (let i = 0; i < layouts.length; i += 1) {
			const master = layouts[i].Master;
			sourceFormattingContent.MastersIndexes.push(masterIndexesMap[master.GetId()]);
		}

		for (let i = 0; i < masters.length; i += 1) {
			sourceFormattingContent.ThemesIndexes.push(i);
		}

		const notesMastersIndexesMap = {};
		for (let i = 0; i < notesMasters.length; i += 1) {
			const noteMaster = notesMasters[i];
			notesMastersIndexesMap[noteMaster.GetId()] = i;
		}
		for (let i = 0; i < notes.length; i += 1) {
			const note = notes[i];
			if (!note) {
				sourceFormattingContent.NotesMastersIndexes.push(-1);
				continue;
			}
			const master = note.Master;
			sourceFormattingContent.NotesMastersIndexes.push(notesMastersIndexesMap[master.GetId()]);
		}
	};
	SlidesCollector.prototype.collect = function (slides) {
	};


	function NormalSlidesCollector(presentation, sourceFormattingContent, endFormattingContent, imagesSelectedContent) {
		SlidesCollector.call(this, presentation, sourceFormattingContent, endFormattingContent, imagesSelectedContent);
		this.notesMastersMap = {};
	}

	AscFormat.InitClassWithoutType(NormalSlidesCollector, SlidesCollector);

	NormalSlidesCollector.prototype.addSlideNotes = function (slide) {
		const sourceFormattingContent = this.sourceFormattingContent;
		const note = slide.notes;
		sourceFormattingContent.Notes.push(note);
		if (!note) {
			return;
		}
		const master = note.Master;
		if (!this.notesMastersMap[master.GetId()]) {
			this.notesMastersMap[master.GetId()] = master;
			sourceFormattingContent.NotesMasters.push(master);
			sourceFormattingContent.NotesThemes.push(master.Theme);
		}
	};
	NormalSlidesCollector.prototype.addSlide = function (slide) {
		this.sourceFormattingContent.SlideObjects.push(slide);
		let endFormattingSlide = slide;
		if (endFormattingSlide.cSld && endFormattingSlide.cSld.Bg) {
			endFormattingSlide = endFormattingSlide.createDuplicate();
			endFormattingSlide.changeBackground(null);
		}
		this.endFormattingContent.SlideObjects.push(endFormattingSlide);
	}

	NormalSlidesCollector.prototype.collect = function (selectedSlides) {
		const presentation = this.presentation;
		for (let i = 0; i < selectedSlides.length; ++i) {
			const slide = selectedSlides[i];
			const master = slide.Layout.Master;
			const masterId = master.Get_Id();
			if (!this.mastersMap[masterId]) {
				this.mastersMap[masterId] = master;
				this.addMaster(master);
			}
			this.addSlide(slide);
			if (i === 0) {
				var rasterImageId = slide.getBase64Img();
				var image = AscFormat.DrawingObjectsController.prototype.createImage(rasterImageId, 0, 0, presentation.GetWidthMM() / 2.0, presentation.GetHeightMM() / 2.0);
				this.imagesSelectedContent.Drawings.push(new DrawingCopyObject(image, 0, 0, presentation.GetWidthMM() / 2.0, presentation.GetHeightMM() / 2.0, rasterImageId));
			}
			this.addSlideNotes(slide);
		}
		this.fillIndexes();
		this.fillEndFormatting();
	};

	function MasterSlideCollector(presentation, sourceFormattingContent, endFormattingContent, imagesSelectedContent) {
		SlidesCollector.call(this, presentation, sourceFormattingContent, endFormattingContent, imagesSelectedContent);
	}

	AscFormat.InitClassWithoutType(MasterSlideCollector, SlidesCollector);
	MasterSlideCollector.prototype.collectMasters = function (selectedSlides) {
		const presentation = this.presentation;
		for (let i = 0; i < selectedSlides.length; ++i) {
			const slide = selectedSlides[i];
			if (slide.isSlideMaster()) {
				this.addMaster(slide);
			}
			if (i === 0) {
				const rasterImageId = slide.getBase64Img();
				const image = AscFormat.DrawingObjectsController.prototype.createImage(rasterImageId, 0, 0, presentation.GetWidthMM() / 2.0, presentation.GetHeightMM() / 2.0);
				this.imagesSelectedContent.Drawings.push(new DrawingCopyObject(image, 0, 0, presentation.GetWidthMM() / 2.0, presentation.GetHeightMM() / 2.0, rasterImageId));
			}
		}
	};
	MasterSlideCollector.prototype.collectLayouts = function (selectedSlides) {
		const presentation = this.presentation;
		for (let i = 0; i < selectedSlides.length; ++i) {
			const slide = selectedSlides[i];
			if (slide.isLayout()) {
				this.sourceFormattingContent.Layouts.push(slide);
			}
			if (i === 0) {
				let rasterImageId = slide.getBase64Img();
				const image = AscFormat.DrawingObjectsController.prototype.createImage(rasterImageId, 0, 0, presentation.GetWidthMM() / 2.0, presentation.GetHeightMM() / 2.0);
				this.imagesSelectedContent.Drawings.push(new DrawingCopyObject(image, 0, 0, presentation.GetWidthMM() / 2.0, presentation.GetHeightMM() / 2.0, rasterImageId));
			}
		}
	};
	MasterSlideCollector.prototype.collect = function (selectedSlides) {
		if (selectedSlides.length > 0) {
			const firstSlide = selectedSlides[0];
			if (firstSlide.isSlideMaster()) {
				this.collectMasters(selectedSlides);
			} else {
				this.collectLayouts(selectedSlides);
			}
			this.fillIndexes();
			this.fillEndFormatting();
		}
	};

	function collectSelectedObjects(spTree, collectArray, isRecursive, idMap, isSourceFormatting) {
		const props = new AscFormat.CCopyObjectProperties();
		props.idMap = idMap;
		props.bSaveSourceFormatting = isSourceFormatting;
		for (let i = 0; i < spTree.length; ++i) {
			const shape = spTree[i];
			if (shape.selected) {
				let copy;
				if (shape.isGroupObject()) {
					copy = shape.copy(props);
					copy.setParent(shape.parent);
				} else {
					if (!isSourceFormatting) {
						copy = shape.copy(props);
						copy.setParent(shape.parent);
						if (shape.isPlaceholder && shape.isPlaceholder()) {
							copy.x = shape.x;
							copy.y = shape.y;
							copy.extX = shape.extX;
							copy.extY = shape.extY;
							copy.rot = shape.rot;
							AscFormat.CheckSpPrXfrm(copy, true);
						}
						copy.convertFromSmartArt(true);
					} else {
						copy = shape.getCopyWithSourceFormatting();
						copy.convertFromSmartArt(true);
						copy.setParent(shape.parent);
					}

				}

				collectArray.push(new DrawingCopyObject(copy, shape.x, shape.y, shape.extX, shape.extY, shape.getBase64Img()));
				if (AscCommon.isRealObject(idMap)) {
					idMap[shape.Id] = copy.Id;
				}
			}
			if (isRecursive && shape.isGroupObject()) {
				collectSelectedObjects(shape.spTree, collectArray, isRecursive, idMap, isSourceFormatting);
			}
		}
	}

	window["AscCommonSlide"] = window["AscCommonSlide"] || {};
	window["AscCommonSlide"].MasterSlideCollector = MasterSlideCollector;
	window["AscCommonSlide"].NormalSlidesCollector = NormalSlidesCollector;
	window["AscCommonSlide"].NormalSlidesCollector = NormalSlidesCollector;
	window["AscCommonSlide"].collectSelectedObjects = collectSelectedObjects;
})();
