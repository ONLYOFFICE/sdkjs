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

(function(undefined) {
	function SlideModeManagerBase(api) {
		this.api = api;
		this.type = null;
	}
	SlideModeManagerBase.prototype.updatePlaceholderCheckboxState = function (placeholderEvent, slide, placeholderTypes) {
		let isMatching = true;
		if (slide) {
			for (let idx = 0; idx < placeholderTypes.length; ++idx) {
				const type = placeholderTypes[idx];
				const shape = slide.getMatchingShape(type, null, false, {});
				if (!shape) {
					isMatching = false;
					break;
				}
			}
		} else {
			isMatching = false;
		}
		const api = this.getApi();
		api.sendEvent(placeholderEvent, isMatching);
	};
	SlideModeManagerBase.prototype.startAction = function(description) {
		const presentation = this.getPresentation();
		presentation.StartAction(description);
	};
	SlideModeManagerBase.prototype.finalizeAction = function(checkEmptyAction) {
		const presentation = this.getPresentation();
		presentation.FinalizeAction(checkEmptyAction);
	};
	SlideModeManagerBase.prototype.getPresentation = function() {
		return this.api.WordControl.m_oLogicDocument;
	};
	SlideModeManagerBase.prototype.getWordControl = function() {
		return this.api.WordControl;
	};
	SlideModeManagerBase.prototype.getApi = function() {
		return this.api;
	};
	SlideModeManagerBase.prototype.getCurrentSlide = function() {
		const presentation = this.getPresentation();
		return this.getAllSlides()[presentation.CurPage];
	};
	SlideModeManagerBase.prototype.getSelectedSlides = function() {
		const wordControl = this.getWordControl();
		let selectedArray = [];
		if (!window["NATIVE_EDITOR_ENJINE"] && wordControl.Thumbnails) {
			selectedArray = wordControl.Thumbnails.GetSelectedArray();
		}

		if (!selectedArray.length) {
			const presentation = this.getPresentation();
			if (presentation.CurPage >= 0) {
				selectedArray = [presentation.CurPage];
			}
		}
		return selectedArray;
	};
	SlideModeManagerBase.prototype.getSlideIndex = function(object) {
		const allSlides = this.getAllSlides();
		for (let i = 0; i < allSlides.length; i += 1) {
			if (allSlides[i] === object) {
				return i;
			}
		}
		return -1;
	};
	SlideModeManagerBase.prototype.getSlide = function(idx) {
		const allSlides = this.getAllSlides();
		return allSlides[idx] || null;
	};
	SlideModeManagerBase.prototype.isMasterSlideMode = function() {
		return false;
	};
	SlideModeManagerBase.prototype.isSlideMode = function() {
		return false;
	};
	SlideModeManagerBase.prototype.isNoteMode = function() {
		return false;
	};
	SlideModeManagerBase.prototype.isMasterNoteMode = function() {
		return false;
	};
	SlideModeManagerBase.prototype.isMasterHandoutMode = function() {
		return false;
	};
	SlideModeManagerBase.prototype.isSorterMode = function() {
		return false;
	};
	SlideModeManagerBase.prototype.isOutlineMode = function() {
		return false;
	};
	SlideModeManagerBase.prototype.getSlidesCount = function() {
		return this.getAllSlides().length;
	};
	SlideModeManagerBase.prototype.getAllSlides = function() {
		return [];
	};
	SlideModeManagerBase.prototype.updateViewMode = function(oldViewManager) {
	};
	SlideModeManagerBase.prototype.setSpecialPasteShowOptions = function(props) {
	};
	SlideModeManagerBase.prototype.isCanAddHyperlinkInContent = function() {
		return false;
	};
	SlideModeManagerBase.prototype.isMasterPlaceholderShape = function(shape) {
		return false;
	};
	SlideModeManagerBase.prototype.addMasterSlide = function() {
	};
	SlideModeManagerBase.prototype.addLayoutSlide = function() {
	};
	SlideModeManagerBase.prototype.startAddPlaceholder = function(nType, bVertical, bStart) {
	};
	SlideModeManagerBase.prototype.applySlideTransition = function(oTransition) {
	};
	SlideModeManagerBase.prototype.applySlideTransitionToAll = function() {
	};
	SlideModeManagerBase.prototype.getCumulativeThumbnailsLength = function(isHorizontalOrientation, thumbnailWidth, thumbnailHeight) {
		return 0;
	};
	SlideModeManagerBase.prototype.isNotesSupported = function() {
		return false;
	};
	SlideModeManagerBase.prototype.isThumbnailsSupported = function() {
		return false;
	};
	SlideModeManagerBase.prototype.getAnimationStartSlideNum = function(startSlideNum) {
		return 0;
	};
	SlideModeManagerBase.prototype.getSavedAnimationStartObject = function() {
		return null;
	};
	SlideModeManagerBase.prototype.goToSavedAnimationStartObject = function(object) {
	};
	SlideModeManagerBase.prototype.getSlideNumber = function(idx) {
		return null;
	};
	SlideModeManagerBase.prototype.insertSlideObjectToPos = function(pos, slide) {
	};
	SlideModeManagerBase.prototype.setTitle = function() {};
	SlideModeManagerBase.prototype.setFooter = function() {};
	SlideModeManagerBase.prototype.setHeader = function() {};
	SlideModeManagerBase.prototype.setDate = function() {};
	SlideModeManagerBase.prototype.setPageNumber = function() {};
	SlideModeManagerBase.prototype.setSlideImage = function() {};
	SlideModeManagerBase.prototype.setBody = function() {};

	SlideModeManagerBase.prototype.setHandoutPageCount = function(val) {
	};
	SlideModeManagerBase.prototype.setPageOrientation = function(val) {
	};
	SlideModeManagerBase.prototype.getCurrentTheme = function() {
		return null;
	};
	SlideModeManagerBase.prototype.getSizesMM = function() {
		return {width: 0, height: 0};
	};
	SlideModeManagerBase.prototype.getSlidesForChangeColorScheme = function() {
		return this.getAllSlides();
	};
	SlideModeManagerBase.prototype.isCanShiftSlides = function() {
		return this.isThumbnailsSupported()
	};
	SlideModeManagerBase.prototype.shiftSlides = function() {
		return []
	};
	SlideModeManagerBase.prototype.applyHFProps = function(props, defaultLang) {
		return false;
	};
	SlideModeManagerBase.prototype.updateInterfaceState = function() {
	};
	SlideModeManagerBase.prototype.isSlideNoteShape = function(shape) {
		return false;
	};
	SlideModeManagerBase.prototype.recalculateThemeObjects = function() {};
	SlideModeManagerBase.prototype.isDrawingSlide = function () {
		return false;
	};
	SlideModeManagerBase.prototype.isFocusOnOutline = function () {
		return false;
	};
	SlideModeManagerBase.prototype.addNextSlide = function (layoutIndex) {};
	SlideModeManagerBase.prototype.changeTheme = function (themeInfo, arrInd) {};
	SlideModeManagerBase.prototype.insertContent2 = function (oContent, aContents, kw, kh, bEndFormatting) {
		return false;
	};
	SlideModeManagerBase.prototype.addNewMasterSlide = function () {};
	SlideModeManagerBase.prototype.addNewLayout = function () {};
	SlideModeManagerBase.prototype.deleteMaster = function () {};
	SlideModeManagerBase.prototype.setPreserve = function (bPr) {};
	SlideModeManagerBase.prototype.recalculateNotesShape = function (slide) {

	};
	SlideModeManagerBase.prototype.applySlideBackgroundToAll = function () {

	};
	SlideModeManagerBase.prototype.isLockBackgroundSlides = function () {
		return true;
	};
	SlideModeManagerBase.prototype.isCanAddHyperlink = function () {
		return false;
	};

	function SlideModeManager(api) {
		SlideModeManagerBase.call(this, api);
		this.type = Asc.c_oAscPresentationViewMode.normal;
	}

	AscFormat.InitClassWithoutType(SlideModeManager, SlideModeManagerBase);
	SlideModeManager.prototype.getChangeModePage = function (oldViewManager) {
		const presentation = this.getPresentation();
		let slide = oldViewManager.getCurrentSlide();
		if (slide) {
			if (slide.isSlide()) {
				return slide.num;
			}
			if (slide.isNote()) {
				const noteIndex = presentation.notes.indexOf(slide);
				return noteIndex !== -1 ? noteIndex : 0;
			}
		}
		return 0;
	};
	SlideModeManager.prototype.updateViewMode = function(oldViewManager) {
		const wordControl = this.getWordControl();
		const presentation = this.getPresentation();
		const slideIdx = this.getChangeModePage(oldViewManager);
		const api = this.getApi();
		wordControl.setNotesEnable(true);
		wordControl.setAnimPaneEnable(true);
		wordControl.setThumbnailsEnable(true);
		wordControl.SetThumbnails(false);
		wordControl.GoToPage(slideIdx);
		api.hideMediaControl();
		api.asc_hideComments();
		presentation.Recalculate({Drawings: {All: true, Map: {}}});
		presentation.Document_UpdateInterfaceState();
	};
	SlideModeManager.prototype.isSlideMode = function() {
		return true;
	};
	SlideModeManager.prototype.setSpecialPasteShowOptions = function(props) {
		const presentation = this.getPresentation();
		const wordControl = this.getWordControl();
		let stateSelection = presentation.GetSelectionState();
		let curPage = stateSelection.CurPage;
		let pos = presentation.GetTargetPosition();
		props = !props ? [Asc.c_oSpecialPasteProps.sourceformatting, Asc.c_oSpecialPasteProps.keepTextOnly] : props;
		let x, y, w, h;
		if (null === pos) {
			pos = presentation.GetSelectedBounds();
			w = pos.w;
			h = pos.h;
			x = pos.x + w;
			y = pos.y + h;
		} else {
			x = pos.X;
			y = pos.Y;
		}
		let screenPos;
		let bThumbnals = presentation.IsFocusOnThumbnails() && !presentation.IsFocusOnOutline();
		let sSlideId = null;

		let aSelectedSlides = presentation.GetSelectedSlides();
		if (bThumbnals && aSelectedSlides.length > 0) {
			let nSlideIndex = aSelectedSlides[aSelectedSlides.length - 1];
			let oSlide = presentation.GetSlide(nSlideIndex);
			sSlideId = oSlide.Get_Id();
		}
		if (sSlideId) {
			screenPos = wordControl.Thumbnails.getSpecialPasteButtonCoords(sSlideId);
			w = 1;
			h = 1;
		} else {
			screenPos = presentation.DrawingDocument.ConvertCoordsToCursorWR(x, y, curPage);
		}

		let specialPasteShowOptions = window['AscCommon'].g_specialPasteHelper.buttonInfo;
		specialPasteShowOptions.asc_setOptions(props);
		specialPasteShowOptions.asc_setLastSelectedPasteProperty(window['AscCommon'].g_specialPasteHelper.isPasteOptions ? window['AscCommon'].g_specialPasteHelper.specialPasteProps : null);

		let targetDocContent = presentation.Get_TargetDocContent();
		if (targetDocContent && targetDocContent.Id) {
			specialPasteShowOptions.setShapeId(targetDocContent.Id);
		} else {
			specialPasteShowOptions.setShapeId(null);
		}

		let curCoord = new AscCommon.asc_CRect(screenPos.X, screenPos.Y, 0, 0);
		specialPasteShowOptions.asc_setCellCoord(curCoord);
		specialPasteShowOptions.setFixPosition({x: x, y: y, pageNum: curPage, w: w, h: h, slideId: sSlideId});
	};
	SlideModeManager.prototype.isCanAddHyperlinkInContent = function() {
		return true;
	};
	SlideModeManager.prototype.applySlideTransition = function(oTransition) {
		const presentation = this.getPresentation();
		if (presentation.IsEmpty())
			return;

		if (!presentation.IsSelectionLocked(AscCommon.changestype_SlideTransition)) {
			presentation.StartAction(AscDFH.historydescription_Presentation_ApplyTransition);
			const selectedSlides = presentation.GetSelectedSlides();
			for (let slideIndex = 0; slideIndex < selectedSlides.length; ++slideIndex) {
				let slide = presentation.GetSlide(selectedSlides[slideIndex]);
				if (slide) {
					slide.applyTransition(oTransition);
				}
			}
			const isShowLoop = oTransition.get_ShowLoop();
			if (AscFormat.isRealBool(isShowLoop) && isShowLoop !== presentation.isLoopShowMode()) {
				presentation.setShowLoop(isShowLoop);
			}
			presentation.UpdateInterface();
			presentation.FinalizeAction();
			const api = this.getApi();
			if (oTransition.get_TransitionType() !== undefined || oTransition.get_TransitionOption() !== undefined)
			{
				api.SlideTransitionPlay();
			}
		}
	};
	SlideModeManager.prototype.applySlideTransitionToAll = function() {
		const presentation = this.getPresentation();

		if (presentation.IsEmpty())
			return;

		const curSlide = presentation.GetCurrentSlide();
		if (!curSlide)
			return;

		let transitionToApply = null;
		if (curSlide.transition) {
			transitionToApply = curSlide.transition.createDuplicate();
		}
		if (!transitionToApply)
			return;


		if (!presentation.IsSelectionLocked(AscCommon.changestype_SlideTransition, {All: true})) {
			presentation.StartAction(AscDFH.historydescription_Presentation_ApplyTransitionToAll);
			let slides = presentation.Slides;
			for (let nSld = 0; nSld < slides.length; ++nSld) {
				let slide = slides[nSld];
				if (slide !== curSlide) {
					slide.applyTransition(transitionToApply);
				}
			}
			presentation.UpdateInterface();
			presentation.FinalizeAction();
		}
	};
	SlideModeManager.prototype.getCumulativeThumbnailsLength = function(isHorizontalOrientation, thumbnailWidth, thumbnailHeight) {
		const slidesCount = this.getSlidesCount();
		return isHorizontalOrientation ? thumbnailWidth * slidesCount : thumbnailHeight * slidesCount;
	};
	SlideModeManager.prototype.isNotesSupported = function() {
		return true;
	};
	SlideModeManager.prototype.getAnimationStartSlideNum = function(startSlideNum) {
		return -1 === startSlideNum ? 0 : startSlideNum;
	};
	SlideModeManager.prototype.getAllSlides = function() {
		const presentation = this.getPresentation();
		return presentation.Slides;
	};
	SlideModeManager.prototype.getSlideNumber = function(idx) {
		const presentation = this.getPresentation();
		return idx + presentation.getFirstSlideNumber();
	};
	SlideModeManager.prototype.insertSlideObjectToPos = function(pos, slide) {
		this.getPresentation().insertSlide(pos, slide);
	};
	SlideModeManager.prototype.isThumbnailsSupported = function() {
		return true;
	};
	SlideModeManager.prototype.getCurrentTheme = function() {
		const slide = this.getCurrentSlide();
		return slide && slide.Layout.Master.Theme || null;
	};
	SlideModeManager.prototype.getSizesMM = function() {
		const presentation = this.getPresentation();
		return presentation.GetSlideSizesMM();
	};
	SlideModeManager.prototype.setPageOrientation = function(val) {
		const presentation = this.getPresentation();
		presentation.SetSlidePageOrientation(val);
	};
	SlideModeManager.prototype.addNextSlide = function(layoutIndex) {
		const presentation = this.getPresentation();
		presentation.addForceNextSlide(layoutIndex);
	}
	SlideModeManager.prototype.changeTheme = function(themeInfo, arrInd) {
		const presentation = this.getPresentation();
		let arr_ind;
		if (!Array.isArray(arrInd)) {
			let oCurMaster;
			let oCurSlide = this.getCurrentSlide();
			arr_ind = [];
			if (oCurSlide) {
				oCurMaster = oCurSlide.Layout && oCurSlide.Layout.Master;
				for (let i = 0; i < presentation.Slides.length; ++i) {
					let oSlide = presentation.Slides[i];
					let oMaster = oSlide.Layout && oSlide.Layout.Master;
					if (oMaster === oCurMaster) {
						arr_ind.push(i);
					}
				}
			}
		} else {
			arr_ind = arrInd;
		}
		let i;
		for (i = 0; i < presentation.slideMasters.length; ++i) {
			if (presentation.slideMasters[i] === themeInfo.Master) {
				break;
			}
		}
		if (i === presentation.slideMasters.length) {
			presentation.pushSlideMaster(themeInfo.Master);
		}
		this.replaceTheme(themeInfo, arr_ind);
		presentation.Document_UpdateInterfaceState();
	}
	SlideModeManager.prototype.replaceTheme = function(themeInfo, arr_ind) {
		const presentation = this.getPresentation();
		presentation.clearThemeTimeouts();

		let oCurSlide = this.getCurrentSlide();
		let oParents = oCurSlide.getParentObjects();
		var oldMaster = oParents.master;
		var _new_master = themeInfo.Master;
		_new_master.presentation = presentation;
		const sizes = this.getSizesMM();
		themeInfo.Master.changeSize(sizes.width, sizes.height);
		var oContent, oMasterSp, oMasterContent, oSp;
		if (oldMaster && oldMaster.hf) {
			themeInfo.Master.setHF(oldMaster.hf.createDuplicate());
			if (oldMaster.hf.dt !== false) {
				oMasterSp = oldMaster.getMatchingShape(AscFormat.phType_dt, null, false, {});
				if (oMasterSp) {
					oMasterContent = oMasterSp.getDocContent && oMasterSp.getDocContent();
					if (oMasterContent) {
						oSp = themeInfo.Master.getMatchingShape(AscFormat.phType_dt, null, false, {});
						if (oSp) {
							oContent = oSp.getDocContent && oSp.getDocContent();
							oContent.Copy2(oMasterContent);
						}
						for (let i = 0; i < themeInfo.Master.sldLayoutLst.length; ++i) {
							oSp = themeInfo.Master.sldLayoutLst[i].getMatchingShape(AscFormat.phType_dt, null, false, {});
							if (oSp) {
								oContent = oSp.getDocContent && oSp.getDocContent();
								oContent.Copy2(oMasterContent);
							}
						}
					}
				}
			}
			if (oldMaster.hf.hdr !== false) {
				oMasterSp = oldMaster.getMatchingShape(AscFormat.phType_hdr, null, false, {});
				if (oMasterSp) {
					oMasterContent = oMasterSp.getDocContent && oMasterSp.getDocContent();
					if (oMasterContent) {
						oSp = themeInfo.Master.getMatchingShape(AscFormat.phType_hdr, null, false, {});
						if (oSp) {
							oContent = oSp.getDocContent && oSp.getDocContent();
							oContent.Copy2(oMasterContent);
						}
						for (let i = 0; i < themeInfo.Master.sldLayoutLst.length; ++i) {
							oSp = themeInfo.Master.sldLayoutLst[i].getMatchingShape(AscFormat.phType_hdr, null, false, {});
							if (oSp) {
								oContent = oSp.getDocContent && oSp.getDocContent();
								oContent.Copy2(oMasterContent);
							}
						}
					}
				}
			}
			if (oldMaster.hf.ftr !== false) {
				oMasterSp = oldMaster.getMatchingShape(AscFormat.phType_ftr, null, false, {});
				if (oMasterSp) {
					oMasterContent = oMasterSp.getDocContent && oMasterSp.getDocContent();
					if (oMasterContent) {
						oSp = themeInfo.Master.getMatchingShape(AscFormat.phType_ftr, null, false, {});
						if (oSp) {
							oContent = oSp.getDocContent && oSp.getDocContent();
							oContent.Copy2(oMasterContent);
						}
						for (let i = 0; i < themeInfo.Master.sldLayoutLst.length; ++i) {
							oSp = themeInfo.Master.sldLayoutLst[i].getMatchingShape(AscFormat.phType_ftr, null, false, {});
							if (oSp) {
								oContent = oSp.getDocContent && oSp.getDocContent();
								oContent.Copy2(oMasterContent);
							}
						}
					}
				}
			}
		}
		for (let i = 0; i < themeInfo.Master.sldLayoutLst.length; ++i) {
			themeInfo.Master.sldLayoutLst[i].changeSize(presentation.GetWidthMM(), presentation.GetHeightMM());
		}
		var slides_array = [];
		for (let i = 0; i < arr_ind.length; ++i) {
			slides_array.push(presentation.Slides[arr_ind[i]]);
		}
		let oReplacedMasters = {};
		let aReplacedMasters = [];
		for (let i = 0; i < slides_array.length; ++i) {
			let oSlide = slides_array[i];
			let oOldMaster = oSlide.getMaster();
			if (oOldMaster) {
				if (!oReplacedMasters[oOldMaster.Id]) {
					oReplacedMasters[oOldMaster.Id] = oOldMaster;
					aReplacedMasters.push(oOldMaster);
				}
			}
			presentation.ChangeSlideSlideMaster(slides_array[i], _new_master);
		}

		for (let nMaster = 0; nMaster < aReplacedMasters.length; ++nMaster) {
			let oMaster = aReplacedMasters[nMaster];
			let bFound = false;
			if (oMaster.isPreserve()) {
				continue;
			}
			for (let nSlide = 0; nSlide < presentation.Slides.length; ++nSlide) {
				if (presentation.Slides[nSlide].getMaster() === oMaster) {
					bFound = true;
					break;
				}
			}
			if (!bFound) {
				presentation.removeSlideMasterObject(oMaster);
			}
		}
		presentation.SetTheme(arr_ind);
		presentation.Recalculate();
	};
	SlideModeManager.prototype.insertContent2 = function(oContent, aContents, kw, kh, bEndFormatting) {
		if (oContent.SlideObjects.length === 0) {
			return;
		}
		const bChangeSize = kw !== 1 || kh !== 1;
		const presentation = this.getPresentation();
		const oCurrentMaster = presentation.GetCurrentMaster();
		if (bEndFormatting) {
			const oSourceContent = aContents[1];
			for (let i = 0; i < oContent.SlideObjects.length; ++i) {
				const oSlide = oContent.SlideObjects[i];
				if (bChangeSize) {
					oSlide.Width = oContent.PresentationWidth;
					oSlide.Height = oContent.PresentationHeight;
					oSlide.changeSize(presentation.GetWidthMM(), presentation.GetHeightMM());
				}
				const nLayoutIndex = oSourceContent.LayoutsIndexes[i];
				const oLayout = oSourceContent.Layouts[nLayoutIndex];
				if (oLayout) {
					oSlide.setLayout(oCurrentMaster.getMatchingLayout(oLayout.type, oLayout.matchingName, oLayout.cSld.name, true));
				} else {
					oSlide.setLayout(oCurrentMaster.sldLayoutLst[0]);
				}
				const oNotes = oContent.Notes[i] || AscCommonSlide.CreateNotes(presentation.notesMasters[0]);
				oSlide.setNotes(oNotes);
				oSlide.notes.setNotesMaster(presentation.notesMasters[0]);
				oSlide.notes.setSlide(oSlide);
			}
		} else {
			for (let i = 0; i < oContent.Masters.length; ++i) {
				if (bChangeSize) {
					oContent.Masters[i].scale(kw, kh);
				}
				presentation.pushSlideMaster(oContent.Masters[i]);
			}
			for (let i = 0; i < oContent.Layouts.length; ++i) {
				const oLayout = oContent.Layouts[i];
				if (bChangeSize) {
					oLayout.scale(kw, kh);
				}
				const nMasterIndex = oContent.MastersIndexes[i];
				const oMaster = oContent.Masters[nMasterIndex];
				if (oMaster) {
					oMaster.addLayout(oLayout);
				}

			}
			for (let i = 0; i < oContent.SlideObjects.length; ++i) {
				const oSlide = oContent.SlideObjects[i];
				if (bChangeSize) {
					oSlide.Width = oContent.PresentationWidth;
					oSlide.Height = oContent.PresentationHeight;
					oSlide.changeSize(presentation.GetWidthMM(), presentation.GetHeightMM());
				}
				const nLayoutIndex = oContent.LayoutsIndexes[i];
				const oLayout = oContent.Layouts[nLayoutIndex];
				oSlide.setLayout(oLayout);

				const nMasterIndex = oContent.MastersIndexes[nLayoutIndex];
				const oMaster = oContent.Masters[nMasterIndex];
				const oTheme = oContent.Themes[nMasterIndex];
				const oNotes = oContent.Notes[i];
				const nNotesMasterIndex = oContent.NotesMastersIndexes[i];
				const oNotesMaster = oContent.NotesMasters[nNotesMasterIndex];
				const oNotesTheme = oContent.NotesThemes[nNotesMasterIndex];
				if (!oMaster.Theme) {
					oMaster.setTheme(oTheme);
				}
				if (!oLayout.Master) {
					oLayout.setMaster(oMaster);
				}
				if (!oNotes || !oNotesMaster || !oNotesTheme) {
					oSlide.setNotes(AscCommonSlide.CreateNotes(presentation.notesMasters[0]));
					oSlide.notes.setNotesMaster(presentation.notesMasters[0]);
					oSlide.notes.setSlide(oSlide);
				} else {
					if (!oNotesMaster.Theme) {
						oNotesMaster.setTheme(oNotesTheme);
					}
					if (!oNotes.Master) {
						oNotes.setNotesMaster(oNotesMaster);
					}
					if (!oSlide.notes) {
						oSlide.setNotes(oNotes);
					}
				}
			}
			return true;
		}

		return false;
	}
	SlideModeManager.prototype.recalculateThemeObjects = function(oThemeObjects) {
		const presentation = this.getPresentation();
		for (let nIdx = 0; nIdx < oThemeObjects.masters.length; ++nIdx) {
			oThemeObjects.masters[nIdx].recalculate();
		}
		for (let nIdx = 0; nIdx < oThemeObjects.layouts.length; ++nIdx) {
			oThemeObjects.layouts[nIdx].recalculate();
		}
		let aIdx = [];
		let nStartIdx = 0;
		for (let nIdx = 0; nIdx < oThemeObjects.slides.length; ++nIdx) {
			if (oThemeObjects.slides[nIdx].num === presentation.CurPage) {
				nStartIdx = aIdx.length;
			}
			aIdx.push(oThemeObjects.slides[nIdx].num);
		}
		AscFormat.redrawSlide(presentation.Slides[aIdx[nStartIdx]], presentation, aIdx, nStartIdx, 0, this.getAllSlides());
	};
	SlideModeManager.prototype.shiftSlides = function(pos, array, bCopyOnMove) {
		const presentation = this.getPresentation();
		const aNewSelected = [];
		let deleted = [], i;
		if (!bCopyOnMove) {
			for (i = array.length - 1; i > -1; --i) {
				deleted.push(presentation.removeSlide(array[i], true));
			}

			for (i = 0; i < array.length; ++i) {
				if (array[i] < pos)
					--pos;
				else
					break;
			}
		} else {
			for (i = array.length - 1; i > -1; --i) {
				let oIdMap = {};
				let oSlideCopy = this.getSlide([array[i]]).createDuplicate(oIdMap, false);
				AscFormat.fResetConnectorsIds(oSlideCopy.cSld.spTree, oIdMap);
				deleted.push(oSlideCopy);
			}
		}
		deleted.reverse();
		for (i = 0; i < deleted.length; ++i) {
			presentation.insertSlideObjectToPos(pos + i, deleted[i]);
			aNewSelected.push(pos + i);
		}
		return aNewSelected;
	};
	SlideModeManager.prototype.applyHFProps = function (props, defaultLang) {
		const presentation = this.getPresentation();
		const selectedSlides = presentation.GetSelectedSlides();
		for (let i = 0; i < selectedSlides.length; i += 1) {
			const slide = this.getSlide(selectedSlides[i]);
			presentation.applyPropsToSlide(slide, slide.Layout, props);
		}
		return true;
	};
	SlideModeManager.prototype.isSlideNoteShape = function(shape) {
		return shape.parent && shape.parent.isNote() && shape.getPlaceholderType() === AscFormat.phType_body;
	};
	SlideModeManager.prototype.isDrawingSlide = function (slideObject) {
		return slideObject.isSlide();
	};
	SlideModeManager.prototype.recalculateNotesShape = function (slide) {
		slide.recalculateNotesShape();
	};
	SlideModeManager.prototype.isLockBackgroundSlides = function () {
		const slides = [];

		const presentation = this.getPresentation();
		const masterMap = {};
		for (let i = 0; i < presentation.Slides.length; i += 1) {
			const slide = presentation.Slides[i];
			slides.push(presentation.Slides[i]);
			const master = slide.Layout && slide.Layout.Master;
			if (master) {
				masterMap[master.GetId()] = master;
			}
		}
		for (let masterId in masterMap) {
			const master = masterMap[masterId];
			slides.push(master);
			for (let i = 0; i < master.sldLayoutLst.length; i += 1) {
				slides.push(master.sldLayoutLst[i]);
			}
		}

		return presentation.Document_Is_SelectionLocked(AscCommon.changestype_SlideBg, slides);
	};
	SlideModeManager.prototype.applySlideBackgroundToAll = function () {
		const presentation = this.getPresentation();
		const currentSlide = this.getCurrentSlide();
		if (!currentSlide) {
			return;
		}
		const bg = currentSlide.cSld.Bg;
		const isShowMasterShapes = currentSlide.showMasterSp;
		const mastersMap = {};
		for (let i = 0; i < presentation.Slides.length; i += 1) {
			const slide = presentation.Slides[i];
			slide.changeBackground(null);
			slide.setShowMasterSp(isShowMasterShapes);
			const master = slide.Layout && slide.Layout.Master;
			if (master) {
				mastersMap[master.GetId()] = master;
			}
		}
		for (let masterId in mastersMap) {
			const master = mastersMap[masterId];
			for (let j = 0; j < master.sldLayoutLst.length; j += 1) {
				const layout = master.sldLayoutLst[j];
				layout.changeBackground(null);
				layout.setShowMasterSp(isShowMasterShapes);
			}
			if (bg) {
				master.changeBackground(bg.createFullCopy());
			} else {
				master.changeBackground(null);
			}
		}
	};
	SlideModeManager.prototype.isCanAddHyperlink = function () {
		return true;
	};

	function MasterSlideModeManager(api) {
		SlideModeManagerBase.call(this, api);
		this.type = Asc.c_oAscPresentationViewMode.masterSlide;
	}

	AscFormat.InitClassWithoutType(MasterSlideModeManager, SlideModeManagerBase);
	MasterSlideModeManager.prototype.updateViewMode = function() {
		const wordControl = this.getWordControl();
		const presentation = this.getPresentation();
		const api = this.getApi();
		let oSlide = presentation.GetCurrentSlide();
		let nIdx = 0;
		if (oSlide) {
			let nCurIdx = presentation.GetSlideIndex(oSlide.Layout);
			if (nCurIdx !== -1) {
				nIdx = nCurIdx;
			}
		}
		wordControl.GoToPage(nIdx);
		wordControl.setNotesEnable(false);
		wordControl.setAnimPaneEnable(false);
		wordControl.setThumbnailsEnable(true);
		wordControl.SetThumbnails(false);
		api.hideMediaControl();
		api.asc_hideComments();
		presentation.Recalculate({Drawings: {All: true, Map: {}}});
		presentation.Document_UpdateInterfaceState();
	};
	MasterSlideModeManager.prototype.isMasterSlideMode = function() {
		return true;
	};
	MasterSlideModeManager.prototype.isMasterPlaceholderShape = function(shape) {
		return shape.isPlaceholder && shape.isPlaceholder();
	};
	MasterSlideModeManager.prototype.addMasterSlide = function() {
		const presentation = this.getPresentation();
		presentation.AddNewMasterSlide();
	};
	MasterSlideModeManager.prototype.addLayoutSlide = function() {
		const presentation = this.getPresentation();
		presentation.AddNewLayout();
	};
	MasterSlideModeManager.prototype.startAddPlaceholder = function(nType, bVertical, bStart) {
		const presentation = this.getPresentation();
		const oCurSlide = presentation.GetCurrentSlide();
		if (oCurSlide.getObjectType() !== AscDFH.historyitem_type_SlideLayout) return;
		presentation.StartAddShape("textRect", bStart, nType, bVertical);
	};
	MasterSlideModeManager.prototype.setTitle = function(bVal) {
		const presentation = this.getPresentation();
		const oCurSlide = presentation.GetCurrentSlide();
		if (oCurSlide.getObjectType() !== AscDFH.historyitem_type_SlideLayout) return;
		presentation.SetLayoutTitle(bVal);
	};
	MasterSlideModeManager.prototype.setFooter = function(bVal) {
		const presentation = this.getPresentation();
		const oCurSlide = presentation.GetCurrentSlide();
		if (oCurSlide.getObjectType() !== AscDFH.historyitem_type_SlideLayout) return;
		presentation.SetLayoutFooter(bVal);
	};
	MasterSlideModeManager.prototype.getMasters = function() {
		const oPresentation = this.getPresentation();
		return oPresentation.slideMasters;
	};
	MasterSlideModeManager.prototype.getCumulativeThumbnailsLength = function(isHorizontalOrientation, thumbnailWidth, thumbnailHeight) {
		let cumulativeThumbnailLength = 0;
		const masters = this.getMasters();
		for (let nIdx = 0; nIdx < masters.length; ++nIdx) {
			const master = masters[nIdx];
			const additionalLength = isHorizontalOrientation
				? thumbnailWidth + master.sldLayoutLst.length * thumbnailWidth * AscCommonSlide.SlideLayout.LAYOUT_THUMBNAIL_SCALE
				: thumbnailHeight + master.sldLayoutLst.length * thumbnailHeight * AscCommonSlide.SlideLayout.LAYOUT_THUMBNAIL_SCALE;
			cumulativeThumbnailLength += additionalLength;
		}
		return Math.round(cumulativeThumbnailLength);
	};
	MasterSlideModeManager.prototype.getSavedAnimationStartObject = function() {
		return this.getCurrentSlide();
	};
	MasterSlideModeManager.prototype.goToSavedAnimationStartObject = function(object) {
		const wordControl = this.getWordControl();
		if (!object) {
			return;
		}
		const idx = this.getSlideIndex(object);
		if (idx > -1) {
			wordControl.GoToPage(idx);
		} else {
			wordControl.GoToPage(0);
		}
	};
	MasterSlideModeManager.prototype.getAllSlides = function() {
		const presentation = this.getPresentation();
		const slides = [];
		for (let idx = 0; idx < presentation.slideMasters.length; ++idx) {
			let master = presentation.slideMasters[idx];
			slides.push(master);
			let aLayouts = master.sldLayoutLst;
			for (let nLt = 0; nLt < aLayouts.length; ++nLt) {
				slides.push(aLayouts[nLt]);
			}
		}
		return slides;
	};
	MasterSlideModeManager.prototype.getSlideNumber = function(idx) {
		const slideMasters = this.getMasters();
		const slide = this.getSlide(idx);
		if (slide.getObjectType() === AscDFH.historyitem_type_SlideMaster) {
			for (let nMaster = 0; nMaster < slideMasters.length; ++nMaster) {
				if (slideMasters[nMaster] === slide) {
					return nMaster + 1;
				}
			}
		}
		return null;
	};
	MasterSlideModeManager.prototype.insertSlideObjectToPos = function(pos, slide) {
		const presentation = this.getPresentation();
		const masters = this.getMasters();
		if (slide.isMaster()) {
			let curSlide = this.getSlide(pos - 1);
			let prevMaster = null;
			if (curSlide) {
				if (curSlide.isMaster()) {
					prevMaster = curSlide;
				} else {
					prevMaster = curSlide.Master;
				}
			}
			let masterPos = 0;
			if (prevMaster) {
				for (let idx = 0; idx < masters.length; ++idx) {
					if (masters[idx] === prevMaster) {
						masterPos = idx + 1;
						break;
					}
				}
			}
			presentation.addSlideMaster(masterPos, slide);
		} else {
			let curSlide = this.getSlide(pos - 1);
			let prevMaster = null;
			let prevLayout = null;
			if (curSlide) {
				if (curSlide.isMaster()) {
					prevMaster = curSlide;
				} else {
					prevLayout = curSlide;
					prevMaster = curSlide.Master;
				}
			}
			if (!prevMaster) {
				prevMaster = masters[0];
			}
			if (prevMaster) {
				let layoutPos = 0;
				if (prevLayout) {
					for (let idx = 0; idx < prevMaster.sldLayoutLst.length; ++idx) {
						if (prevMaster.sldLayoutLst[idx] === prevLayout) {
							layoutPos = idx + 1;
							break;
						}
					}
				}
				prevMaster.addToSldLayoutLstToPos(layoutPos, slide);
			}
		}
	};
	MasterSlideModeManager.prototype.isThumbnailsSupported = function() {
		return true;
	};
	MasterSlideModeManager.prototype.getCurrentTheme = function() {
		const slide = this.getCurrentSlide();
		if (slide) {
			if (slide.getObjectType() === AscDFH.historyitem_type_SlideLayout) {
				return slide.Master.Theme;
			} else {
				return slide.Theme;
			}
		}
		return null;
	};
	MasterSlideModeManager.prototype.getSizesMM = SlideModeManager.prototype.getSizesMM;
	MasterSlideModeManager.prototype.setPageOrientation = SlideModeManager.prototype.setPageOrientation;
	MasterSlideModeManager.prototype.addNewMasterSlide = function() {
		const presentation = this.getPresentation();
		presentation.AddForceNewMasterSlide();
	};
	MasterSlideModeManager.prototype.addNewLayout = function() {
		const presentation = this.getPresentation();
		presentation.AddForceNewLayout();
	};
	MasterSlideModeManager.prototype.updateInterfaceState = function() {
		const curSlide = this.getCurrentSlide();
		this.updateTitleCheckboxState(curSlide);
		this.updateFooterCheckBoxState(curSlide);
		this.updateCanDeleteMasterState();
		this.updateCanDeleteLayoutState(curSlide);
	};
	MasterSlideModeManager.prototype.updateCanDeleteLayoutState = function (curSlide) {
		const api = this.getApi();
		let isCanDeleteLayout = false;
		if (curSlide && curSlide.isLayout()) {
			const presentation = this.getPresentation();
			isCanDeleteLayout = true;
			let selectedSlides = presentation.GetSelectedSlides();
			for (let nIdx = 0; nIdx < selectedSlides.length; ++nIdx) {
				const slide = this.getSlide(selectedSlides[nIdx]);
				if (slide.isLayout()) {
					if (!presentation.CanRemoveLayout(slide)) {
						isCanDeleteLayout = false;
						break;
					}
				} else {
					isCanDeleteLayout = false;
					break;
				}
			}
		}
		api.sendEvent("asc_onCanDeleteLayout", isCanDeleteLayout);
	};
	MasterSlideModeManager.prototype.updateCanDeleteMasterState = function () {
		const presentation = this.getPresentation();
		const api = this.getApi();
		api.sendEvent("asc_onCanDeleteMaster", presentation.CanRemoveMaster(presentation.lastMaster));
	};
	MasterSlideModeManager.prototype.updateFooterCheckBoxState = function (curSlide) {
		const api = this.getApi();
		const eventName = "asc_onFooter";
		if (curSlide && curSlide.isLayout()) {
			this.updatePlaceholderCheckboxState(eventName, curSlide, [AscFormat.phType_ftr, AscFormat.phType_dt, AscFormat.phType_sldNum]);
		} else {
			api.sendEvent(eventName, false);
		}
	};
	MasterSlideModeManager.prototype.updateTitleCheckboxState = function (curSlide) {
		const api = this.getApi();
		const eventName = "asc_onTitle";
		if (curSlide && curSlide.isLayout()) {
			this.updatePlaceholderCheckboxState(eventName, curSlide, [AscFormat.phType_title]);
		} else {
			api.sendEvent(eventName, false);
		}
	};
	MasterSlideModeManager.prototype.getSlidesForChangeColorScheme = function() {
		return this.getSelectedSlides();
	};
	MasterSlideModeManager.prototype.changeTheme = function(themeInfo, arrInd) {
		let oCurSlide = this.getCurrentSlide();
		let oCurMaster = oCurSlide.getMaster();
		if (oCurMaster === themeInfo.Master) {
			return;
		}
		const presentation = this.getPresentation();
		for (let i = 0; i < presentation.slideMasters.length; ++i) {
			if (presentation.slideMasters[i] === themeInfo.Master) {
				return;
			}
		}
		const bReplace = !(oCurMaster.Theme.name === "Blank" || oCurMaster.Theme.name === "Office Theme" || oCurMaster.isPreserve());
		const arr_ind = [];
		for (let nSlide = 0; nSlide < presentation.Slides.length; ++nSlide) {
			if (presentation.Slides[nSlide].getMaster() === oCurMaster) {
				arr_ind.push(nSlide);
			}
		}
		if (!bReplace) {
			presentation.pushSlideMaster(themeInfo.Master);
		} else {
			for (let nMaster = 0; nMaster < presentation.slideMasters.length; ++nMaster) {
				if (presentation.slideMasters[nMaster] === oCurMaster) {
					presentation.removeSlideMaster(nMaster, 1);
					presentation.addSlideMaster(nMaster, themeInfo.Master)
					break;
				}
			}
		}
		this.replaceTheme(themeInfo, arr_ind);
		const nIdx = this.getSlideIndex(themeInfo.Master);
		if (nIdx !== -1) {
			const wordControl = this.getWordControl();
			wordControl.GoToPage(nIdx);
		}
		presentation.Document_UpdateInterfaceState();
	};
	MasterSlideModeManager.prototype.replaceTheme = SlideModeManager.prototype.replaceTheme;
	MasterSlideModeManager.prototype.setPreserve = function(bPr) {
		const oThis = this;
		const presentation = this.getPresentation();
		const arrIndexes = this.getSelectedSlides();
		if (!arrIndexes.length) {
			return;
		}
		arrIndexes.sort(AscCommon.fSortAscending);
		for (let i = 0; i < arrIndexes.length; i += 1) {
			const nIdx = arrIndexes[i];
			const oSlideObject = this.getSlide(nIdx);
			if (!(oSlideObject instanceof AscCommonSlide.MasterSlide)) {
				return;
			}
		}
		const arrMasterIndexesForDelete = [];
		const arrMastersForDelete = [];
		let nSlideIndex = null;
		if (bPr) {
			nSlideIndex = arrIndexes[arrIndexes.length - 1];
		} else {
			for (let i = 0; i < arrIndexes.length; i++) {
				const nIdx = arrIndexes[i];
				const oSlideObject = this.getSlide(nIdx);
				if (oSlideObject.IsUseInSlides()) {
					nSlideIndex = arrIndexes[i];
				} else {
					arrMasterIndexesForDelete.push(nIdx);
					arrMastersForDelete.push(oSlideObject);
				}
			}
		}

		function callback(bDelete) {
			if (bDelete && presentation.Document_Is_SelectionLocked(AscCommon.changestype_RemoveSlide, arrMastersForDelete)) {
				return;
			}

			presentation.StartAction(AscDFH.historydescription_Presentation_SetPreserveSlideMaster);
			for (let i = 0; i < arrIndexes.length; i++) {
				const nIdx = arrIndexes[i];
				const oSlideObject = oThis.getSlide(nIdx);
				oSlideObject.setPreserve(bPr);
			}

			if (bDelete) {
				for (let i = arrMasterIndexesForDelete.length - 1; i >= 0; i -= 1) {
					presentation.removeSlide(arrMasterIndexesForDelete[i]);
				}

				presentation.DrawingDocument.UpdateThumbnailsAttack();
				if (nSlideIndex === null) {
					nSlideIndex = Math.max(arrMasterIndexesForDelete[0] - 1, 0);
				}
				presentation.DrawingDocument.m_oWordControl.GoToPage(nSlideIndex, undefined, undefined, true);
			}

			presentation.Document_UpdateUndoRedoState();
			presentation.FinalizeAction();
		}

		if (arrMastersForDelete.length) {
			presentation.Api.sendEvent("asc_onRemoveUnpreserveMasters", callback);
		} else {
			callback();
		}
	};
	MasterSlideModeManager.prototype.deleteMaster = function() {
		const presentation = this.getPresentation();
		let oMaster = presentation.GetCurrentMaster();
		if (!oMaster) return;
		let nMasterIdx = -1;
		for (let nMaster = 0; nMaster < presentation.slideMasters.length; ++nMaster) {
			if (presentation.slideMasters[nMaster] === oMaster) {
				nMasterIdx = nMaster;
				break;
			}
		}
		let nIdx = presentation.GetSlideIndex(oMaster);
		presentation.StartAction(AscDFH.historydescription_Presentation_DeleteSlides, nIdx);
		presentation.removeSlide(nIdx);
		let nNewIdx = presentation.GetSlideIndex(oMaster);
		if (nNewIdx === -1) {
			let nCurIdx = 0;
			if (nMasterIdx > 0) {
				nCurIdx = nMasterIdx - 1;
			}
			let nPageIdx = presentation.GetSlideIndex(presentation.slideMasters[nCurIdx]);
			presentation.DrawingDocument.m_oWordControl.GoToPage(nPageIdx, undefined, undefined, true);
			presentation.Api.sync_HideComment();
			presentation.Document_UpdateUndoRedoState();
			presentation.Recalculate();
		}
		presentation.FinalizeAction();
	};
	MasterSlideModeManager.prototype.insertContent2 = function(oContent, aContents, kw, kh, bEndFormatting) {
		if (!(oContent.Layouts.length > 0 || oContent.Masters.length > 0)) {
			return false;
		}
		const presentation = this.getPresentation();
		const bChangeSize = kw !== 1 || kh !== 1;
		for (let i = 0; i < oContent.Layouts.length; ++i) {
			const oLayout = oContent.Layouts[i];
			if (bChangeSize) {
				oLayout.scale(kw, kh);
			}
			const sizes = this.getSizesMM();
			oLayout.setSlideSize(sizes.width, sizes.height);
			const nMasterIndex = oContent.MastersIndexes[i];
			const oMaster = oContent.Masters[nMasterIndex];
			if (oMaster) {
				oMaster.addLayout(oLayout);
			}
		}

		const oCurSlide = presentation.GetCurrentSlide();
		const oCurMaster = presentation.GetCurrentMaster();
		if (oContent.Masters.length > 0) {
			let nPos = presentation.slideMasters.length;
			for (let nMaster = 0; nMaster < presentation.slideMasters.length; ++nMaster) {
				if (presentation.slideMasters[nMaster] === oCurMaster) {
					nPos = nMaster + 1;
					break;
				}
			}
			for (let i = 0; i < oContent.Masters.length; ++i) {
				if (bChangeSize) {
					oContent.Masters[i].scale(kw, kh);
				}
				oContent.Masters[i].setSlideSize(presentation.GetWidthMM(), presentation.GetHeightMM());
				const oTheme = oContent.Themes[i];
				oContent.Masters[i].setTheme(oTheme);
				presentation.addSlideMaster(nPos + i, oContent.Masters[i]);
			}

			let nPage = presentation.GetSlideIndex(oContent.Masters[0]);
			if (nPage > -1) {
				presentation.CurPage = presentation.GetSlideIndex(oContent.Masters[0]);
				presentation.bGoToPage = true;
			}
		} else {
			if (oCurMaster) {
				let nPos = oCurMaster.sldLayoutLst.length;
				if (oCurSlide.isMaster()) {
					nPos = 0;
				} else {
					for (let nLt = 0; nLt < oCurMaster.sldLayoutLst.length; ++nLt) {
						if (oCurMaster.sldLayoutLst[nLt] === oCurSlide) {
							nPos = nLt + 1;
						}
					}
				}
				for (let i = 0; i < oContent.Layouts.length; ++i) {
					const oLayout = oContent.Layouts[i];
					oCurMaster.addToSldLayoutLstToPos(nPos + i, oLayout);
				}
				let nPage = presentation.GetSlideIndex(oContent.Layouts[0]);
				if (nPage > -1) {
					presentation.CurPage = presentation.GetSlideIndex(oContent.Layouts[0]);
					presentation.bGoToPage = true;
				}
			}
		}
		return false;
	};
	MasterSlideModeManager.prototype.recalculateThemeObjects = function(oThemeObjects) {
		const presentation = this.getPresentation();
		for (let nIdx = 0; nIdx < oThemeObjects.slides.length; ++nIdx) {
			oThemeObjects.slides[nIdx].recalculate();
		}
		let aIdx = [];
		let nStartIdx = 0;
		for (let nIdx = 0; nIdx < oThemeObjects.masters.length; ++nIdx) {
			let nMasterIdx = this.getSlideIndex(oThemeObjects.masters[nIdx]);
			if (nMasterIdx === presentation.CurPage) {
				nStartIdx = aIdx.length;
			}
			aIdx.push(nMasterIdx);
		}
		for (let nIdx = 0; nIdx < oThemeObjects.layouts.length; ++nIdx) {
			let nLayoutIdx = this.getSlideIndex(oThemeObjects.layouts[nIdx]);
			if (nLayoutIdx === presentation.CurPage) {
				nStartIdx = aIdx.length;
			}
			aIdx.push(nLayoutIdx);
		}
		AscFormat.redrawSlide(this.getSlide(aIdx[nStartIdx]), presentation, aIdx, nStartIdx, 0, this.getAllSlides());
	};
	MasterSlideModeManager.prototype.shiftSlides = function(pos, array, bCopyOnMove) {
		const presentation = this.getPresentation();
		const aNewSelected = [];
		let aToInsert = [];
		let oSlideLikeObject;
		if (bCopyOnMove) {
			for (let nIdx = 0; nIdx < array.length; ++nIdx) {
				let nIndexInSlides = array[nIdx];
				oSlideLikeObject = this.getSlide(nIndexInSlides);
				aToInsert.push(oSlideLikeObject.createDuplicate({}, false));
			}
		} else {
			for (let nIdx = array.length - 1; nIdx > -1; --nIdx) {
				let nIndexInSlides = array[nIdx];
				oSlideLikeObject = this.getSlide(nIndexInSlides);
				aToInsert.splice(0, 0, oSlideLikeObject);
				if (oSlideLikeObject.isMaster()) {
					presentation.removeSlideMasterObject(oSlideLikeObject);
				} else {
					oSlideLikeObject.Master.removeLayout(oSlideLikeObject);
				}
				if (nIndexInSlides < pos) {
					--pos;
				}
			}
		}
		let oPrevSlideLikeObj = this.getSlide(pos - 1);
		let oPrevMaster;
		if (oPrevSlideLikeObj) {
			if (oPrevSlideLikeObj.isMaster()) {
				oPrevMaster = oPrevSlideLikeObj;
			} else {
				oPrevMaster = oPrevSlideLikeObj.Master;
			}
		}
		if (oSlideLikeObject.isMaster()) {
			let nInsertPos = null;
			if (!oPrevSlideLikeObj) {
				nInsertPos = 0;
			} else {
				if (oPrevMaster) {
					for (let nIdx = 0; nIdx < presentation.slideMasters.length; ++nIdx) {
						if (presentation.slideMasters[nIdx] === oPrevMaster) {
							nInsertPos = nIdx + 1;
							break;
						}
					}
				}
			}
			if (nInsertPos !== null) {
				for (let nIdx = 0; nIdx < aToInsert.length; ++nIdx) {
					presentation.addSlideMaster(nInsertPos + nIdx, aToInsert[nIdx]);
					aNewSelected.push(this.getSlideIndex(aToInsert[nIdx]));
				}
			}
		} else {
			if (oPrevMaster) {
				let nInsertPos = null;
				let oMaster = null;
				if (oPrevSlideLikeObj.isMaster()) {
					nInsertPos = 0;
					oMaster = oPrevSlideLikeObj;
				} else {
					oMaster = oPrevSlideLikeObj.Master;
					for (let nIdx = 0; nIdx < oMaster.sldLayoutLst.length; ++nIdx) {
						if (oMaster.sldLayoutLst[nIdx] === oPrevSlideLikeObj) {
							nInsertPos = nIdx + 1;
							break;
						}
					}
				}
				if (oMaster !== null && nInsertPos !== null) {
					for (let nIdx = 0; nIdx < aToInsert.length; ++nIdx) {
						oMaster.addToSldLayoutLstToPos(nInsertPos + nIdx, aToInsert[nIdx]);
						aNewSelected.push(this.getSlideIndex(aToInsert[nIdx]));
					}
				}
			}
		}
		return aNewSelected;
	};
	MasterSlideModeManager.prototype.applyHFProps = function (props, defaultLang) {
		const depMap = {};
		const presentation = this.getPresentation();
		const selectedSlides = presentation.GetSelectedSlides();
		const slides = [];
		for (let i = 0; i < selectedSlides.length; i += 1) {
			const slide = this.getSlide(selectedSlides[i]);
			if (!slide.hf) {
				slide.setHF(new AscFormat.HF());
			}
			slide.hf.applySettings(props);
			slides.push(slide);
			depMap[slide.GetId()] = true;
		}
		presentation.applyHierarchyHFProperties(props, slides, defaultLang);
		for (let i = 0; i < presentation.Slides.length; i += 1) {
			const slide = presentation.Slides[i];
			if (slide.Layout && (depMap[slide.Layout.GetId()] || slide.Layout.Master && depMap[slide.Layout.Master.GetId()])) {
				presentation.applyPropsToSlide(slide, slide.Layout, props);
			}
		}
		return true;
	};
	MasterSlideModeManager.prototype.isDrawingSlide = function (slideObject) {
		return slideObject.isMaster() || slideObject.isLayout();
	};
	MasterSlideModeManager.prototype.isLockBackgroundSlides = function () {
		const slides = [];

		const presentation = this.getPresentation();
		for (let i = 0; i < presentation.Slides.length; i += 1) {
			const slide = presentation.Slides[i];
			slides.push(slide);
		}
		for (let i = 0; i < presentation.slideMasters.length; i += 1) {
			const master = presentation.slideMasters[i];
			slides.push(master);
			for (let j = 0; j < master.sldLayoutLst.length; j += 1) {
				slides.push(master.sldLayoutLst[j]);
			}
		}
		return presentation.Document_Is_SelectionLocked(AscCommon.changestype_SlideBg, slides);
	};
	MasterSlideModeManager.prototype.applySlideBackgroundToAll = function () {
		const presentation = this.getPresentation();
		const currentSlide = this.getCurrentSlide();
		if (!currentSlide) {
			return;
		}
		const bg = currentSlide.cSld.Bg;
		const isShowMasterShapes = currentSlide.isLayout() ? currentSlide.showMasterSp : true;
		for (let i = 0; i < presentation.slideMasters.length; i += 1) {
			const master = presentation.slideMasters[i];
			if (bg) {
				master.changeBackground(bg.createFullCopy());
			} else {
				master.changeBackground(null);
			}
			for (let j = 0; j < master.sldLayoutLst.length; j += 1) {
				const layout = master.sldLayoutLst[j];
				layout.changeBackground(null);
				layout.setShowMasterSp(isShowMasterShapes);
			}
		}
		for (let i = 0; i < presentation.Slides.length; i += 1) {
			const slide = presentation.Slides[i];
			slide.changeBackground(null);
			slide.setShowMasterSp(isShowMasterShapes);
		}
	};


	function NoteModeManager(api) {
		SlideModeManagerBase.call(this, api);
		this.type = Asc.c_oAscPresentationViewMode.notesPage;
	}

	AscFormat.InitClassWithoutType(NoteModeManager, SlideModeManagerBase);
	NoteModeManager.prototype.isNoteMode = function() {
		return true;
	};
	NoteModeManager.prototype.getAllSlides = function() {
		const presentation = this.getPresentation();
		return presentation.notes;
	};
	NoteModeManager.prototype.getSizesMM = function() {
		const presentation = this.getPresentation();
		return presentation.GetNotesSizesMM();
	};
	NoteModeManager.prototype.setPageOrientation = function(val) {
		const presentation = this.getPresentation();
		presentation.SetNotePageOrientation(val);
	};
	NoteModeManager.prototype.changeTheme = function(themeInfo, arrInd) {
		const presentation = this.getPresentation();
		if (!Array.isArray(arrInd)) {
			arrInd = [presentation.CurPage];
		}
		const masterMap = {};
		for (let i = 0; i < arrInd.length; i += 1) {
			const slide = this.getSlide(arrInd[i]);
			if (slide.Master && !masterMap[slide.Master.Get_Id()]) {
				masterMap[slide.Master.Get_Id()] = true;
				slide.Master.setTheme(themeInfo.Theme);
			}
		}
		presentation.Document_UpdateInterfaceState();
		presentation.Recalculate();
	};
	NoteModeManager.prototype.isDrawingSlide = function (slideObject) {
		return slideObject.isNote();
	};
	NoteModeManager.prototype.recalculateThemeObjects = function(oThemeObjects) {
		const presentation = this.getPresentation();
		for (let nIdx = 0; nIdx < oThemeObjects.notesMasters.length; ++nIdx) {
			oThemeObjects.notesMasters[nIdx].recalculate();
		}
		let aIdx = [];
		let nStartIdx = 0;
		for (let nIdx = 0; nIdx < oThemeObjects.notes.length; ++nIdx) {
			const note = oThemeObjects.notes[nIdx];
			const slideIndex = presentation.GetSlideIndex(note);
			if (slideIndex === presentation.CurPage) {
				nStartIdx = aIdx.length;
			}
			aIdx.push(slideIndex);
		}
		AscFormat.redrawSlide(presentation.notes[aIdx[nStartIdx]], presentation, aIdx, nStartIdx, 0, this.getAllSlides());
	};
	NoteModeManager.prototype.updateViewMode = function (oldViewManager) {
		const wordControl = this.getWordControl();
		const presentation = this.getPresentation();
		const api = this.getApi();
		let oSlide = oldViewManager.getCurrentSlide();
		let nIdx = 0;
		if (oSlide) {
			let nCurIdx = this.getSlideIndex(oSlide.notes);
			if (nCurIdx !== -1) {
				nIdx = nCurIdx;
			}
		}
		for (let i = 0; i < presentation.notes.length; i += 1) {
			const note = presentation.notes[i];
			note.recalcBody();
			note.graphicObjects.resetSelection();
		}
		wordControl.setNotesEnable(false);
		wordControl.setAnimPaneEnable(false);
		wordControl.setThumbnailsEnable(false);
		wordControl.GoToPage(nIdx);
		api.hideMediaControl();
		api.asc_hideComments();
		presentation.Recalculate({Drawings: {All: true, Map: {}}});
		presentation.Document_UpdateInterfaceState();
	};

	NoteModeManager.prototype.addNextSlide = function (layoutIndex) {
		const presentation = this.getPresentation();
		presentation.addForceNextSlide(layoutIndex);
	};
	NoteModeManager.prototype.isLockBackgroundSlides = function () {
		const notes = [];

		const presentation = this.getPresentation();
		const masterMap = {};
		for (let i = 0; i < presentation.notes.length; i += 1) {
			const note = presentation.notes[i];
			notes.push(note);
			if (note.Master) {
				masterMap[note.Master.GetId()] = note.Master;
			}
		}
		for (let masterId in masterMap) {
			const master = masterMap[masterId];
			notes.push(master);
		}
		return presentation.Document_Is_SelectionLocked(AscCommon.changestype_SlideBg, notes);
	};
	NoteModeManager.prototype.applySlideBackgroundToAll = function () {
		const presentation = this.getPresentation();
		const currentSlide = this.getCurrentSlide();
		if (!currentSlide) {
			return;
		}
		const bg = currentSlide.cSld.Bg;
		const isShowMasterShapes = currentSlide.showMasterSp;
		const mastersMap = {};
		for (let i = 0; i < presentation.notes.length; i += 1) {
			const note = presentation.notes[i];
			note.changeBackground(null);
			note.setShowMasterSp(isShowMasterShapes);
			const master = note.Master;
			if (master) {
				mastersMap[master.GetId()] = master;
			}
		}
		for (let masterId in mastersMap) {
			const master = mastersMap[masterId];
			if (bg) {
				master.changeBackground(bg.createFullCopy());
			} else {
				master.changeBackground(null);
			}
		}
	};


	function MasterNoteModeManager(api) {
		SlideModeManagerBase.call(this, api);
		this.type = Asc.c_oAscPresentationViewMode.notesMaster;
	}

	AscFormat.InitClassWithoutType(MasterNoteModeManager, SlideModeManagerBase);
	MasterNoteModeManager.prototype.isMasterNoteMode = function() {
		return true;
	};
	MasterNoteModeManager.prototype.getAllSlides = function() {
		const presentation = this.getPresentation();
		return presentation.notesMasters;
	};
	MasterNoteModeManager.prototype.getSizesMM = NoteModeManager.prototype.getSizesMM;
	MasterNoteModeManager.prototype.setPageOrientation = NoteModeManager.prototype.setPageOrientation;
	MasterNoteModeManager.prototype.isMasterPlaceholderShape = function(shape) {
		return shape.isPlaceholder && shape.isPlaceholder();
	};
	MasterNoteModeManager.prototype.setFooter = function(val) {
		this.setPlaceholder(val, AscFormat.phType_ftr, AscCommonSlide.addFooterShape);
	};
	MasterNoteModeManager.prototype.setHeader = function(val) {
		this.setPlaceholder(val, AscFormat.phType_hdr, AscCommonSlide.addHeaderShape);
	};
	MasterNoteModeManager.prototype.setDate = function(val) {
		this.setPlaceholder(val, AscFormat.phType_dt, AscCommonSlide.addDateShape);
	};
	MasterNoteModeManager.prototype.setPageNumber = function(val) {
		this.setPlaceholder(val, AscFormat.phType_sldNum, AscCommonSlide.addNumberShape);
	};

	MasterNoteModeManager.prototype.setBody = function(val) {
		const slide = this.getCurrentSlide();
		if (!slide) {
			return;
		}
		this.startAction(0);
		const shape = slide.getBodyShape();
		if (!val) {
			if (shape) {
				slide.graphicObjects.deselectObject(shape);
				slide.removeFromSpTreeById(shape.Get_Id());
			}
		} else if (!shape) {
			AscCommonSlide.addBodyShape(slide);
		}
		this.finalizeAction(true);
	};
	MasterNoteModeManager.prototype.setSlideImage = function(val) {
		this.setPlaceholder(val, AscFormat.phType_sldImg, AscCommonSlide.addSlideImage);
	};
	MasterNoteModeManager.prototype.setPlaceholder = function(val, type, addCallback) {
		const slide = this.getCurrentSlide();
		if (!slide) {
			return;
		}
		this.startAction(0);
		const shape = slide.getMatchingShape(type, null, false, {});
		if (!val) {
			if (shape) {
				slide.graphicObjects.deselectObject(shape);
				slide.removeFromSpTreeById(shape.Get_Id());
			}
		} else if (!shape) {
			addCallback(slide);
		}
		this.finalizeAction(true);
	};
	MasterNoteModeManager.prototype.changeTheme = function(themeInfo, arrInd) {
		const presentation = this.getPresentation();
		if (!Array.isArray(arrInd)) {
			arrInd = [presentation.CurPage];
		}
		for (let i = 0; i < arrInd.length; i += 1) {
			const slide = this.getSlide(arrInd[i]);
			slide.setTheme(themeInfo.Theme);
			slide.checkSlideTheme();
		}
		presentation.Document_UpdateInterfaceState();
		presentation.Recalculate();
	};
	MasterNoteModeManager.prototype.recalculateThemeObjects = function(oThemeObjects) {
		const presentation = this.getPresentation();
		for (let nIdx = 0; nIdx < oThemeObjects.notes.length; ++nIdx) {
			oThemeObjects.notes[nIdx].recalculate();
		}
		let aIdx = [];
		let nStartIdx = 0;
		for (let nIdx = 0; nIdx < oThemeObjects.notesMasters.length; ++nIdx) {
			const masterIndex = presentation.GetSlideIndex(oThemeObjects.notesMasters[nIdx]);
			if (masterIndex === presentation.CurPage) {
				nStartIdx = aIdx.length;
			}
			aIdx.push(masterIndex);
		}
		AscFormat.redrawSlide(presentation.notesMasters[aIdx[nStartIdx]], presentation, aIdx, nStartIdx, 0, this.getAllSlides());
	};
	MasterNoteModeManager.prototype.isDrawingSlide = function (slideObject) {
		return slideObject.isNoteMaster();
	};
	MasterNoteModeManager.prototype.updateViewMode = function (oldViewManager) {
		const wordControl = this.getWordControl();
		const presentation = this.getPresentation();
		const api = this.getApi();
		wordControl.GoToPage(0);
		wordControl.setNotesEnable(false);
		wordControl.setAnimPaneEnable(false);
		wordControl.setThumbnailsEnable(false);
		api.hideMediaControl();
		api.asc_hideComments();
		presentation.Recalculate({Drawings: {All: true, Map: {}}});
		presentation.Document_UpdateInterfaceState();
	};
	MasterNoteModeManager.prototype.updateInterfaceState = function() {
		const currentSlide = this.getCurrentSlide();
		this.updateHeaderCheckboxState(currentSlide);
		this.updateSlideImageCheckboxState(currentSlide);
		this.updateFooterCheckboxState(currentSlide);
		this.updateDateCheckboxState(currentSlide);
		this.updateBodyCheckboxState(currentSlide);
		this.updatePageNumberCheckboxState(currentSlide);
		this.updatePageOrientationState();
	};
	MasterNoteModeManager.prototype.updateHeaderCheckboxState = function (currentSlide) {
		this.updatePlaceholderCheckboxState("asc_onHeader", currentSlide, [AscFormat.phType_hdr]);
	};
	MasterNoteModeManager.prototype.updateSlideImageCheckboxState = function (currentSlide) {
		this.updatePlaceholderCheckboxState("asc_onSlideImage", currentSlide, [AscFormat.phType_sldImg]);
	};
	MasterNoteModeManager.prototype.updateFooterCheckboxState = function (currentSlide) {
		this.updatePlaceholderCheckboxState("asc_onFooter", currentSlide, [AscFormat.phType_ftr]);
	};
	MasterNoteModeManager.prototype.updateDateCheckboxState = function (currentSlide) {
		this.updatePlaceholderCheckboxState("asc_onDate", currentSlide, [AscFormat.phType_dt]);
	};
	MasterNoteModeManager.prototype.updateBodyCheckboxState = function (currentSlide) {
		let isMatching = false;
		if (currentSlide) {
			const body = currentSlide.getBodyShape();
			if (body) {
				isMatching = true;
			}
		}
		const api = this.getApi();
		api.sendEvent("asc_onBody", isMatching);
	};
	MasterNoteModeManager.prototype.updatePageNumberCheckboxState = function (currentSlide) {
		this.updatePlaceholderCheckboxState("asc_onPageNumber", currentSlide, [AscFormat.phType_sldNum]);
	};
	MasterNoteModeManager.prototype.updatePageOrientationState = function () {
		const presentation = this.getPresentation();
		if (presentation) {
			presentation.UpdateNotePageOrientationState();
		}
	};
	MasterNoteModeManager.prototype.isLockBackgroundSlides = function () {
		const notes = [];

		const presentation = this.getPresentation();
		for (let i = 0; i < presentation.notes.length; i += 1) {
			const note = presentation.notes[i];
			notes.push(note);
		}
		for (let i = 0; i < presentation.notesMasters.length; i += 1) {
			const master = presentation.notesMasters[i];
			notes.push(master);
		}
		return presentation.Document_Is_SelectionLocked(AscCommon.changestype_SlideBg, notes);
	};
	MasterNoteModeManager.prototype.applySlideBackgroundToAll = function () {
		const presentation = this.getPresentation();
		const currentSlide = this.getCurrentSlide();
		if (!currentSlide) {
			return;
		}
		const bg = currentSlide.cSld.Bg;
		for (let i = 0; i < presentation.notesMasters.length; i += 1) {
			const master = presentation.notesMasters[i];
			if (bg) {
				master.changeBackground(bg.createFullCopy());
			} else {
				master.changeBackground(null);
			}
		}
		for (let i = 0; i < presentation.notes.length; i += 1) {
			const note = presentation.notes[i];
			note.changeBackground(null);
			note.setShowMasterSp(true);
		}
	};

	function MasterHandoutModeManager(api) {
		SlideModeManagerBase.call(this, api);
		this.type = Asc.c_oAscPresentationViewMode.handoutMaster;
	}

	AscFormat.InitClassWithoutType(MasterHandoutModeManager, SlideModeManagerBase);
	MasterHandoutModeManager.prototype.isMasterHandoutMode = function() {
		return true;
	};
	MasterHandoutModeManager.prototype.getAllSlides = function() {
		const presentation = this.getPresentation();
		return presentation.handoutMasters;
	};
	MasterHandoutModeManager.prototype.isMasterPlaceholderShape = function(shape) {
		return shape.isPlaceholder && shape.isPlaceholder();
	};

	MasterHandoutModeManager.prototype.setFooter = MasterNoteModeManager.prototype.setFooter;
	MasterHandoutModeManager.prototype.setHeader = MasterNoteModeManager.prototype.setHeader;
	MasterHandoutModeManager.prototype.setDate = MasterNoteModeManager.prototype.setDate;
	MasterHandoutModeManager.prototype.setPageNumber = MasterNoteModeManager.prototype.setPageNumber;
	MasterHandoutModeManager.prototype.setPlaceholder = MasterNoteModeManager.prototype.setPlaceholder;
	MasterHandoutModeManager.prototype.setHandoutPageCount = function(val) {
		const slide = this.getCurrentSlide();
		if (slide) {
			const oldCount = slide.getSlideCount();
			if (oldCount !== val) {
				const presentation = this.getPresentation();
				slide.setSlideCount(val);
				presentation.DrawingDocument.OnRecalculateSlide(presentation.CurPage);
			}
		}
	};
	MasterHandoutModeManager.prototype.getSizesMM = NoteModeManager.prototype.getSizesMM;
	MasterHandoutModeManager.prototype.setPageOrientation = NoteModeManager.prototype.setPageOrientation;
	MasterHandoutModeManager.prototype.isDrawingSlide = function (slideObject) {
		return slideObject.isHandoutMaster();
	};
	MasterHandoutModeManager.prototype.recalculateThemeObjects = function(oThemeObjects) {
		const presentation = this.getPresentation();
		let nStartIdx = 0;
		const aIdx = [];
		for (let nIdx = 0; nIdx < oThemeObjects.handoutMasters.length; ++nIdx) {
			const masterIndex = presentation.GetSlideIndex(oThemeObjects.handoutMasters[nIdx]);
			if (masterIndex === presentation.CurPage) {
				nStartIdx = aIdx.length;
			}
			aIdx.push(masterIndex);
		}
		AscFormat.redrawSlide(presentation.handoutMasters[aIdx[nStartIdx]], presentation, aIdx, nStartIdx, 0, this.getAllSlides());
	};
	MasterHandoutModeManager.prototype.updateViewMode = MasterNoteModeManager.prototype.updateViewMode;
	MasterHandoutModeManager.prototype.updateInterfaceState = function() {
		const currentSlide = this.getCurrentSlide();
		this.updateHeaderCheckboxState(currentSlide);
		this.updateFooterCheckboxState(currentSlide);
		this.updateDateCheckboxState(currentSlide);
		this.updatePageNumberCheckboxState(currentSlide);
		this.updateSlidesPerPageState(currentSlide);
		this.updatePageOrientationState();
	};
	MasterHandoutModeManager.prototype.updateHeaderCheckboxState = MasterNoteModeManager.prototype.updateHeaderCheckboxState;
	MasterHandoutModeManager.prototype.updateFooterCheckboxState = MasterNoteModeManager.prototype.updateFooterCheckboxState;
	MasterHandoutModeManager.prototype.updateDateCheckboxState = MasterNoteModeManager.prototype.updateDateCheckboxState;
	MasterHandoutModeManager.prototype.updatePageNumberCheckboxState = MasterNoteModeManager.prototype.updatePageNumberCheckboxState;
	MasterHandoutModeManager.prototype.updatePageOrientationState = MasterNoteModeManager.prototype.updatePageOrientationState;
	MasterHandoutModeManager.prototype.updateSlidesPerPageState = function (currentPage) {
		let slidesCount = 1;
		if (currentPage) {
			slidesCount = currentPage.getSlidesCount();
		}
		const api = this.getApi();
		api.sendEvent("asc_onSlidesPerPage", slidesCount);
	};
	MasterHandoutModeManager.prototype.isLockBackgroundSlides = function () {
		const presentation = this.getPresentation();
		return presentation.Document_Is_SelectionLocked(AscCommon.changestype_SlideBg, presentation.handoutMasters);
	};
	MasterHandoutModeManager.prototype.applySlideBackgroundToAll = function () {
		const presentation = this.getPresentation();
		const currentSlide = this.getCurrentSlide();
		if (!currentSlide) {
			return;
		}
		const bg = currentSlide.cSld.Bg;
		for (let i = 0; i < presentation.handoutMasters.length; i += 1) {
			const master = presentation.handoutMasters[i];
			if (bg) {
				master.changeBackground(bg.createFullCopy());
			} else {
				master.changeBackground(null);
			}
		}
	};

	function SorterModeManager(api) {
		SlideModeManagerBase.call(this, api);
		this.type = Asc.c_oAscPresentationViewMode.sorter;
	}

	AscFormat.InitClassWithoutType(SorterModeManager, SlideModeManagerBase);
	SorterModeManager.prototype.isSorterMode = function() {
		return true;
	};

	function OutlineModeManager(api) {
		SlideModeManager.call(this, api);
		this.type = Asc.c_oAscPresentationViewMode.outline;
		this.outlineView = new AscCommonSlide.OutlineView(api);
	}

	AscFormat.InitClassWithoutType(OutlineModeManager, SlideModeManager);
	OutlineModeManager.prototype.isOutlineMode = function() {
		return true;
	};
	OutlineModeManager.prototype.isSlideMode = function() {
		return true;
	};
	OutlineModeManager.prototype.isFocusOnOutline = function () {
		const api = this.getApi();
		return api.WordControl.Thumbnails.isOutline();
	};
	OutlineModeManager.prototype.updateViewMode = function (oldViewManager) {
		const wordControl = this.getWordControl();
		const presentation = this.getPresentation();
		const slideIdx = this.getChangeModePage(oldViewManager);
		const api = this.getApi();
		wordControl.setNotesEnable(true);
		wordControl.setAnimPaneEnable(true);
		wordControl.setThumbnailsEnable(true);
		wordControl.SetThumbnails(true);
		wordControl.GoToPage(slideIdx);
		api.hideMediaControl();
		api.asc_hideComments();
		presentation.Recalculate({Drawings: {All: true, Map: {}}});
		presentation.Document_UpdateInterfaceState();
	};

	function getViewModeByType(type, api) {
		switch (type) {
			case Asc.c_oAscPresentationViewMode.normal: {
				return new SlideModeManager(api);
			}
			case Asc.c_oAscPresentationViewMode.notesPage: {
				return new NoteModeManager(api);
			}
			case Asc.c_oAscPresentationViewMode.masterSlide: {
				return new MasterSlideModeManager(api);
			}
			case Asc.c_oAscPresentationViewMode.handoutMaster: {
				return new MasterHandoutModeManager(api);
			}
			case Asc.c_oAscPresentationViewMode.notesMaster: {
				return new MasterNoteModeManager(api);
			}
			case Asc.c_oAscPresentationViewMode.sorter: {
				return new SorterModeManager(api);
			}
			case Asc.c_oAscPresentationViewMode.outline: {
				return new OutlineModeManager(api);
			}
		}
	}

	window["AscCommon"] = window["AscCommon"] || {};
	window["AscCommon"].getViewModeByType = getViewModeByType;
})();
