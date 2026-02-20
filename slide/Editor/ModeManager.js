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

(function (undefined) {
	function SlideModeManagerBase(api) {
		this.api = api;
		this.type = null;
	}

	SlideModeManagerBase.prototype.getPresentation = function () {
		return this.api.WordControl.m_oLogicDocument;
	};
	SlideModeManagerBase.prototype.getWordControl = function () {
		return this.api.WordControl;
	};
	SlideModeManagerBase.prototype.getApi = function () {
		return this.api;
	};
	SlideModeManagerBase.prototype.getCurrentSlide = function () {
		const presentation = this.getPresentation();
		return this.getAllSlides()[presentation.CurPage];
	};
	SlideModeManagerBase.prototype.getSlideIndex = function (object) {
		const allSlides = this.getAllSlides();
		const presentation = this.getPresentation();
		for (let i = 0; i < allSlides.length; i += 1) {
			if (allSlides[i] === object) {
				return i;
			}
		}
		return -1;
	};
	SlideModeManagerBase.prototype.isMasterSlideMode = function () {
		return false;
	};
	SlideModeManagerBase.prototype.isSlideMode = function () {
		return false;
	};
	SlideModeManagerBase.prototype.isNoteMode = function () {
		return false;
	};
	SlideModeManagerBase.prototype.isMasterNoteMode = function () {
		return false;
	};
	SlideModeManagerBase.prototype.isMasterHandoutMode = function () {
		return false;
	};
	SlideModeManagerBase.prototype.isSorterMode = function () {
		return false;
	};
	SlideModeManagerBase.prototype.getSlidesCount = function () {
		return this.getAllSlides().length;
	};
	SlideModeManagerBase.prototype.getAllSlides = function () {
		return [];
	};
	SlideModeManagerBase.prototype.updateViewMode = function () {};
	SlideModeManagerBase.prototype.setSpecialPasteShowOptions = function (props) {};
	SlideModeManagerBase.prototype.isCanAddHyperlinkInContent = function () {
		return false;
	};
	SlideModeManagerBase.prototype.isMasterPlaceholderShape = function (shape) {
		return false;
	};
	SlideModeManagerBase.prototype.addMasterSlide = function () {};
	SlideModeManagerBase.prototype.addLayoutSlide = function () {};
	SlideModeManagerBase.prototype.startAddPlaceholder = function (nType, bVertical, bStart) {};
	SlideModeManagerBase.prototype.setLayoutTitle = function (bVal) {};
	SlideModeManagerBase.prototype.setLayoutFooter = function (bVal) {};
	SlideModeManagerBase.prototype.applySlideTransition = function (oTransition) {};
	SlideModeManagerBase.prototype.applySlideTransitionToAll = function () {};
	SlideModeManagerBase.prototype.getCumulativeThumbnailsLength = function (isHorizontalOrientation, thumbnailWidth, thumbnailHeight) {return 0;};
	SlideModeManagerBase.prototype.isNotesSupported = function () {return false;};
	SlideModeManagerBase.prototype.getAnimationStartSlideNum = function (startSlideNum) {return 0;};
	SlideModeManagerBase.prototype.getSavedAnimationStartObject = function () {return null;};
	SlideModeManagerBase.prototype.goToSavedAnimationStartObject = function (object) {};
	SlideModeManagerBase.prototype.getSlideNumber = function (object) {return null;};
	SlideModeManagerBase.prototype.insertSlideObjectToPos = function (pos, slide) {};


	function SlideModeManager(api) {
		SlideModeManagerBase.call(this, api);
		this.type = Asc.c_oAscPresentationViewMode.normal;
	}

	AscFormat.InitClassWithoutType(SlideModeManager, SlideModeManagerBase);
	SlideModeManager.prototype.updateViewMode = function () {
		const wordControl = this.getWordControl();
		const presentation = this.getPresentation();
		const api = this.getApi();
		wordControl.GoToPage(0);
		wordControl.setNotesEnable(true);
		wordControl.setAnimPaneEnable(true);
		api.hideMediaControl();
		api.asc_hideComments();
		presentation.Recalculate({Drawings: {All: true, Map: {}}});
		presentation.Document_UpdateInterfaceState();
	};
	SlideModeManager.prototype.isSlideMode = function () {
		return true;
	};
	SlideModeManager.prototype.setSpecialPasteShowOptions = function (props) {
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
		let bThumbnals = presentation.IsFocusOnThumbnails();
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
	SlideModeManager.prototype.isCanAddHyperlinkInContent = function () {
		return true;
	};
	SlideModeManager.prototype.applySlideTransition = function (oTransition) {
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
		}
	};
	SlideModeManager.prototype.applySlideTransitionToAll = function () {
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
	SlideModeManager.prototype.getCumulativeThumbnailsLength = function (isHorizontalOrientation, thumbnailWidth, thumbnailHeight) {
		const slidesCount = this.getSlidesCount();
		return isHorizontalOrientation ? thumbnailWidth * slidesCount : thumbnailHeight * slidesCount;
	};
	SlideModeManager.prototype.isNotesSupported = function () {
		return true;
	};
	SlideModeManager.prototype.getAnimationStartSlideNum = function (startSlideNum) {
		return -1 === startSlideNum ? 0 : startSlideNum;
	};
	SlideModeManager.prototype.getAllSlides = function () {
		const presentation = this.getPresentation();
		return presentation.Slides;
	};
	SlideModeManager.prototype.getSlideNumber = function (idx) {
		const presentation = this.getPresentation();
		return idx + presentation.getFirstSlideNumber();
	};
	SlideModeManager.prototype.insertSlideObjectToPos = function (pos, slide) {
		this.getPresentation().insertSlide(pos, slide);
	};


	function MasterSlideModeManager(api) {
		SlideModeManagerBase.call(this, api);
		this.type = Asc.c_oAscPresentationViewMode.masterSlide;
	}

	AscFormat.InitClassWithoutType(MasterSlideModeManager, SlideModeManagerBase);
	MasterSlideModeManager.prototype.updateViewMode = function () {
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
		presentation.Recalculate({Drawings: {All: true, Map: {}}});
		wordControl.setNotesEnable(false);
		wordControl.setAnimPaneEnable(false);
		api.hideMediaControl();
		api.asc_hideComments();
		presentation.Document_UpdateInterfaceState();
	};
	MasterSlideModeManager.prototype.isMasterSlideMode = function () {
		return true;
	};
	MasterSlideModeManager.prototype.isMasterPlaceholderShape = function (shape) {
		return shape.isPlaceholder && shape.isPlaceholder();
	};
	MasterSlideModeManager.prototype.addMasterSlide = function () {
		const presentation = this.getPresentation();
		presentation.AddNewMasterSlide();
	};
	MasterSlideModeManager.prototype.addLayoutSlide = function () {
		const presentation = this.getPresentation();
		presentation.AddNewLayout();
	};
	MasterSlideModeManager.prototype.startAddPlaceholder = function (nType, bVertical, bStart) {
		const presentation = this.getPresentation();
		const oCurSlide = presentation.GetCurrentSlide();
		if (oCurSlide.getObjectType() !== AscDFH.historyitem_type_SlideLayout) return;
		presentation.StartAddShape("textRect", bStart, nType, bVertical);
	};
	MasterSlideModeManager.prototype.setLayoutTitle = function (bVal) {
		const presentation = this.getPresentation();
		const oCurSlide = presentation.GetCurrentSlide();
		if (oCurSlide.getObjectType() !== AscDFH.historyitem_type_SlideLayout) return;
		presentation.SetLayoutTitle(bVal);
	};
	MasterSlideModeManager.prototype.setLayoutFooter = function (bVal) {
		const presentation = this.getPresentation();
		const oCurSlide = presentation.GetCurrentSlide();
		if (oCurSlide.getObjectType() !== AscDFH.historyitem_type_SlideLayout) return;
		presentation.SetLayoutFooter(bVal);
	};
	MasterSlideModeManager.prototype.getMasters = function () {
		const oPresentation = this.getPresentation();
		return oPresentation.slideMasters;
	};
	MasterSlideModeManager.prototype.getCumulativeThumbnailsLength = function (isHorizontalOrientation, thumbnailWidth, thumbnailHeight) {
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
	MasterSlideModeManager.prototype.getSavedAnimationStartObject = function () {
		return this.getCurrentSlide();
	};
	MasterSlideModeManager.prototype.goToSavedAnimationStartObject = function (object) {
		const wordControl = this.getWordControl;
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
	MasterSlideModeManager.prototype.getAllSlides = function () {
		const presentation = this.getPresentation();
		const slides = [];
		for(let idx = 0; idx < presentation.slideMasters.length; ++idx) {
			let master = presentation.slideMasters[idx];
			slides.push(master);
			let aLayouts = master.sldLayoutLst;
			for(let nLt = 0; nLt < aLayouts.length; ++nLt) {
				slides.push(aLayouts[nLt]);
			}
		}
		return slides;
	};
	MasterSlideModeManager.prototype.getSlideNumber = function (idx) {
		const slideMasters = this.getMasters();
		const slide = this.getSlide(idx);
		if(slide.getObjectType() === AscDFH.historyitem_type_SlideMaster) {
			for(let nMaster = 0; nMaster < slideMasters.length; ++nMaster) {
				if(slideMasters[nMaster] === slide) {
					return nMaster + 1;
				}
			}
		}
	};
	MasterSlideModeManager.prototype.insertSlideObjectToPos = function (pos, slide) {
		const presentation = this.getPresentation();
		const masters = this.getMasters();
		if(slide.isMaster()) {
			let curSlide = this.getSlide(pos - 1);
			let prevMaster = null;
			if(curSlide) {
				if(curSlide.isMaster()) {
					prevMaster = curSlide;
				}
				else {
					prevMaster = curSlide.Master;
				}
			}
			let masterPos = 0;
			if(prevMaster) {
				for(let idx = 0; idx < masters.length; ++idx) {
					if(masters[idx] === prevMaster) {
						masterPos = idx + 1;
						break;
					}
				}
			}
			presentation.addSlideMaster(masterPos, slide);
		}
		else {
			let curSlide = this.getSlide(pos - 1);
			let prevMaster = null;
			let prevLayout = null;
			if(curSlide) {
				if(curSlide.isMaster()) {
					prevMaster = curSlide;
				}
				else {
					prevLayout = curSlide;
					prevMaster = curSlide.Master;
				}
			}
			if(!prevMaster) {
				prevMaster = masters[0];
			}
			if(prevMaster) {
				let layoutPos = 0;
				if(prevLayout) {
					for(let idx = 0; idx < prevMaster.sldLayoutLst.length; ++idx) {
						if(prevMaster.sldLayoutLst[idx] === prevLayout) {
							layoutPos = idx + 1;
							break;
						}
					}
				}
				prevMaster.addToSldLayoutLstToPos(layoutPos, slide);
			}
		}
	};


	function NoteModeManager(api) {
		SlideModeManagerBase.call(this, api);
		this.type = Asc.c_oAscPresentationViewMode.notesPage;
	}

	AscFormat.InitClassWithoutType(NoteModeManager, SlideModeManagerBase);
	NoteModeManager.prototype.isNoteMode = function () {
		return true;
	};
	NoteModeManager.prototype.getAllSlides = function () {
		const presentation = this.getPresentation();
		return presentation.notes;
	};

	function MasterNoteModeManager(api) {
		SlideModeManagerBase.call(this, api);
		this.type = Asc.c_oAscPresentationViewMode.notesMaster;
	}

	AscFormat.InitClassWithoutType(MasterNoteModeManager, SlideModeManagerBase);
	MasterNoteModeManager.prototype.isMasterNoteMode = function () {
		return true;
	};
	MasterNoteModeManager.prototype.getAllSlides = function () {
		const presentation = this.getPresentation();
		return presentation.notesMasters;
	};

	function MasterHandoutModeManager(api) {
		SlideModeManagerBase.call(this, api);
		this.type = Asc.c_oAscPresentationViewMode.handoutMaster;
	}

	AscFormat.InitClassWithoutType(MasterHandoutModeManager, SlideModeManagerBase);
	MasterHandoutModeManager.prototype.isMasterHandoutMode = function () {
		return true;
	};
	MasterHandoutModeManager.prototype.getAllSlides = function () {
		const presentation = this.getPresentation();
		return presentation.handoutMasters;
	};

	function SorterModeManager(api) {
		SlideModeManagerBase.call(this, api);
		this.type = Asc.c_oAscPresentationViewMode.sorter;
	}

	AscFormat.InitClassWithoutType(SorterModeManager, SlideModeManagerBase);
	SorterModeManager.prototype.isSorterMode = function () {
		return true;
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
		}
	}

	window["AscCommon"] = window["AscCommon"] || {};
	window["AscCommon"].getViewModeByType = getViewModeByType;
})();
