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

(function(undefined) {
	function SlideModeManagerBase(api) {
		this.api = api;
		this.type = null;
	}

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
		if (!window["NATIVE_EDITOR_ENJINE"] && wordControl.Thumbnails) {
			return wordControl.Thumbnails.GetSelectedArray();
		} else {
			const presentation = this.getPresentation();
			return [presentation.CurPage];
		}
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
	SlideModeManagerBase.prototype.getSlidesCount = function() {
		return this.getAllSlides().length;
	};
	SlideModeManagerBase.prototype.getAllSlides = function() {
		return [];
	};
	SlideModeManagerBase.prototype.updateViewMode = function() {
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
	SlideModeManagerBase.prototype.getSlideNumber = function(object) {
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
	SlideModeManagerBase.prototype.applySlideProps = function(oProps, nDefaultLang, bAll) {
		return false
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

	function SlideModeManager(api) {
		SlideModeManagerBase.call(this, api);
		this.type = Asc.c_oAscPresentationViewMode.normal;
	}

	AscFormat.InitClassWithoutType(SlideModeManager, SlideModeManagerBase);
	SlideModeManager.prototype.updateViewMode = function() {
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
		const sizes = this.getSizesMM();
		const presentation = this.getPresentation();
		if (sizes.width > sizes.height && val === Asc.c_oAscPageOrientation.PagePortrait || sizes.width < sizes.height && val === Asc.c_oAscPageOrientation.PageLandscape) {
			presentation.changeSlideSize(sizes.height * AscCommonWord.g_dKoef_mm_to_emu, sizes.width * AscCommonWord.g_dKoef_mm_to_emu, sizes.type);
		}
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
		const presentation = this.getPresentation();
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
				const oNotes = oContent.Notes[i] || AscCommonSlide.CreateNotes();
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
				const oNotesMaster = oContent.NotesMastersIndexes[nNotesMasterIndex];
				const oNotesTheme = oContent.NotesThemes[nNotesMasterIndex];
				if (!oMaster.Theme) {
					oMaster.setTheme(oTheme);
				}
				if (!oLayout.Master) {
					oLayout.setMaster(oMaster);
				}
				if (!oNotes || !oNotesMaster || !oNotesTheme) {
					oSlide.setNotes(AscCommonSlide.CreateNotes());
					oSlide.notes.setNotesMaster(presentation.notesMasters[0]);
					oSlide.notes.setSlide(oSlide);
				} else {
					if (!oNotesMaster.Themes) {
						oNotesMaster.setTheme(oNotesTheme);
					}
					if (!oNotes.Master) {
						oNotes.setNotes(oNotesMaster);
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
	SlideModeManager.prototype.applySlideProps = function(oProps, nDefaultLang, bAll) {
		let oSlideProps = oProps.get_Slide();
		if (!oSlideProps) {
			return false;
		}
		const presentation = this.getPresentation();
		if (bAll) {
			var oMastersMap = {};
			for (let i = 0; i < presentation.Slides.length; ++i) {
				const oSlide = presentation.Slides[i];
				const oParents = oSlide.getParentObjects();
				const oMaster = oParents.master;
				const oLayout = oParents.layout;
				const bRemoveOnTitle = oLayout.type === AscFormat.nSldLtTTitle && presentation.showSpecialPlsOnTitleSld === false;
				if (oMaster) {
					if (!oMaster.hf) {
						oMaster.setHF(new AscFormat.HF());
					}
					oMaster.hf.applySettings(oSlideProps);
					if (oSlideProps.get_ShowSlideNum()) {
						const oSp = oSlide.getMatchingShape(AscFormat.phType_sldNum, null, false, {});
						if (!bRemoveOnTitle) {
							if (!oSp) {
								AscCommonSlide.copyPlaceholderToLikeObject(oSlide, oLayout, AscFormat.phType_sldNum);
							}
						} else {
							if (oSp) {
								oSlide.removeFromSpTreeById(oSp.Get_Id());
								oSp.setBDeleted(true);
							}
						}
					} else {
						const oSp = oSlide.getMatchingShape(AscFormat.phType_sldNum, null, false, {});
						if (oSp) {
							oSlide.removeFromSpTreeById(oSp.Get_Id());
							oSp.setBDeleted(true);
						}
					}

					if (oSlideProps.get_ShowFooter()) {
						const sText = oSlideProps.get_Footer();
						if (!oMastersMap[oMaster.Get_Id()]) {
							if (typeof sText === "string") {
								for (let j = 0; j < oMaster.sldLayoutLst.length; ++j) {
									AscCommonSlide.addFooterToSlideLikeObject(oMaster.sldLayoutLst[j], sText);
								}
								AscCommonSlide.addFooterToSlideLikeObject(oMaster, sText);
							}
						}
						const oSp = oSlide.getMatchingShape(AscFormat.phType_ftr, null, false, {});
						if (!bRemoveOnTitle) {
							if (!oSp) {
								AscCommonSlide.copyPlaceholderToLikeObject(oSlide, oLayout, AscFormat.phType_ftr);
							} else {
								AscCommonSlide.addFooterToSlideLikeObject(oSlide, sText);
							}
						} else {
							if (oSp) {
								oSlide.removeFromSpTreeById(oSp.Get_Id());
								oSp.setBDeleted(true);
							}
						}
					} else {
						const oSp = oSlide.getMatchingShape(AscFormat.phType_ftr, null, false, {});
						if (oSp) {
							oSlide.removeFromSpTreeById(oSp.Get_Id());
							oSp.setBDeleted(true);
						}
					}

					if (oSlideProps.get_ShowHeader()) {
						const sText = oSlideProps.get_Header();
						if (!oMastersMap[oMaster.Get_Id()]) {
							if (typeof sText === "string") {
								for (let j = 0; j < oMaster.sldLayoutLst.length; ++j) {
									AscCommonSlide.addHeaderToSlideLikeObject(oMaster.sldLayoutLst[j], sText);
								}
								AscCommonSlide.addHeaderToSlideLikeObject(oMaster, sText);
							}
						}


						const oSp = oSlide.getMatchingShape(AscFormat.phType_hdr, null, false, {});
						if (!bRemoveOnTitle) {
							if (!oSp) {
								AscCommonSlide.copyPlaceholderToLikeObject(oSlide, oLayout, AscFormat.phType_hdr);
							} else {
								AscCommonSlide.addHeaderToSlideLikeObject(oSlide, sText);
							}
						} else {
							if (oSp) {
								oSlide.removeFromSpTreeById(oSp.Get_Id());
								oSp.setBDeleted(true);
							}
						}
					} else {
						const oSp = oSlide.getMatchingShape(AscFormat.phType_hdr, null, false, {});
						if (oSp) {
							oSlide.removeFromSpTreeById(oSp.Get_Id());
							oSp.setBDeleted(true);
						}
					}

					if (oSlideProps.get_ShowDateTime()) {
						const oDateTime = oSlideProps.get_DateTime();

						let sDateTime = "";
						let sCustomDateTime = "";
						let nLang = nDefaultLang;
						if (oDateTime) {
							sDateTime = oDateTime.get_DateTime();
							sCustomDateTime = oDateTime.get_CustomDateTime();
							nLang = oDateTime.get_Lang();
							if (!oMastersMap[oMaster.Get_Id()]) {
								if (typeof sDateTime === "string" || typeof sCustomDateTime === "string") {
									if (sDateTime) {
										sCustomDateTime = oDateTime.get_DateTimeExamples()[sDateTime];
									}
									for (let j = 0; j < oMaster.sldLayoutLst.length; ++j) {
										AscCommonSlide.addDateTimeToSlideLikeObject(oMaster.sldLayoutLst[j], sDateTime, sCustomDateTime, nLang);
									}
									AscCommonSlide.addDateTimeToSlideLikeObject(oMaster, sDateTime, sCustomDateTime, nLang);
								}
							}
						}
						const oSp = oSlide.getMatchingShape(AscFormat.phType_dt, null, false, {});
						if (!bRemoveOnTitle) {
							if (!oSp) {
								AscCommonSlide.copyPlaceholderToLikeObject(oSlide, oLayout, AscFormat.phType_dt);
							} else {
								AscCommonSlide.addDateTimeToSlideLikeObject(oSlide, sDateTime, sCustomDateTime, oDateTime.get_Lang());
							}
						} else {
							if (oSp) {
								oSlide.removeFromSpTreeById(oSp.Get_Id());
								oSp.setBDeleted(true);
							}
						}
					} else {
						const oSp = oSlide.getMatchingShape(AscFormat.phType_dt, null, false, {});
						if (oSp) {
							oSlide.removeFromSpTreeById(oSp.Get_Id());
							oSp.setBDeleted(true);
						}
					}

					if (!oMastersMap[oMaster.Get_Id()]) {
						for (let nLayout = 0; nLayout < oMaster.sldLayoutLst.length; ++nLayout) {
							const oCurLayout = oMaster.sldLayoutLst[nLayout];
							if (oCurLayout.hf) {
								oCurLayout.setHF(null);
							}
						}
					}
					oMastersMap[oMaster.Get_Id()] = oMaster;
				}
			}
		} else {
			const aSelectedSlides = presentation.GetSelectedSlides();
			for (let nSlideIndex = 0; nSlideIndex < aSelectedSlides.length; ++nSlideIndex) {
				const oSlide = presentation.GetSlide(aSelectedSlides[nSlideIndex]);
				presentation.applyPropsToSlide(oSlide, oSlideProps);
			}
		}
		return true;
	};
	SlideModeManager.prototype.isSlideNoteShape = function(shape) {
		return shape.parent && shape.parent.isNote();
	};
	SlideModeManager.prototype.isDrawingSlide = function (slideObject) {
		return slideObject.isSlide();
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
		presentation.Recalculate({Drawings: {All: true, Map: {}}});
		wordControl.setNotesEnable(false);
		wordControl.setAnimPaneEnable(false);
		api.hideMediaControl();
		api.asc_hideComments();
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
		let isTitle = false;
		let isFooter = false;
		let isCanDeleteLayout = false;
		const curSlide = this.getCurrentSlide();
		const presentation = this.getPresentation();
		const api = this.getApi();
		if (curSlide && curSlide.isLayout()) {
			let shape = curSlide.getMatchingShape(AscFormat.phType_title, null, false, {});
			if (!shape) {
				shape = curSlide.getMatchingShape(AscFormat.phType_ctrTitle, null, false, {});
			}
			if (shape) {
				isTitle = true;
			}
			isFooter = true;
			let types = [AscFormat.phType_ftr, AscFormat.phType_dt, AscFormat.phType_sldNum];
			for (let nIdx = 0; nIdx < types.length; ++nIdx) {
				const type = types[nIdx];
				const shape = curSlide.getMatchingShape(type, null, false, {});
				if (!shape) {
					isFooter = false;
					break;
				}
			}


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
		api.sendEvent("asc_onLayoutTitle", isTitle);
		api.sendEvent("asc_onLayoutFooter", isFooter);
		api.sendEvent("asc_onCanDeleteMaster", presentation.CanRemoveMaster(presentation.lastMaster));
		api.sendEvent("asc_onCanDeleteLayout", isCanDeleteLayout);
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
		const presentation = this.getPresention();
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
			for (let nMaster = presentation.slideMasters.length; nMaster < presentation.slideMasters.length; ++nMaster) {
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
	MasterSlideModeManager.prototype.applySlideProps = function(oProps, nDefaultLang, bAll) {
		let oSlideProps = oProps.get_Slide();
		if (!oSlideProps) {
			return false;
		}
		const presentation = this.getPresentation();
		if (bAll) {
			for (let nMaster = 0; nMaster < presentation.slideMasters.length; ++nMaster) {
				const oMaster = presentation.slideMasters[nMaster];
				if (!oMaster.hf) {
					oMaster.setHF(new AscFormat.HF());
				}
				oMaster.hf.applySettings(oSlideProps);
				if (oSlideProps.get_ShowFooter()) {
					const sText = oSlideProps.get_Footer();
					if (typeof sText === "string") {
						for (let j = 0; j < oMaster.sldLayoutLst.length; ++j) {
							AscCommonSlide.addFooterToSlideLikeObject(oMaster.sldLayoutLst[j], sText);
						}
						AscCommonSlide.addFooterToSlideLikeObject(oMaster, sText);
					}
				}
				if (oSlideProps.get_ShowHeader()) {
					const sText = oSlideProps.get_Header();
					if (typeof sText === "string") {
						for (let j = 0; j < oMaster.sldLayoutLst.length; ++j) {
							AscCommonSlide.addHeaderToSlideLikeObject(oMaster.sldLayoutLst[j], sText);
						}
						AscCommonSlide.addHeaderToSlideLikeObject(oMaster, sText);
					}
				}
				if (oSlideProps.get_ShowDateTime()) {
					const oDateTime = oSlideProps.get_DateTime();

					let sDateTime = "";
					let sCustomDateTime = "";
					let nLang = nDefaultLang;
					if (oDateTime) {
						sDateTime = oDateTime.get_DateTime();
						sCustomDateTime = oDateTime.get_CustomDateTime();
						nLang = oDateTime.get_Lang();
						if (typeof sDateTime === "string" || typeof sCustomDateTime === "string") {
							if (sDateTime) {
								sCustomDateTime = oDateTime.get_DateTimeExamples()[sDateTime];
							}
							for (let j = 0; j < oMaster.sldLayoutLst.length; ++j) {
								AscCommonSlide.addDateTimeToSlideLikeObject(oMaster.sldLayoutLst[j], sDateTime, sCustomDateTime, nLang);
							}
							AscCommonSlide.addDateTimeToSlideLikeObject(oMaster, sDateTime, sCustomDateTime, nLang);
						}
					}
				}
				for (let nLayout = 0; nLayout < oMaster.sldLayoutLst.length; ++nLayout) {
					const oLayout = oMaster.sldLayoutLst[nLayout];
					if (oLayout.hf) {
						oLayout.setHF(null);
					}
				}
			}
			for (let nSlide = 0; nSlide < presentation.Slides.length; ++nSlide) {
				presentation.applyPropsToSlide(presentation.Slides[nSlide], oSlideProps);
			}
		} else {
			let aSelected = presentation.GetSelectedSlideObjects();
			let oDependentSlides = {};
			for (let nIdx = 0; nIdx < aSelected.length; ++nIdx) {
				let oSlideLikeObject = aSelected[nIdx];
				if (!oSlideLikeObject.hf) {
					oSlideLikeObject.setHF(new AscFormat.HF());
				}
				oSlideLikeObject.hf.applySettings(oSlideProps);
				if (oSlideProps.get_ShowFooter()) {
					const sText = oSlideProps.get_Footer();
					if (typeof sText === "string") {
						AscCommonSlide.addFooterToSlideLikeObject(oSlideLikeObject, sText);
					}
				}
				if (oSlideProps.get_ShowHeader()) {
					const sText = oSlideProps.get_Header();
					if (typeof sText === "string") {
						AscCommonSlide.addHeaderToSlideLikeObject(oSlideLikeObject, sText);
					}
				}
				if (oSlideProps.get_ShowDateTime()) {
					const oDateTime = oSlideProps.get_DateTime();
					let sDateTime = "";
					let sCustomDateTime = "";
					if (oDateTime) {
						sDateTime = oDateTime.get_DateTime();
						sCustomDateTime = oDateTime.get_CustomDateTime();
						if (typeof sDateTime === "string" || typeof sCustomDateTime === "string") {
							if (sDateTime) {
								sCustomDateTime = oDateTime.get_DateTimeExamples()[sDateTime];
							}
							AscCommonSlide.addDateTimeToSlideLikeObject(oSlideLikeObject, sDateTime, sCustomDateTime, oDateTime.get_Lang());
						}
					}
				}
				for (let nSlide = 0; nSlide < presentation.Slides.length; ++nSlide) {
					let oSlide = presentation.Slides[nSlide];
					let oParentObjects = oSlide.getParentObjects();
					if (oParentObjects.layout === oSlideLikeObject || oParentObjects.master === oSlideLikeObject) {
						oDependentSlides[oSlide.Id] = oSlide;
					}
				}
			}
			for (let sId in oDependentSlides) {
				if (oDependentSlides.hasOwnProperty(sId)) {
					let oSlide = oDependentSlides[sId];
					presentation.applyPropsToSlide(oSlide, oSlideProps);
				}
			}
		}
		return true;
	};
	MasterSlideModeManager.prototype.isDrawingSlide = function (slideObject) {
		return slideObject.isMaster() || slideObject.isLayout();
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
		const sizes = this.getSizesMM();
		const presentation = this.getPresentation();
		if (sizes.width > sizes.height && val === Asc.c_oAscPageOrientation.PagePortrait || sizes.width < sizes.height && val === Asc.c_oAscPageOrientation.PageLandscape) {
			presentation.changeNoteSize(sizes.height * AscCommonWord.g_dKoef_mm_to_emu, sizes.width * AscCommonWord.g_dKoef_mm_to_emu, sizes.type);
		}
	};
	NoteModeManager.prototype.applySlideProps = function(oProps, nDefaultLang, bAll) {
		const oNotesMastersMap = {};
		const presentation = this.getPresentation();
		const oNotesProps = oProps.get_Notes();
		for (let nSlide = 0; nSlide < presentation.notes.length; ++nSlide) {
			const oNotes = presentation.notes[nSlide];
			const oNotesMaster = oNotes.Master;
			if (!oNotesMaster) {
				continue;
			}
			if (!oNotesMaster.hf) {
				oNotesMaster.setHF(new AscFormat.HF());
			}
			oNotesMaster.hf.applySettings(oNotesProps);
			if (oNotesProps.get_ShowSlideNum()) {
				let oSp = oNotes.getMatchingShape(AscFormat.phType_sldNum, null, false, {});
				if (!oSp) {
					oSp = oNotesMaster.getMatchingShape(AscFormat.phType_sldNum, null, false, {});
					if (oSp) {
						oSp = oSp.copy(undefined);
						oSp.clearLang();
						oNotes.addToSpTreeToPos(undefined, oSp);
						oSp.setParent(oNotes);
					}
				}
			} else {
				const oSp = oNotes.getMatchingShape(AscFormat.phType_sldNum, null, false, {});
				if (oSp) {
					oNotes.removeFromSpTreeById(oSp.Get_Id());
					oSp.setBDeleted(true);
				}
			}

			if (oNotesProps.get_ShowFooter()) {
				const sText = oNotesProps.get_Footer();
				if (!oNotesMastersMap[oNotesMaster.Get_Id()]) {
					if (typeof sText === "string") {
						const oSp = oNotesMaster.getMatchingShape(AscFormat.phType_ftr, null, false, {});
						const oContent = oSp && oSp.getDocContent && oSp.getDocContent();
						if (oContent) {
							AscFormat.CheckContentTextAndAdd(oContent, sText);
						}
					}
				}
				const oSp = oNotes.getMatchingShape(AscFormat.phType_ftr, null, false, {});
				if (!oSp) {
					let oSp = oNotesMaster.getMatchingShape(AscFormat.phType_ftr, null, false, {});
					if (oSp) {
						oSp = oSp.copy(undefined);
						oSp.clearLang();
						oNotes.addToSpTreeToPos(undefined, oSp);
						oSp.setParent(oNotes);
					}
				} else {
					const oContent = oSp.getDocContent && oSp.getDocContent();
					if (oContent && typeof sText === "string") {
						AscFormat.CheckContentTextAndAdd(oContent, sText);
					}
				}
			} else {
				const oSp = oNotes.getMatchingShape(AscFormat.phType_ftr, null, false, {});
				if (oSp) {
					oNotes.removeFromSpTreeById(oSp.Get_Id());
					oSp.setBDeleted(true);
				}
			}

			if (oNotesProps.get_ShowHeader()) {
				const sText = oNotesProps.get_Header();
				if (!oNotesMastersMap[oNotesMaster.Get_Id()]) {
					AscCommonSlide.addHeaderToSlideLikeObject(oNotesMaster, sText);
				}


				let oSp = oNotes.getMatchingShape(AscFormat.phType_hdr, null, false, {});
				if (!oSp) {
					oSp = oNotesMaster.getMatchingShape(AscFormat.phType_hdr, null, false, {});
					if (oSp) {
						oSp = oSp.copy(undefined);
						oSp.clearLang();
						oNotes.addToSpTreeToPos(undefined, oSp);
						oSp.setParent(oNotes);
					}
				} else {
					AscCommonSlide.addHeaderToSlideLikeObject(oNotes, sText);
				}
			} else {
				const oSp = oNotes.getMatchingShape(AscFormat.phType_hdr, null, false, {});
				if (oSp) {
					oNotes.removeFromSpTreeById(oSp.Get_Id());
					oSp.setBDeleted(true);
				}
			}


			if (oNotesProps.get_ShowDateTime()) {
				const oDateTime = oNotesProps.get_DateTime();

				let sDateTime = "";
				let sCustomDateTime = "";
				let nLang = nDefaultLang;
				if (oDateTime) {
					sDateTime = oDateTime.get_DateTime();
					sCustomDateTime = oDateTime.get_CustomDateTime();
					nLang = oDateTime.get_Lang();
					if (!oNotesMastersMap[oNotesMaster.Get_Id()]) {
						if (typeof sDateTime === "string" || typeof sCustomDateTime === "string") {
							if (sDateTime) {
								sCustomDateTime = oDateTime.get_DateTimeExamples()[sDateTime];
							}
							AscCommonSlide.addDateTimeToSlideLikeObject(oNotesMaster, sDateTime, sCustomDateTime, nLang);
						}
					}
				}
				let oSp = oNotes.getMatchingShape(AscFormat.phType_dt, null, false, {});

				if (!oSp) {
					oSp = oNotesMaster.getMatchingShape(AscFormat.phType_dt, null, false, {});
					if (oSp) {
						oSp = oSp.copy(undefined);
						oSp.clearLang();
						oNotes.addToSpTreeToPos(undefined, oSp);
						oSp.setParent(oNotes);
					}
				} else {
					AscCommonSlide.addDateTimeToSlideLikeObject(oNotes, sDateTime, sCustomDateTime, nLang);
				}
			} else {
				const oSp = oNotes.getMatchingShape(AscFormat.phType_dt, null, false, {});
				if (oSp) {
					oNotes.removeFromSpTreeById(oSp.Get_Id());
					oSp.setBDeleted(true);
				}
			}
			oNotesMastersMap[oNotesMaster.Get_Id()] = oNotesMaster;
		}
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
	MasterNoteModeManager.prototype.setFooter = MasterHandoutModeManager.prototype.setFooter;
	MasterNoteModeManager.prototype.setHeader = MasterHandoutModeManager.prototype.setHeader;
	MasterNoteModeManager.prototype.setDate = MasterHandoutModeManager.prototype.setDate;
	MasterNoteModeManager.prototype.setNumber = MasterHandoutModeManager.prototype.setNumber;
	MasterNoteModeManager.prototype.setBody = function(val) {
		this.setPlaceholder(val, AscFormat.phType_body, AscCommonSlide.addBodyShape);
	};
	MasterNoteModeManager.prototype.setSlideImage = function(val) {
		this.setPlaceholder(val, AscFormat.phType_body, AscCommonSlide.addSlideImage);
	};
	MasterNoteModeManager.prototype.setPlaceholder = MasterHandoutModeManager.prototype.setPlaceholder;
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
	MasterHandoutModeManager.prototype.setPlaceholder = function(val, type, addCallback) {
		const slide = this.getCurrentSlide();
		if (!slide) {
			return;
		}
		this.startAction(0);
		let shape = slide.getMatchingShape(type, null, false, {});
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
	MasterHandoutModeManager.prototype.setFooter = function(val) {
		this.setPlaceholder(val, AscFormat.phType_ftr, AscCommonSlide.addFooterShape);
	};
	MasterHandoutModeManager.prototype.setHeader = function(val) {
		this.setPlaceholder(val, AscFormat.phType_hdr, AscCommonSlide.addHeaderShape);
	};
	MasterHandoutModeManager.prototype.setDate = function(val) {
		this.setPlaceholder(val, AscFormat.phType_dt, AscCommonSlide.addDateShape);
	};
	MasterHandoutModeManager.prototype.setNumber = function(val) {
		this.setPlaceholder(val, AscFormat.phType_sldNum, AscCommonSlide.addNumberShape);
	};
	MasterHandoutModeManager.prototype.setHandoutPageCount = function(val) {
		const slide = this.getCurrentSlide();
		slide && slide.setSlideCount(val);
	};
	MasterHandoutModeManager.prototype.getSizesMM = NoteModeManager.prototype.getSizesMM;
	MasterHandoutModeManager.prototype.setPageOrientation = NoteModeManager.prototype.setPageOrientation;
	MasterHandoutModeManager.prototype.applySlideProps = function() {
		//todo
	};
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
		AscFormat.redrawSlide(presentation.handoutMaster[aIdx[nStartIdx]], presentation, aIdx, nStartIdx, 0, this.getAllSlides());
	};

	function SorterModeManager(api) {
		SlideModeManagerBase.call(this, api);
		this.type = Asc.c_oAscPresentationViewMode.sorter;
	}

	AscFormat.InitClassWithoutType(SorterModeManager, SlideModeManagerBase);
	SorterModeManager.prototype.isSorterMode = function() {
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
