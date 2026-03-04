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

(function() {
	function SlideBase() {
		AscFormat.CBaseFormatObject.call(this);
		this.cSld = new AscFormat.CSld(this);
		this.clrMap = new AscFormat.ClrMap();
		this.presentation = editor && editor.WordControl && editor.WordControl.m_oLogicDocument;
		this.graphicObjects = new AscFormat.DrawingObjectsController(this);

		this.ImageBase64 = "";
		this.Width64 = 0;
		this.Height64 = 0;

		this.Width = 254;
		this.Height = 190.5;

		this.backgroundFill = null;

		this.recalcInfo = {
			recalculateBackground: true,
			recalculateSpTree:     true
		};

		this.deleteLock = new PropLocker(this.Id);
		this.backgroundLock = new PropLocker(this.Id);
		this.timingLock = new PropLocker(this.Id);
		this.transitionLock = new PropLocker(this.Id);
		this.layoutLock = new PropLocker(this.Id);
		this.showLock = new PropLocker(this.Id);
	}

	AscFormat.InitClass(SlideBase, AscFormat.CBaseFormatObject, 0);
	SlideBase.prototype.getDrawingDocument = function() {
		return editor.WordControl.m_oLogicDocument.DrawingDocument;
	};
	SlideBase.prototype.OnUpdateOverlay = function() {
		this.presentation.DrawingDocument.m_oWordControl.OnUpdateOverlay();
	};
	SlideBase.prototype.getNum = function() {
		let aSlides = this.presentation.GetAllSlides();
		for (let nIdx = 0; nIdx < aSlides.length; ++nIdx) {
			if (aSlides[nIdx] === this)
				return nIdx;
		}
		return -1;
	};
	SlideBase.prototype.drawSelect = function(_type) {
		if (_type === undefined) {
			this.graphicObjects.drawTextSelection(this.getNum());
			this.graphicObjects.drawSelect(0, this.presentation.DrawingDocument);
		} else if (_type == 1)
			this.graphicObjects.drawTextSelection(this.getNum());
		else if (_type == 2)
			this.graphicObjects.drawSelect(0, this.presentation.DrawingDocument);
	};
	SlideBase.prototype.showDrawingObjects = function() {
		editor.WordControl.m_oDrawingDocument.OnRecalculateSlide(this.getNum());
	};
	SlideBase.prototype.sendGraphicObjectProps = function() {
		editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
	};
	SlideBase.prototype.getMaster = function() {
		return this.getParentObjects().master;
	};
	SlideBase.prototype.addToRecalculate = function() {
		AscCommon.History.RecalcData_Add({Type: AscDFH.historyitem_recalctype_Drawing, Object: this});
	};
	SlideBase.prototype.convertPixToMM = function(pix) {
		return editor.WordControl.m_oDrawingDocument.GetMMPerDot(pix);
	};
	//todo
	SlideBase.prototype.shapeAdd = function(pos, item) {
		let pos_ = pos;
		if (!AscFormat.isRealNumber(pos)) {
			pos_ = this.cSld.spTree.length;
		}
		this.checkDrawingUniNvPr(item);
		AscCommon.History.Add(new AscDFH.CChangesDrawingsContent(this, AscDFH.historyitem_SlideLayoutAddToSpTree, pos_, [item], true));
		this.cSld.spTree.splice(pos_, 0, item);
		item.setParent2(this);
		this.recalcInfo.recalculateSpTree = true;
	};

	//todo check slides methods
	SlideBase.prototype.changeSize = function(width, height) {
		var kw = width / this.Width, kh = height / this.Height;
		this.setSlideSize(width, height);
		this.cSld.forEachSp(function(oSp) {
			oSp.changeSize(kw, kh);
		});
	};
	SlideBase.prototype.getAllRasterImages = function(images) {
		let oBgPr = this.cSld.Bg && this.cSld.Bg.bgPr;
		if (oBgPr) {
			oBgPr.checkBlipFillRasterImage(images);
		}
		this.cSld.forEachSp(function(oSp) {
			oSp.getAllRasterImages(images);
		});
	};
	SlideBase.prototype.Reassign_ImageUrls = function(images_rename) {
		this.cSld.forEachSp(function(oSp) {
			oSp.Reassign_ImageUrls(images_rename);
		});
		let bg = this.cSld && this.cSld.Bg;
		let bgPr = bg && bg.bgPr;
		let fill = bgPr && bgPr.Fill && bgPr.Fill.fill;

		if (fill instanceof AscFormat.CBlipFill &&
			typeof fill.RasterImageId === "string" &&
			images_rename[fill.RasterImageId]) {
			let oBg = bg.createFullCopy();
			oBg.bgPr.Fill.fill.RasterImageId = images_rename[oBg.bgPr.Fill.fill.RasterImageId];
			this.changeBackground(oBg);

			if (this.recalculateBackground) {
				this.recalculateBackground();
			}
		}
	};
	SlideBase.prototype.checkDrawingUniNvPr = function(drawing) {
		if (drawing) {
			drawing.checkDrawingUniNvPr();
		}
	};
	SlideBase.prototype.handleAllContents = function(fCallback) {
		this.cSld.handleAllContents(fCallback);
		if (this.notesShape) {
			this.notesShape.handleAllContents(fCallback);
		}
	};
	SlideBase.prototype.getAllRasterImagesForDraw = function(images) {
		let aImages = images;
		if (!aImages) {
			aImages = [];
		}

		if (this.recalcInfo.recalculateBackground) {
			this.recalculateBackground();
			this.recalcInfo.recalculateBackground = false;
		}
		if (this.backgroundFill) {
			let sImageId = this.backgroundFill.checkRasterImageId();
			if (sImageId) {
				aImages.push(sImageId);
			}
		}
		this.cSld.forEachSp(function(oSp) {
			oSp.getAllRasterImages(aImages);
		});
		if (this.Layout) {
			if (this.needLayoutSpDraw()) {
				this.Layout.getAllRasterImagesForDraw(aImages);
			}
			if (this.Layout.Master) {
				if (this.needMasterSpDraw()) {
					this.Layout.Master.getAllRasterImagesForDraw(aImages);
				}
			}
		}
		return aImages;
	};
	SlideBase.prototype.checkImageDraw = function(sImageSrc) {
		const aImages = this.getAllRasterImagesForDraw();
		for (let nIdx = 0; nIdx < aImages.length; ++nIdx) {
			let sImage = aImages[nIdx];
			if (AscCommon.getFullImageSrc2(sImage) === sImageSrc) {
				return true;
			}
		}
		return false;
	};
	SlideBase.prototype.openChartEditor = function() {
		this.graphicObjects.openChartEditor();
	};
	SlideBase.prototype.openOleEditor = function() {
		this.graphicObjects.openOleEditor();
	};
	SlideBase.prototype.getBase64Img = function() {
		if (window["NATIVE_EDITOR_ENJINE"]) {
			return "";
		}
		if (typeof this.cachedImage === "string" && this.cachedImage.length > 0)
			return this.cachedImage;

		AscCommon.IsShapeToImageConverter = true;

		var dKoef = AscCommon.g_dKoef_mm_to_pix;
		var _need_pix_width = ((this.Width * dKoef / 3.0 + 0.5) >> 0);
		var _need_pix_height = ((this.Height * dKoef / 3.0 + 0.5) >> 0);

		if (_need_pix_width <= 0 || _need_pix_height <= 0)
			return null;

		var _canvas = document.createElement('canvas');
		_canvas.width = _need_pix_width;
		_canvas.height = _need_pix_height;

		var _ctx = _canvas.getContext('2d');

		var g = new AscCommon.CGraphics();
		g.init(_ctx, _need_pix_width, _need_pix_height, this.Width, this.Height);
		g.m_oFontManager = AscCommon.g_fontManager;

		g.m_oCoordTransform.tx = 0.0;
		g.m_oCoordTransform.ty = 0.0;
		g.transform(1, 0, 0, 1, 0, 0);

		this.draw(g, /*pageIndex*/0);

		if (AscCommon.g_fontManager) {
			AscCommon.g_fontManager.m_pFont = null;
		}
		if (AscCommon.g_fontManager2) {
			AscCommon.g_fontManager2.m_pFont = null;
		}
		AscCommon.IsShapeToImageConverter = false;

		var _ret = {ImageNative: _canvas, ImageUrl: ""};
		try {
			_ret.ImageUrl = _canvas.toDataURL("image/png");
		} catch (err) {
			_ret.ImageUrl = "";
		}
		return _ret.ImageUrl;
	};
	SlideBase.prototype.getMatchingShape = function(type, idx, bSingleBody, info) {
		var _input_reduced_type;
		if (type == null) {
			_input_reduced_type = AscFormat.phType_body;
		} else {
			if (type == AscFormat.phType_ctrTitle) {
				_input_reduced_type = AscFormat.phType_title;
			} else {
				_input_reduced_type = type;
			}
		}

		var _input_reduced_index;
		if (idx == null) {
			_input_reduced_index = 0;
		} else {
			_input_reduced_index = idx;
		}

		var _sp_tree = this.cSld.spTree;
		var _shape_index;
		var _index, _type;
		var _final_index, _final_type;
		var _glyph;
		var body_count = 0;
		var last_body;
		var oNvObjPr;
		for (_shape_index = 0; _shape_index < _sp_tree.length; ++_shape_index) {
			_glyph = _sp_tree[_shape_index];
			if (_glyph.isPlaceholder()) {
				oNvObjPr = _glyph.getUniNvProps();
				if (oNvObjPr) {
					_index = oNvObjPr.nvPr.ph.idx;
					_type = oNvObjPr.nvPr.ph.type;
					if (_type == null) {
						_final_type = AscFormat.phType_body;
					} else {
						if (_type == AscFormat.phType_ctrTitle) {
							_final_type = AscFormat.phType_title;
						} else {
							_final_type = _type;
						}
					}

					if (_index == null) {
						_final_index = 0;
					} else {
						_final_index = _index;
					}

					if (_input_reduced_type == _final_type && _input_reduced_index == _final_index) {
						if (info) {
							info.bBadMatch = !(_type === type && _index === idx);
						}
						return _glyph;
					}
					if (_input_reduced_type == AscFormat.phType_title && _input_reduced_type == _final_type) {
						if (info) {
							info.bBadMatch = !(_type === type && _index === idx);
						}
						return _glyph;
					}
					if (AscFormat.phType_body === _type) {
						++body_count;
						last_body = _glyph;
					}
				}
			}
		}


		if (_input_reduced_type == AscFormat.phType_sldNum || _input_reduced_type == AscFormat.phType_dt || _input_reduced_type == AscFormat.phType_ftr || _input_reduced_type == AscFormat.phType_hdr) {
			for (_shape_index = 0; _shape_index < _sp_tree.length; ++_shape_index) {
				_glyph = _sp_tree[_shape_index];
				if (_glyph.isPlaceholder()) {
					oNvObjPr = _glyph.getUniNvProps();
					if (oNvObjPr) {
						_type = oNvObjPr.nvPr.ph.type;
						if (_input_reduced_type == _type) {
							if (info) {
								info.bBadMatch = !(_type === type && _index === idx);
							}
							return _glyph;
						}
					}
				}
			}
		}

		if (info) {
			return null;
		}
		if (body_count === 1 && _input_reduced_type === AscFormat.phType_body && bSingleBody) {
			if (info) {
				info.bBadMatch = !(_type === type && _index === idx);
			}
			return last_body;
		}

		for (_shape_index = 0; _shape_index < _sp_tree.length; ++_shape_index) {
			_glyph = _sp_tree[_shape_index];
			if (_glyph.isPlaceholder()) {
				if (_glyph instanceof AscFormat.CShape) {
					_index = _glyph.nvSpPr.nvPr.ph.idx;
					_type = _glyph.nvSpPr.nvPr.ph.type;
				}
				if (_glyph instanceof AscFormat.CImageShape) {
					_index = _glyph.nvPicPr.nvPr.ph.idx;
					_type = _glyph.nvPicPr.nvPr.ph.type;
				}
				if (_glyph instanceof AscFormat.CGroupShape) {
					_index = _glyph.nvGrpSpPr.nvPr.ph.idx;
					_type = _glyph.nvGrpSpPr.nvPr.ph.type;
				}

				if (_index == null) {
					_final_index = 0;
				} else {
					_final_index = _index;
				}

				if (_input_reduced_index == _final_index) {
					if (info) {
						info.bBadMatch = true;
					}
					return _glyph;
				}
			}
		}


		if (body_count === 1 && bSingleBody) {
			if (info) {
				info.bBadMatch = !(_type === type && _index === idx);
			}
			return last_body;
		}

		return null;
	};
	SlideBase.prototype.checkSlideSize = function() {
		this.recalcInfo.recalculateSpTree = true;
		this.cSld.forEachSp(function(oSp) {
			oSp.handleUpdateExtents();
		});
	};
	SlideBase.prototype.recalcText = function() {
		this.recalcInfo.recalculateSpTree = true;
		this.cSld.forEachSp(function(oSp) {
			oSp.recalcText();
		});
	};
	SlideBase.prototype.checkSlideTheme = function() {

		this.bChangeLayout = undefined;
		this.recalcInfo.recalculateSpTree = true;
		this.recalcInfo.recalculateBackground = true;
		this.cSld.forEachSp(function(oSp) {
			oSp.handleUpdateTheme();
		});
	};
	SlideBase.prototype.copySelectedObjects = function() {
		var aSelectedObjects, i, fShift = 5.0;
		var oSelector = this.graphicObjects.selection.groupSelection ? this.graphicObjects.selection.groupSelection : this.graphicObjects;
		aSelectedObjects = [].concat(oSelector.selectedObjects);
		oSelector.resetSelection(undefined, false);
		var bGroup = !!this.graphicObjects.selection.groupSelection;
		if (bGroup) {
			oSelector.normalize();
		}
		for (i = 0; i < aSelectedObjects.length; ++i) {
			var oCopy = aSelectedObjects[i].copy(undefined);
			oCopy.x = aSelectedObjects[i].x;
			oCopy.y = aSelectedObjects[i].y;
			oCopy.extX = aSelectedObjects[i].extX;
			oCopy.extY = aSelectedObjects[i].extY;
			AscFormat.CheckSpPrXfrm(oCopy, true);
			oCopy.spPr.xfrm.setOffX(oCopy.x + fShift);
			oCopy.spPr.xfrm.setOffY(oCopy.y + fShift);
			oCopy.setParent(this);
			if (!bGroup) {
				this.addToSpTreeToPos(undefined, oCopy);
			} else {
				oCopy.setGroup(aSelectedObjects[i].group);
				aSelectedObjects[i].group.addToSpTree(undefined, oCopy);
			}
			oSelector.selectObject(oCopy, 0);
		}
		if (bGroup) {
			oSelector.updateCoordinatesAfterInternalResize();
		}
	};
	SlideBase.prototype.getPlaceholdersControls = function() {
		var ret = [];
		var aSpTree = this.cSld.spTree;
		for (var i = 0; i < aSpTree.length; ++i) {
			var oSp = aSpTree[i];
			oSp.createPlaceholderControl(ret);
		}
		return ret;
	};
	SlideBase.prototype.getDrawingObjects = function() {
		return this.cSld.spTree;
	};
	SlideBase.prototype.removeFromSpTreeById = function(id) {
		var sp_tree = this.cSld.spTree;
		for (var i = 0; i < sp_tree.length; ++i) {
			if (sp_tree[i].Get_Id() === id) {
				this.removeFromSpTreeByPos(i);
				return i;
			}
		}
		return null;
	};
	SlideBase.prototype.removeFromSpTreeByPos = function(pos) {
		if (pos > -1 && pos < this.cSld.spTree.length) {
			var oSp = this.cSld.spTree[pos];
			if (oSp.isPlaceholder() || this.isMaster() || this.isLayout()) {
				let oMap = {};
				oMap[oSp.Id] = oSp;
				let oPres = Asc.editor.private_GetLogicDocument();
				let aSlides = oPres.Slides;
				if (this.isMaster()) {
					for (let nLt = 0; nLt < this.sldLayoutLst.length; ++nLt) {
						let oLt = this.sldLayoutLst[nLt];
						oLt.cSld.forEachSp(function(oSp) {
							oSp.checkOnDeletePlaceholder(oMap);
						});
					}
					for (let nSld = 0; nSld < aSlides.length; ++nSld) {
						let oSld = aSlides[nSld];
						if (oSld.getMaster() === this) {
							oSld.cSld.forEachSp(function(oSp) {
								oSp.checkOnDeletePlaceholder(oMap);
							});
						}
					}
				}
				if (this.isLayout()) {
					for (let nSld = 0; nSld < aSlides.length; ++nSld) {
						let oSld = aSlides[nSld];
						if (oSld.Layout === this) {
							oSld.cSld.forEachSp(function(oSp) {
								oSp.checkOnDeletePlaceholder(oMap);
							});
						}
					}
				}
			}
			this.shapeRemove(pos, 1);
			if (this.timing && !AscCommon.IsChangingDrawingZIndex) {
				this.checkNeedCopyTimingBeforeEdit();
				this.timing.onRemoveObject(oSp.Get_Id());
			}
			if (this.collaborativeMarks) {
				this.collaborativeMarks.Update_OnRemove(pos, 1);
			}

			return oSp;
		}
		return null;
	};
	SlideBase.prototype.RestartSpellCheck = function() {
		for (var i = 0; i < this.cSld.spTree.length; ++i) {
			this.cSld.spTree[i].RestartSpellCheck();
		}
		if (this.notes) {
			var spTree = this.notes.cSld.spTree;
			for (i = 0; i < spTree.length; ++i) {
				spTree[i].RestartSpellCheck();
			}
		}
	};
	SlideBase.prototype.Search = function(Engine, Type) {
		var sp_tree = this.cSld.spTree;
		for (var i = 0; i < sp_tree.length; ++i) {
			if (sp_tree[i].Search) {
				sp_tree[i].Search(Engine, Type);
			}
		}
		if (this.notesShape) {
			this.notesShape.Search(Engine, Type);
		}
	};
	SlideBase.prototype.GetSearchElementId = function(isNext, StartPos) {
		var sp_tree = this.cSld.spTree, i, Id;
		if (isNext) {
			for (i = StartPos; i < sp_tree.length; ++i) {
				if (sp_tree[i].GetSearchElementId) {
					Id = sp_tree[i].GetSearchElementId(isNext, false);
					if (Id !== null) {
						return Id;
					}
				}
			}
		} else {
			for (i = StartPos; i > -1; --i) {
				if (sp_tree[i].GetSearchElementId) {
					Id = sp_tree[i].GetSearchElementId(isNext, false);
					if (Id !== null) {
						return Id;
					}
				}
			}
		}
		return null;
	};
	SlideBase.prototype.replaceSp = function(oPh, oObject) {
		var aSpTree = this.cSld.spTree;
		for (var i = 0; i < aSpTree.length; ++i) {
			if (aSpTree[i] === oPh) {
				break;
			}
		}
		this.removeFromSpTreeByPos(i);
		this.addToSpTreeToPos(i, oObject);
		var oNvProps = oObject.getNvProps && oObject.getNvProps();
		if (oNvProps) {
			if (oPh) {
				var oNvPropsPh = oPh.getNvProps && oPh.getNvProps();
				var oPhPr = oNvPropsPh.ph;
				if (oPhPr) {
					oNvProps.setPh(oPhPr.createDuplicate());
				}
			}
		}
	};
	SlideBase.prototype.drawViewPrMarks = function(oGraphics) {
		let oContext = oGraphics.m_oContext;
		if (!oContext ||
			AscCommon.IsShapeToImageConverter ||
			oGraphics.animationDrawer ||
			oGraphics.IsThumbnail ||
			oGraphics.IsDemonstrationMode ||
			oGraphics.isBoundsChecker() ||
			oGraphics.IsNoDrawingEmptyPlaceholder) {
			return;
		}

		let oApi = editor;
		if (!oApi) {
			return;
		}
		if (!oApi.WordControl) {
			return;
		}
		let oPresentation = oApi.WordControl.m_oLogicDocument;
		if (!oPresentation) {
			return;
		}
		if (oApi.asc_getShowGridlines()) {
			oPresentation.checkGridCache(oGraphics);
		}
		if (oApi.asc_getShowGuides()) {
			oPresentation.drawGuides(oGraphics);
		}
	};
	SlideBase.prototype.removeAllInks = function() {
		this.cSld.removeAllInks();
	};
	SlideBase.prototype.getAllInks = function(arrInks) {
		arrInks = arrInks || [];
		this.cSld.getAllInks(arrInks);
		return arrInks;
	};

	// todo end check slides methods

	SlideBase.prototype.getDrawingsForController = function() {
	};
	SlideBase.prototype.getTheme = function() {
	};
	SlideBase.prototype.getColorMap = function() {
	};
	SlideBase.prototype.recalculateBackground = function() {
	};
	SlideBase.prototype.checkSlideColorScheme = function() {
	};
	SlideBase.prototype.getAllFonts = function(fonts) {
	};
	SlideBase.prototype.createFontMap = function(oFontsMap, oCheckedMap) {
	};
	SlideBase.prototype.recalculate = function() {
	};
	SlideBase.prototype.isAnimated = function() {
		return false;
	};
	SlideBase.prototype.isLockedObject = function() {
		return false;
	};
	SlideBase.prototype.needRecalc = function() {
		if (this.recalcInfo) {
			for (let flag in this.recalcInfo) {
				if (flag) {
					return true;
				}
			}
		}
		return false;
	};
	SlideBase.prototype.Clear_ContentChanges = function() {
	};
	SlideBase.prototype.Add_ContentChanges = function(Changes) {
	};
	SlideBase.prototype.Refresh_ContentChanges = function() {
	};
	SlideBase.prototype.checkSlideColorScheme = function() {
	};
	SlideBase.prototype.isSlide = function() {
		return false;
	};
	SlideBase.prototype.isLayout = function() {
		return false;
	};
	SlideBase.prototype.isMaster = function() {
		return false;
	};
	SlideBase.prototype.getParentObjects = function() {
	};
	SlideBase.prototype.recalculateNotesShape = function() {
	};
	SlideBase.prototype.getNotesHeight = function() {
		return 0;
	};
	SlideBase.prototype.isVisible = function() {
		return true;
	};
	SlideBase.prototype.getWorksheet = function() {
		return null;
	};
	SlideBase.prototype.Get_ColorMap = function() {
		return AscFormat.GetDefaultColorMap();
	};
	SlideBase.prototype.IsUseInDocument = function() {
		return false;
	};
	SlideBase.prototype.IsUseInSlides = function() {
		return false
	};
	SlideBase.prototype.isPreserve = function() {
		return false
	};
	SlideBase.prototype.draw = function(graphics, slide) {
	};
	SlideBase.prototype.shapeRemove = function(pos, count) {
	};
	SlideBase.prototype.shapeAdd = function(pos, item) {
	};
	SlideBase.prototype.addToSpTreeToPos = function(pos, item) {
		this.shapeAdd(pos, item);
	};
	SlideBase.prototype.setSlideSize = function(w, h) {
	};
	SlideBase.prototype.changeBackground = function(bg) {
	};

	function fAddTextToPhInSlideLikeObject(oSlideLikeObject, nPhType, sText) {
		if(typeof sText === "string") {
			let oSp = oSlideLikeObject.getMatchingShape(nPhType, null, false, {});
			let oContent = oSp && oSp.getDocContent && oSp.getDocContent();
			if (oContent) {
				let oParaPr = null;
				let oFirstParaOld = oContent.GetFirstParagraph();
				if (oFirstParaOld) {
					oParaPr = oFirstParaOld.GetDirectParaPr();
				}
				AscFormat.CheckContentTextAndAdd(oContent, sText);
				if(oParaPr) {
					let oFirstParaNew = oContent.GetFirstParagraph();
					if (oFirstParaNew) {
						oFirstParaNew.SetPr(oParaPr);
					}
				}
			}
		}
	}
	function addFooterToSlideLikeObject(oSlideLikeObject, sFooterText) {
		fAddTextToPhInSlideLikeObject(oSlideLikeObject, AscFormat.phType_ftr, sFooterText);
	}
	function addHeaderToSlideLikeObject(oSlideLikeObject, sHeaderText) {
		fAddTextToPhInSlideLikeObject(oSlideLikeObject, AscFormat.phType_hdr, sHeaderText);
	}
	function addDateTimeToSlideLikeObject(oSlideLikeObject, sDateTimeFieldType, sCustomDateTime, nLang) {

		let oSp = oSlideLikeObject.getMatchingShape(AscFormat.phType_dt, null, false, {});
		if (oSp) {
			let oContent = oSp.getDocContent && oSp.getDocContent();
			if (oContent) {
				if (sDateTimeFieldType) {
					oContent.ClearContent(true);
					let oParagraph = oContent.Content[0];
					let oFld = new AscCommonWord.CPresentationField(oParagraph);
					oFld.SetGuid(AscCommon.CreateGUID());
					oFld.SetFieldType(sDateTimeFieldType);
					if(AscFormat.isRealNumber(nLang)) {
						oFld.Set_Lang_Val(nLang);
					}
					else {
					const presentation = oSlideLikeObject.presentation;
					const nDefaultLang = presentation && presentation.GetDefaultLanguage() || null;
						oFld.Set_Lang_Val(nDefaultLang);
					}
					if (typeof sCustomDateTime === "string") {
						oFld.CanAddToContent = true;
						oFld.AddText(sCustomDateTime);
						oFld.CanAddToContent = false;
					}
					oParagraph.Internal_Content_Add(0, oFld);
				} else {
					AscFormat.CheckContentTextAndAdd(oContent, sCustomDateTime);
				}
			}
		}
	}
	function copyPlaceholderToLikeObject(oSlide, oTemplate, nPhType) {
		let oSp = oTemplate.getMatchingShape(nPhType, null, false, {});
		if (oSp) {
			oSp = oSp.copy(undefined);
			oSp.clearLang();
			oSlide.addToSpTreeToPos(undefined, oSp);
			oSp.setParent(oSlide);
		}
	}

	const PLACEHOLDERSHAPE_WIDTH_COEFFICIENT = 2971800 / 6858000;
	const PLACEHOLDERSHAPE_HEIGHT_COEFFICIENT = 458787 / 9144000;

	function addHeaderShape(slideObject, x, y, width, height) {
		var oNvSpPr = new AscFormat.UniNvPr();
		var oCNvPr = oNvSpPr.cNvPr;
		oCNvPr.setId(2);
		oCNvPr.setName("Header Placeholder 1");
		var oPh = new AscFormat.Ph();
		oPh.setType(AscFormat.phType_hdr);
		oPh.setSz(2);
		oNvSpPr.nvPr.setPh(oPh);
		const notesHeight = slideObject.Height;
		const notesWidth = slideObject.Width;
		const placeholderWidth = notesWidth * PLACEHOLDERSHAPE_WIDTH_COEFFICIENT;
		const placeholderHeight = notesHeight * PLACEHOLDERSHAPE_HEIGHT_COEFFICIENT;
		var oSp = createPlaceholderShape(oNvSpPr, 0, 0, placeholderWidth, placeholderHeight, AscCommon.align_Left, AscFormat.VERTICAL_ANCHOR_TYPE_TOP);
		oSp.setParent(slideObject);
		slideObject.addToSpTreeToPos(0, oSp);
	}

	function addDateShape(slideObject) {
		const oNvSpPr = new AscFormat.UniNvPr();
		const oCNvPr = oNvSpPr.cNvPr;
		oCNvPr.setId(3);
		oCNvPr.setName("Date Placeholder 2");
		const oPh = new AscFormat.Ph();
		oPh.setType(AscFormat.phType_dt);
		oPh.setIdx(3 + "");
		oNvSpPr.nvPr.setPh(oPh);
		const notesHeight = slideObject.Height;
		const notesWidth = slideObject.Width;
		const placeholderWidth = notesWidth * PLACEHOLDERSHAPE_WIDTH_COEFFICIENT;
		const placeholderHeight = notesHeight * PLACEHOLDERSHAPE_HEIGHT_COEFFICIENT;
		const oSp = createPlaceholderShape(oNvSpPr, notesWidth - placeholderWidth, 0, placeholderWidth, placeholderHeight, AscCommon.align_Right, AscFormat.VERTICAL_ANCHOR_TYPE_TOP);
		oSp.setParent(slideObject);
		slideObject.addToSpTreeToPos(1, oSp);
		AscCommonSlide.addDateTimeToSlideLikeObject(slideObject, "datetimeFigureOut");
	}

	function addFooterShape(slideObject) {
		const oNvSpPr = new AscFormat.UniNvPr();
		const oCNvPr = oNvSpPr.cNvPr;
		oCNvPr.setId(4);
		oCNvPr.setName("Footer Placeholder 3");
		const oPh = new AscFormat.Ph();
		oPh.setType(AscFormat.phType_ftr);
		oPh.setIdx(2 + "");
		oPh.setSz(2);
		oNvSpPr.nvPr.setPh(oPh);
		const notesHeight = slideObject.Height;
		const notesWidth = slideObject.Width;
		const placeholderWidth = notesWidth * PLACEHOLDERSHAPE_WIDTH_COEFFICIENT;
		const placeholderHeight = notesHeight * PLACEHOLDERSHAPE_HEIGHT_COEFFICIENT;
		const oSp = createPlaceholderShape(oNvSpPr, 0, notesHeight - placeholderHeight, placeholderWidth, placeholderHeight, AscCommon.align_Left, AscFormat.VERTICAL_ANCHOR_TYPE_BOTTOM);
		oSp.setParent(slideObject);
		slideObject.addToSpTreeToPos(2, oSp);
	}

	function addNumberShape(slideObject) {
		const oNvSpPr = new AscFormat.UniNvPr();
		const oCNvPr = oNvSpPr.cNvPr;
		oCNvPr.setId(5);
		oCNvPr.setName("Slide Number Placeholder 4");
		const oPh = new AscFormat.Ph();
		oPh.setType(AscFormat.phType_sldNum);
		oPh.setIdx(3 + "");
		oPh.setSz(2);
		oNvSpPr.nvPr.setPh(oPh);
		const notesHeight = slideObject.Height;
		const notesWidth = slideObject.Width;
		const placeholderWidth = notesWidth * PLACEHOLDERSHAPE_WIDTH_COEFFICIENT;
		const placeholderHeight = notesHeight * PLACEHOLDERSHAPE_HEIGHT_COEFFICIENT;
		const oSp = createPlaceholderShape(oNvSpPr, notesWidth - placeholderWidth, notesHeight - placeholderHeight, placeholderWidth, placeholderHeight, AscCommon.align_Right, AscFormat.VERTICAL_ANCHOR_TYPE_BOTTOM);
		const oContent = oSp.getDocContent();
		if(oContent) {
			oContent.ClearContent(true);
			const oParagraph = oContent.Content[0];
			const oFld = new AscCommonWord.CPresentationField(oParagraph);
			oFld.SetGuid(AscCommon.CreateGUID());
			oFld.SetFieldType("slidenum");
			oParagraph.Internal_Content_Add(0, oFld);
		}
		oSp.setParent(slideObject);
		slideObject.addToSpTreeToPos(3, oSp);
	}

	function createPlaceholderShape(oNvSpPr, x, y, width, height, textAlign, bodyPrAlign) {
		const oSp = new AscFormat.CShape();
		oSp.setBDeleted(false);
		oSp.setNvSpPr(oNvSpPr);
		oSp.setLockValue(AscFormat.LOCKS_MASKS.noGrp, true);
		oSp.setSpPr(new AscFormat.CSpPr());
		oSp.spPr.setParent(oSp);
		oSp.spPr.setXfrm(new AscFormat.CXfrm());
		oSp.spPr.xfrm.setParent(oSp.spPr);
		oSp.spPr.xfrm.setOffX(x);
		oSp.spPr.xfrm.setOffY(y);
		oSp.spPr.xfrm.setExtX(width);
		oSp.spPr.xfrm.setExtY(height);
		oSp.spPr.setGeometry(AscFormat.CreateGeometry("rect"));
		oSp.spPr.geometry.setParent(oSp.spPr);
		oSp.createTextBody();
		const oBodyPr = oSp.txBody.bodyPr.createDuplicate();
		oBodyPr.vert = AscFormat.nVertTThorz;
		oBodyPr.lIns = 91440 / 36000;
		oBodyPr.tIns = 45720 / 36000;
		oBodyPr.rIns = 91440 / 36000;
		oBodyPr.bIns = 45720 / 36000;
		oBodyPr.anchor = bodyPrAlign;
		oSp.txBody.setBodyPr(oBodyPr);
		const oTxLstStyle = new AscFormat.TextListStyle();
		oTxLstStyle.levels[0] = new CParaPr();
		oTxLstStyle.levels[0].Jc = textAlign;
		oTxLstStyle.levels[0].DefaultRunPr = new AscCommonWord.CTextPr();
		oTxLstStyle.levels[0].DefaultRunPr.FontSize = 12;
		//endParaPr
		oSp.txBody.setLstStyle(oTxLstStyle);
		return oSp;
	}

//--------------------------------------------------------export----------------------------------------------------
	window['AscCommonSlide'] = window['AscCommonSlide'] || {};
	window['AscCommonSlide'].SlideBase = SlideBase;
	window['AscCommonSlide'].addFooterToSlideLikeObject = addFooterToSlideLikeObject;
	window['AscCommonSlide'].addHeaderToSlideLikeObject = addHeaderToSlideLikeObject;
	window['AscCommonSlide'].addDateTimeToSlideLikeObject = addDateTimeToSlideLikeObject;
	window['AscCommonSlide'].copyPlaceholderToLikeObject = copyPlaceholderToLikeObject;
	window['AscCommonSlide'].addHeaderShape = addHeaderShape;
	window['AscCommonSlide'].addDateShape = addDateShape;
	window['AscCommonSlide'].addFooterShape = addFooterShape;
	window['AscCommonSlide'].addNumberShape = addNumberShape;
})();
