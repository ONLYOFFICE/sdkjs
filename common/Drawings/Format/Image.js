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

(
	/**
	 * @param {Window} window
	 * @param {undefined} undefined
	 */
	function (window, undefined) {
// Import
		var CShape = AscFormat.CShape;
		var global_MatrixTransformer = AscCommon.global_MatrixTransformer;

		var isRealObject = AscCommon.isRealObject;

		AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetSpPr] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetBlipFill] = AscDFH.CChangesDrawingsObjectNoId;
		AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetParent] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetGroup] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetStyle] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_ImageShapeSetNvPicPr] = AscDFH.CChangesDrawingsObject;

		AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetSpPr] = function (oClass, value) {
			oClass.spPr = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetBlipFill] = function (oClass, value, FromLoad) {
			oClass.blipFill = value;
			if (FromLoad) {
				if (typeof AscCommon.CollaborativeEditing !== "undefined") {
					if (value && (typeof value.RasterImageId === "string") && value.RasterImageId.length > 0) {
						AscCommon.CollaborativeEditing.Add_NewImage(value.RasterImageId);
					}
				}
			}
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetParent] = function (oClass, value) {
			oClass.oldParent = oClass.parent;
			oClass.parent = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetGroup] = function (oClass, value) {
			oClass.group = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetStyle] = function (oClass, value) {
			oClass.style = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_ImageShapeSetNvPicPr] = function (oClass, value) {
			oClass.nvPicPr = value;
		};


		AscDFH.drawingsConstructorsMap[AscDFH.historyitem_ImageShapeSetBlipFill] = AscFormat.CBlipFill;


		function CImageShape() {
			AscFormat.CGraphicObjectBase.call(this);
			this.nvPicPr = null;
			this.blipFill = null;
			this.style = null;

			this.cropBrush = false;
			this.isCrop = false;
			this.parentCrop = null;

			this.shdwSp = null;
		}

		AscFormat.InitClass(CImageShape, AscFormat.CGraphicObjectBase, AscDFH.historyitem_type_ImageShape);
		CImageShape.prototype.setNvPicPr = function (pr) {
			AscCommon.History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ImageShapeSetNvPicPr, this.nvPicPr, pr));
			this.nvPicPr = pr;
		};

		CImageShape.prototype.setSpPr = function (pr) {
			AscCommon.History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ImageShapeSetSpPr, this.spPr, pr));
			this.spPr = pr;

			if (pr) {
				pr.setParent(this);
			}
		};

		CImageShape.prototype.setBlipFill = function (pr) {
			AscCommon.History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_ImageShapeSetBlipFill, this.blipFill, pr));
			this.blipFill = pr;
		};

		CImageShape.prototype.setParent = function (pr) {
			AscCommon.History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ImageShapeSetParent, this.parent, pr));
			this.parent = pr;
		};

		CImageShape.prototype.setGroup = function (pr) {
			AscCommon.History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ImageShapeSetGroup, this.group, pr));
			this.group = pr;
		};

		CImageShape.prototype.setStyle = function (pr) {
			AscCommon.History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ImageShapeSetStyle, this.style, pr));
			this.style = pr;
		};

		CImageShape.prototype.copy = function (oPr) {
			var copy = new CImageShape();

			if (this.nvPicPr) {
				copy.setNvPicPr(this.nvPicPr.createDuplicate());
			}
			if (this.spPr) {
				copy.setSpPr(this.spPr.createDuplicate());
				copy.spPr.setParent(copy);
			}
			if (this.blipFill) {
				copy.setBlipFill(this.blipFill.createDuplicate());
			}
			if (this.style) {
				copy.setStyle(this.style.createDuplicate());
			}
			copy.setBDeleted(this.bDeleted);
			if (this.macro !== null) {
				copy.setMacro(this.macro);
			}
			if (this.textLink !== null) {
				copy.setTextLink(this.textLink);
			}
			if (this.clientData) {
				copy.setClientData(this.clientData.createDuplicate());
			}
			if (this.fLocksText !== null) {
				copy.setFLocksText(this.fLocksText);
			}
			if (!oPr || false !== oPr.cacheImage) {
				copy.cachedImage = this.getBase64Img();
				copy.cachedPixH = this.cachedPixH;
				copy.cachedPixW = this.cachedPixW;
			}
			copy.setLocks(this.locks);
			return copy;
		};

		CImageShape.prototype.getImageUrl = function () {
			if (isRealObject(this.blipFill))
				return this.blipFill.RasterImageId;
			return null;
		};
		CImageShape.prototype.getTransparent = function () {
			if (isRealObject(this.blipFill)) {
				return this.blipFill.getTransparent();
			}
			return null;
		};
		CImageShape.prototype.setTransparent = function (v) {
			if(this.blipFill) {
				let oNewBlipFill = this.blipFill.createDuplicate();
				oNewBlipFill.setTransparent(v);
				this.setBlipFill(oNewBlipFill);
			}
		};

		CImageShape.prototype.getSnapArrays = function (snapX, snapY) {
			var transform = this.getTransformMatrix();
			snapX.push(transform.tx);
			snapX.push(transform.tx + this.extX * 0.5);
			snapX.push(transform.tx + this.extX);
			snapY.push(transform.ty);
			snapY.push(transform.ty + this.extY * 0.5);
			snapY.push(transform.ty + this.extY);
		};

		CImageShape.prototype.getWatermarkProps = function () {
			var oProps = new Asc.CAscWatermarkProperties();
			oProps.put_Type(Asc.c_oAscWatermarkType.Image);
			oProps.setXfrmRot(AscFormat.normalizeRotate(this.getXfrmRot() || 0));
			oProps.put_ImageUrl2(this.blipFill.RasterImageId);
			oProps.put_Scale(-1);
			oProps.put_ImageSize(this.getXfrmExtX() * 36000 + 0.5 >> 0, this.getXfrmExtY() * 36000 + 0.5 >> 0);
			var oApi;
			if (window["Asc"] && window["Asc"]["editor"]) {
				oApi = window["Asc"]["editor"];
			} else {
				oApi = editor;
			}
			if (oApi) {
				var oImgP = new Asc.asc_CImgProperty();
				oImgP.ImageUrl = this.blipFill.RasterImageId;
				var oSize = oImgP.asc_getOriginSize(oApi);

				if (oSize) {
					var dScale = (((this.extX / oSize.Width) * 100 + 0.5) >> 0) / 100;
					oProps.put_Scale(dScale);
					var dAspect = this.extX / this.extY;
					var dAspect2 = oSize.Width / oSize.Height;
					if (AscFormat.fApproxEqual(dAspect, dAspect2, 0.01)) {
						var oParaDrawing = AscFormat.getParaDrawing(this);
						if (oParaDrawing) {
							var oParentParagraph = oParaDrawing.Get_ParentParagraph();
							if (oParentParagraph) {
								var oSectPr = oParentParagraph.Get_SectPr();
								if (oSectPr) {
									var Width = oSectPr.GetContentFrameWidth();
									if (AscFormat.fApproxEqual(this.extX, Width, 1)) {
										oProps.put_Scale(-1);
									}
								}
							}
						}
					}
				}
			}
			return oProps;
		};

		CImageShape.prototype.getParentObjects = CShape.prototype.getParentObjects;

		CImageShape.prototype.hitInPath = CShape.prototype.hitInPath;

		CImageShape.prototype.hitInInnerArea = CShape.prototype.hitInInnerArea;

		CImageShape.prototype.getRotateAngle = CShape.prototype.getRotateAngle;

		CImageShape.prototype.changeSize = CShape.prototype.changeSize;

		CImageShape.prototype.canRotate = function () {
			if (this.isCrop) {
				return false;
			}
			if (this.cropObject) {
				return false;
			}
			return AscFormat.CGraphicObjectBase.prototype.canRotate.call(this);
		};


		CImageShape.prototype.createRotateTrack = function () {
			return new AscFormat.RotateTrackShapeImage(this);
		};

		CImageShape.prototype.createResizeTrack = function (cardDirection) {
			return new AscFormat.ResizeTrackShapeImage(this, cardDirection);
		};

		CImageShape.prototype.createMoveTrack = function () {
			return new AscFormat.MoveShapeImageTrack(this);
		};

		CImageShape.prototype.getInvertTransform = function () {
			if (this.recalcInfo.recalculateTransform) {
				this.recalculateTransform();
				this.recalcInfo.recalculateTransform = true;
			}
			return this.invertTransform;
		};

		CImageShape.prototype.hitInTextRect = function (x, y) {
			return false;
		};

		CImageShape.prototype.convertToWord = function (document) {
			this.setBDeleted(true);
			var oCopy = this.copy(undefined);
			oCopy.removePlaceholder();
			oCopy.setBDeleted(false);
			return oCopy;
		};

		CImageShape.prototype.convertToPPTX = function (drawingDocument, worksheet) {
			var ret = this.copy(undefined);
			ret.setWorksheet(worksheet);
			ret.setParent(null);
			ret.setBDeleted(false);
			return ret;
		};

		CImageShape.prototype.convertToPdf = function () {
			let copy = new AscPDF.CPdfImage();
			
			copy.setParent(null);
			copy.setBDeleted(false);

			if (this.nvPicPr) {
				copy.setNvPicPr(this.nvPicPr.createDuplicate());
			}
			if (this.spPr) {
				copy.setSpPr(this.spPr.createDuplicate());
				copy.spPr.setParent(copy);
			}
			if (this.blipFill) {
				copy.setBlipFill(this.blipFill.createDuplicate());
			}
			if (this.style) {
				copy.setStyle(this.style.createDuplicate());
			}
			if (this.macro !== null) {
				copy.setMacro(this.macro);
			}
			if (this.textLink !== null) {
				copy.setTextLink(this.textLink);
			}
			if (this.clientData) {
				copy.setClientData(this.clientData.createDuplicate());
			}
			if (this.fLocksText !== null) {
				copy.setFLocksText(this.fLocksText);
			}
			copy.setLocks(this.locks);
			return copy;
		};

		CImageShape.prototype.recalculateBrush = CShape.prototype.recalculateBrush;

		CImageShape.prototype.recalculatePen = function () {
			CShape.prototype.recalculatePen.call(this);
			if (this.pen) {
				// if(AscFormat.isRealNumber(this.pen.w)){
				//     this.pen.w *= 2;
				// }
				// if(!this.pen.Fill){
				//     this.pen = null;
				// }
			}
		};

		CImageShape.prototype.getCompiledLine = CShape.prototype.getCompiledLine;

		CImageShape.prototype.getCompiledFill = CShape.prototype.getCompiledFill;

		CImageShape.prototype.getCompiledTransparent = CShape.prototype.getCompiledTransparent;

		CImageShape.prototype.getAllRasterImages = function (images) {
			this.blipFill && typeof this.blipFill.RasterImageId === "string" && this.blipFill.RasterImageId.length > 0 && images.push(this.blipFill.RasterImageId);
		};

		CImageShape.prototype.getHierarchy = function () {

			let hierarchy = [];
			if (this.isPlaceholder()) {
				var ph_type = this.getPlaceholderType();
				var ph_index = this.getPlaceholderIndex();
				var b_is_single_body = this.getIsSingleBody();
				switch (this.parent.kind) {
					case AscFormat.TYPE_KIND.SLIDE: {
						hierarchy.push(this.parent.Layout.getMatchingShape(ph_type, ph_index, b_is_single_body));
						hierarchy.push(this.parent.Layout.Master.getMatchingShape(ph_type, ph_index, b_is_single_body));
						break;
					}

					case AscFormat.TYPE_KIND.LAYOUT: {
						hierarchy.push(this.parent.Master.getMatchingShape(ph_type, ph_index, b_is_single_body));
						break;
					}
				}
			}
			return hierarchy;
		};

		CImageShape.prototype.recalculateTransform = function () {
			this.cachedImage = null;
			if (!isRealObject(this.group)) {
				if (this.spPr.xfrm.isNotNull()) {
					var xfrm = this.spPr.xfrm;
					this.x = xfrm.offX;
					this.y = xfrm.offY;
					this.extX = xfrm.extX;
					this.extY = xfrm.extY;
					this.rot = AscFormat.isRealNumber(xfrm.rot) ? xfrm.rot : 0;
					this.flipH = xfrm.flipH === true;
					this.flipV = xfrm.flipV === true;
				} else {
					if (this.isPlaceholder()) {
						var hierarchy = this.getHierarchy();
						for (var i = 0; i < hierarchy.length; ++i) {
							var hierarchy_sp = hierarchy[i];
							if (isRealObject(hierarchy_sp) && hierarchy_sp.spPr.xfrm.isNotNull()) {
								var xfrm = hierarchy_sp.spPr.xfrm;
								this.x = xfrm.offX;
								this.y = xfrm.offY;
								this.extX = xfrm.extX;
								this.extY = xfrm.extY;
								this.rot = AscFormat.isRealNumber(xfrm.rot) ? xfrm.rot : 0;
								this.flipH = xfrm.flipH === true;
								this.flipV = xfrm.flipV === true;
								break;
							}
						}
						if (i === hierarchy.length) {
							this.x = 0;
							this.y = 0;
							this.extX = 5;
							this.extY = 5;
							this.rot = 0;
							this.flipH = false;
							this.flipV = false;
						}
					} else {
						this.x = 0;
						this.y = 0;
						this.extX = 5;
						this.extY = 5;
						this.rot = 0;
						this.flipH = false;
						this.flipV = false;
					}
				}
			} else {
				var xfrm;
				if (this.spPr.xfrm.isNotNull()) {
					xfrm = this.spPr.xfrm;
				} else {
					if (this.isPlaceholder()) {
						var hierarchy = this.getHierarchy();
						for (var i = 0; i < hierarchy.length; ++i) {
							var hierarchy_sp = hierarchy[i];
							if (isRealObject(hierarchy_sp) && hierarchy_sp.spPr.xfrm.isNotNull()) {
								xfrm = hierarchy_sp.spPr.xfrm;
								break;
							}
						}
						if (i === hierarchy.length) {
							xfrm = new AscFormat.CXfrm();
							xfrm.offX = 0;
							xfrm.offX = 0;
							xfrm.extX = 5;
							xfrm.extY = 5;
						}
					} else {
						xfrm = new AscFormat.CXfrm();
						xfrm.offX = 0;
						xfrm.offY = 0;
						xfrm.extX = 5;
						xfrm.extY = 5;
					}
				}
				var scale_scale_coefficients = this.group.getResultScaleCoefficients();
				this.x = scale_scale_coefficients.cx * (xfrm.offX - this.group.spPr.xfrm.chOffX);
				this.y = scale_scale_coefficients.cy * (xfrm.offY - this.group.spPr.xfrm.chOffY);
				this.extX = scale_scale_coefficients.cx * xfrm.extX;
				this.extY = scale_scale_coefficients.cy * xfrm.extY;
				this.rot = AscFormat.isRealNumber(xfrm.rot) ? xfrm.rot : 0;
				this.flipH = xfrm.flipH === true;
				this.flipV = xfrm.flipV === true;
			}
			this.transform.Reset();
			var hc = this.extX * 0.5;
			var vc = this.extY * 0.5;
			global_MatrixTransformer.TranslateAppend(this.transform, -hc, -vc);
			if (this.flipH)
				global_MatrixTransformer.ScaleAppend(this.transform, -1, 1);
			if (this.flipV)
				global_MatrixTransformer.ScaleAppend(this.transform, 1, -1);
			global_MatrixTransformer.RotateRadAppend(this.transform, -this.rot);
			global_MatrixTransformer.TranslateAppend(this.transform, this.x + hc, this.y + vc);
			if (isRealObject(this.group)) {
				global_MatrixTransformer.MultiplyAppend(this.transform, this.group.getTransformMatrix());
			}
			this.invertTransform = global_MatrixTransformer.Invert(this.transform);
			//if(this.drawingBase && !this.group)
			//{
			//    this.drawingBase.setGraphicObjectCoords();
			//}
		};

		CImageShape.prototype.Refresh_RecalcData = function (data) {
			switch (data.Type) {
				case AscDFH.historyitem_ImageShapeSetBlipFill: {
					this.recalcBrush();
					this.recalcFill();
					this.addToRecalculate();
					break;
				}
				case AscDFH.historyitem_ShapeSetBDeleted: {
					if (!this.bDeleted) {
						this.addToRecalculate();
					}
					break;
				}
			}
		};

		CImageShape.prototype.recalculateGeometry = function () {
			this.calcGeometry = null;
			if (isRealObject(this.spPr && this.spPr.geometry)) {
				this.calcGeometry = this.spPr.geometry;
			} else {
				var hierarchy = this.getHierarchy();
				for (var i = 0; i < hierarchy.length; ++i) {
					if (hierarchy[i] && hierarchy[i].spPr && hierarchy[i].spPr.geometry) {
						var _g = hierarchy[i].spPr.geometry;
						this.calcGeometry = AscFormat.ExecuteNoHistory(function () {
							var _r = _g.createDuplicate();
							_r.setParent(this);
							return _r;
						}, this, []);
						break;
					}
				}
			}
			if (isRealObject(this.calcGeometry)) {
				var transform = this.getTransform();
				this.calcGeometry.Recalculate(transform.extX, transform.extY);
			}
		};

		CImageShape.prototype.getTransformMatrix = function () {
			if (this.recalcInfo.recalculateTransform) {
				this.recalculateTransform();
				this.recalcInfo.recalculateTransform = false;
			}
			return this.transform;
		};

		CImageShape.prototype.getTransform = function () {
			if (this.recalcInfo.recalculateTransform) {
				this.recalculateTransform();
				this.recalcInfo.recalculateTransform = false;
			}
			return {
				x: this.x,
				y: this.y,
				extX: this.extX,
				extY: this.extY,
				rot: this.rot,
				flipH: this.flipH,
				flipV: this.flipV
			};
		};

		CImageShape.prototype.draw = function (graphics, transform) {
			if (this.checkNeedRecalculate && this.checkNeedRecalculate()) {
				return;
			}
			var oUR = graphics.updatedRect;
			if (oUR && this.bounds) {
				if (!oUR.isIntersectOther(this.bounds)) {
					return;
				}
			}
			if (graphics.animationDrawer) {
				graphics.animationDrawer.drawObject(this, graphics);
				return;
			}


			var oClipRect;
			if (!graphics.isBoundsChecker()) {
				oClipRect = this.getClipRect();
			}
			if (oClipRect) {
				graphics.SaveGrState();
				graphics.AddClipRect(oClipRect.x, oClipRect.y, oClipRect.w, oClipRect.h);
			}
			this.drawShdw && this.drawShdw(graphics);
			var _transform = transform ? transform : this.transform;
			graphics.SetIntegerGrid(false);
			graphics.transform3(_transform, false);
			var shape_drawer = new AscCommon.CShapeDrawer();
			if (this.getObjectType() !== AscDFH.historyitem_type_OleObject && (this.pen || this.brush)) {
				shape_drawer.fromShape2(this, graphics, this.calcGeometry);
				shape_drawer.draw(this.calcGeometry);
				shape_drawer.Clear();
			}
			var oldBrush = this.brush;
			var oldPen = this.pen;

			if (this.getObjectType() === AscDFH.historyitem_type_OleObject && !graphics.isBoundsChecker()) {
				var sImageId = this.blipFill && this.blipFill.RasterImageId;
				if (sImageId) {
					var oApi = editor || window['Asc']['editor'];
					if (oApi) {
						sImageId = AscCommon.getFullImageSrc2(sImageId);
						var _img = oApi.ImageLoader.map_image_index[sImageId];
						if ((_img && _img.Status === AscFonts.ImageLoadStatus.Loading) || (_img && _img.Image) || window["NATIVE_EDITOR_ENJINE"]) {
							this.brush = CreateBrushFromBlipFill(this.blipFill);
							this.pen = null;
						} else {
							this.brush = AscFormat.CreateNoFillUniFill();
						}
					}
				} else {
					this.brush = CreateBrushFromBlipFill(this.blipFill);
					this.pen = null;
				}
			} else {
				this.brush = CreateBrushFromBlipFill(this.blipFill);
				//this.pen = null;
			}

			shape_drawer.fromShape2(this, graphics, this.calcGeometry);
			shape_drawer.draw(this.calcGeometry);
			if (this.cropBrush) {
				this.brush = this.cropBrush;
				this.pen = null;
				shape_drawer.Clear();
				shape_drawer.fromShape2(this, graphics, this.calcGeometry);
				shape_drawer.draw(this.calcGeometry);
			}

			this.brush = oldBrush;
			this.pen = oldPen;

			this.drawLocks(_transform, graphics);
			if (oClipRect) {
				graphics.RestoreGrState();
			}
			graphics.reset();
			graphics.SetIntegerGrid(true);
		};


		CImageShape.prototype.handleUpdateLn = function () {
			this.recalcLine();
			this.recalcPen();
			this.addToRecalculate();
		};

		CImageShape.prototype.changePresetGeom = function (sPreset) {
			if (sPreset === "textRect") {
				return;
			}
			this.spPr.setGeometry(AscFormat.CreateGeometry(sPreset));
		};

		CImageShape.prototype.changeShadow = function (oShadow) {


			this.spPr && this.spPr.changeShadow(oShadow);
		};

		CImageShape.prototype.recalculateLocalTransform = CShape.prototype.recalculateLocalTransform;
		CImageShape.prototype.hit = CShape.prototype.hit;

		CImageShape.prototype.changeLine = function (line) {
			if (this.recalcInfo.recalculatePen) {
				this.recalculatePen();
			}
			var stroke = AscFormat.CorrectUniStroke(line, this.pen);
			if (stroke.Fill) {
				stroke.Fill.convertToPPTXMods();
			}
			this.spPr.setLn(stroke);
		};
		CImageShape.prototype.changeFill = function (unifill) {

			if (this.recalcInfo.recalculateBrush) {
				this.recalculateBrush();
			}
			var unifill2 = AscFormat.CorrectUniFill(unifill, this.brush, this.getEditorType());
			unifill2.convertToPPTXMods();
			this.setFill(unifill2);
		};
		CImageShape.prototype.setFill = function (fill) {

			this.spPr.setFill(fill);
		};

		CImageShape.prototype.hasCrop = function () {
			if(this.blipFill && this.blipFill.srcRect) {
				return true;
			}
			return false;
		};


		CImageShape.prototype.drawAdjustments = function (drawingDocument) {
			if (!this.canChangeAdjustments()) {
				return;
			}
			if (this.calcGeometry) {
				this.calcGeometry.drawAdjustments(drawingDocument, this.transform, false);
			}
		};

		CImageShape.prototype.hitToAdjustment = CShape.prototype.hitToAdjustment;

		CImageShape.prototype.getMediaFileName = function () {
			let oUniMedia = this.getUniMedia();
			if (oUniMedia) {
				if (oUniMedia.type === 7 || oUniMedia.type === 8) {
					if (typeof oUniMedia.media === "string" && oUniMedia.media.length > 0) {
						let sExt = AscCommon.GetFileExtension(oUniMedia.media);
						if (this.blipFill && typeof this.blipFill.RasterImageId === 'string') {
							let sName = AscCommon.GetFileName(this.blipFill.RasterImageId);
							let sMediaFile = sName + '.' + sExt;
							return sMediaFile;
						}
					}
				}
			}
			return null;
		};


		CImageShape.prototype.isVideo = function() {
			let oUniMedia = this.getUniMedia();
			if (oUniMedia && oUniMedia.type === 7) {
				return true;
			}
			return false;
		};

		CImageShape.prototype.getMediaData = function () {
			let oMediaData = new CMediaData(this);
			if (!oMediaData.isValid()) {
				return null;
			}
			return oMediaData;
		};

		CImageShape.prototype.setNvSpPr = function (pr) {
			AscCommon.History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_ImageShapeSetNvPicPr, this.nvPicPr, pr));
			this.nvPicPr = pr;
		};

		CImageShape.prototype.getAllImages = function (images) {
			if (this.blipFill instanceof AscFormat.CBlipFill && typeof this.blipFill.RasterImageId === "string") {
				images[AscCommon.getFullImageSrc2(this.blipFill.RasterImageId)] = true;
			}
		};
		CImageShape.prototype.checkTypeCorrect = function () {
			if (!this.blipFill) {
				return false;
			}
			if (!this.spPr) {
				return false;
			}
			return true;
		};

		CImageShape.prototype.getTypeName = function () {
			return AscCommon.translateManager.getValue("Picture");
		};

		CImageShape.prototype.getPictureBase64Data = function () {
			if (!this.blipFill) {
				return null;
			}
			//check tile
			if (this.blipFill.tile) {
				return null;
			}
			//check crop
			let oSrcRect = this.blipFill.srcRect;
			// if (oSrcRect) {
			// 	let fAE = AscFormat.fApproxEqual;
			// 	if (!fAE(oSrcRect.l, 0) || !fAE(oSrcRect.t, 0) ||
			// 				!fAE(oSrcRect.r, 100) || fAE(oSrcRect.b, 100)) {
			// 		return null;
			// 	}
			// }
			//check geometry
			if (this.calcGeometry && this.calcGeometry.preset !== "rect") {
				return null;
			}
			//check outline
			if (this.pen && this.pen.isVisible()) {
				return null;
			}

			//try draw source image
			let sRasterImageId = this.blipFill.RasterImageId;
			if (!sRasterImageId) {
				return null;
			}
			let oApi = Asc.editor || editor;
			if (!oApi) {
				return null;
			}
			let oBase64Data = this.blipFill.getBase64Data(false, false);
			let isNum = AscFormat.isRealNumber;
			if (oBase64Data.img && isNum(oBase64Data.w) && isNum(oBase64Data.h)) {
				return {ImageUrl: oBase64Data.img, width: oBase64Data.w, height: oBase64Data.h};
			}
			return null;
		};
		CImageShape.prototype.replacePictureData = function (sData, dW, dH, bWord, replaceMode) {
			let sOldRasterImageId = this.blipFill && this.blipFill.RasterImageId;
			this.setBlipFill(AscFormat.CreateBlipFillRasterImageId(sData));
			let oXfrm = this.spPr && this.spPr.xfrm;
			if (!oXfrm) {
				return;
			}
			if (!sOldRasterImageId) {
				return;
			}
			let oApi = Asc.editor || editor;
			if (!oApi) {
				return null;
			}
			let oImageLoader = oApi.ImageLoader;
			if (!oImageLoader) {
				return;
			}
			let oImage = oImageLoader.map_image_index[AscCommon.getFullImageSrc2(sOldRasterImageId)];
			if (!oImage || !oImage.Image || oImage.Status !== AscFonts.ImageLoadStatus.Complete) {
				return;
			}
			if(replaceMode === "stretch") {
				return;
			}
			else if(replaceMode === "fill") {
				this.cropFill();
				return;
			}
			else if(replaceMode === "fit") {
				this.cropFit();
				return;
			}
			else if(replaceMode === "original") {
				var oImgP = new Asc.asc_CImgProperty();
				oImgP.ImageUrl = this.blipFill.RasterImageId;
				var oSize = oImgP.asc_getOriginSize(oApi);
				if(oSize.IsCorrect) {
					let dX = this.getXfrmOffX();
					let dY = this.getXfrmOffY();
					let dExtX = oSize.Width;
					let dExtY = oSize.Height;
					let nRot = this.getXfrmRot();
					let bFlipH = this.getXfrmFlipH();
					let bFlipV = this.getXfrmFlipV();
					this.setTransformParams(dX, dY, dExtX, dExtY, nRot, bFlipH, bFlipV);
					this.checkDrawingBaseCoords();
					if (this.parent && this.parent.CheckWH) {
						this.parent.CheckWH();
					}
				}
				return;
			}
			let nPixWOld = Math.max(oImage.Image.width, 1);
			let nPixHOld = Math.max(oImage.Image.height, 1);

			let nPixW = dW / AscCommon.g_dKoef_pix_to_mm;
			let nPixH = dH / AscCommon.g_dKoef_pix_to_mm;

			let dWKoeff = nPixW / nPixWOld;
			let dHKoeff = nPixH / nPixHOld;
			if (!AscFormat.fApproxEqual(dWKoeff, 1.0, 0.01) ||
				!AscFormat.fApproxEqual(dHKoeff, 1.0, 0.01)) {

				let oMainGroup = this.getMainGroup();
				if (oMainGroup) {
					oMainGroup.normalize();
				}
				let dExtX = this.extX;
				let dExtY = this.extY;
				let dX = oXfrm.offX;
				let dY = oXfrm.offY;
				let bFlipH = oXfrm.flipH;
				let bFlipV = oXfrm.flipV;
				let nRot = oXfrm.rot;
				let dCenterX = dX + dExtX / 2.0;
				let dCenterY = dY + dExtY / 2.0;
				dExtX *= dWKoeff;
				dExtY *= dHKoeff;
				let bChangePosition = !bWord || !!this.group;
				if (bChangePosition) {
					dX = dCenterX - dExtX / 2.0;
					dY = dCenterY - dExtY / 2.0;
				}
				this.setTransformParams(dX, dY, dExtX, dExtY, nRot, bFlipH, bFlipV);
				this.checkDrawingBaseCoords();
				if (this.parent && this.parent.CheckWH) {
					this.parent.CheckWH();
				}
			}
		};
		CImageShape.prototype.pasteFormatting = function (oFormatData) {
			if (!oFormatData) {
				return;
			}
			this.pasteDrawingFormatting(oFormatData.Drawing);
		};
		CImageShape.prototype.compareForMorph = function(oDrawingToCheck, oCurCandidate, oMapPaired) {
			return AscFormat.CShape.prototype.compareForMorph.call(this, oDrawingToCheck, oCurCandidate, oMapPaired);
		};
		CImageShape.prototype.getText = function() {
			return null;
		};
		CImageShape.prototype.canFill = function () {
			return true;
		};

		function CreateBrushFromBlipFill(oBlipFill) {
			if (!oBlipFill) {
				return null;
			}
			var oBrush = new AscFormat.CUniFill();
			oBrush.fill = oBlipFill;

			let nTransparent = oBlipFill.getTransparent();
			if(AscFormat.isRealNumber(nTransparent)) {
				oBrush.setTransparent(nTransparent);
			}
			return oBrush;
		}

		function CMediaData(drawing) {
			this.drawing = drawing;
			this.additionalData = null;
		}
		CMediaData.prototype.setDrawing = function(oDrawing) {
			this.drawing = oDrawing;
		};
		CMediaData.prototype.setAdditionalData = function(oAdditionalData) {
			this.additionalData = oAdditionalData;
		};
		CMediaData.prototype.getName = function() {
			if(!this.drawing)
				return null;
			return this.drawing.getMediaFileName();
		};
		CMediaData.prototype.isValid = function() {
			if(!this.getName())
				return false;
			let oDrawing = this.getDrawing();
			if(!oDrawing || !oDrawing.IsUseInDocument())
				return false;
			return true;
		};
		CMediaData.prototype.getDrawing = function() {
			return this.drawing;
		};
		CMediaData.prototype.getTransform = function() {
			if(!this.drawing)
				return null;
			return this.drawing.getTransformMatrix();
		};
		CMediaData.prototype.isFullScreen = function() {
			if(!this.additionalData)
				return false;

			return this.additionalData.isFullScreen();
		};
		CMediaData.prototype.isMute = function() {
			if(!this.additionalData)
				return false;

			return this.additionalData.isMute();
		};
		CMediaData.prototype.isVideo = function() {
			if(!this.additionalData)
				return this.drawing.isVideo();

			return this.additionalData.isVideo();
		};
		CMediaData.prototype.getVol = function() {
			if(!this.additionalData)
				return null;

			return this.additionalData.getVol();
		};
		CMediaData.prototype.getStartTime = function() {
			if(!this.drawing)
				return null;
			return this.drawing.getMediaStartTime();
		};
		CMediaData.prototype.getEndTime = function() {
			if(!this.drawing)
				return null;
			return this.drawing.getMediaEndTime();
		};
		CMediaData.prototype.getPlayerData = function(oFrameRect, oControlRect) {
			if(!this.isValid())
				return null;
			let oDrawing = this.getDrawing();
			let oData = CreateMediaData();
			oData["DrawingId"] = oDrawing.Id;
			oData["FullScreen"] = this.isFullScreen();
			oData["Mute"] = this.isMute();
			oData["IsVideo"] = this.isVideo();
			oData["Volume"] = this.getVol();
			oData["StartTime"] = this.getStartTime();
			oData["EndTime"] = this.getEndTime();
			oData["FrameRect"] = oFrameRect.toObject();
			oData["ControlRect"] = oControlRect.toObject();
			oData["Rotation"] = oDrawing.getFullRotate();
			oData["FlipH"] = oDrawing.getFullFlipH();
			oData["FlipV"] = oDrawing.getFullFlipV();
			oData["IsSelected"] = oDrawing.selected;
			oData["MediaFile"] = this.getName();
			oData["From"] = 0;
			oData["Theme"] = AscCommon.GlobalSkin.Type;
			return oData;
		};

		function CreateMediaData() {
			let oData = {};
			oData["DrawingId"] = null;
			oData["FullScreen"] = null;
			oData["Mute"] = null;
			oData["IsVideo"] = null;
			oData["Volume"] = null;
			oData["StartTime"] = null;
			oData["EndTime"] = null;
			oData["FrameRect"] = null;
			oData["ControlRect"] = null;
			oData["DrawingTransform"] = null;
			oData["MediaFile"] = null;
			oData["From"] = null;
			return oData;
		}
		//--------------------------------------------------------export----------------------------------------------------
		window['AscFormat'] = window['AscFormat'] || {};
		window['AscFormat'].CImageShape = CImageShape;
		window['AscFormat'].CreateMediaData = CreateMediaData;
	})(window);
