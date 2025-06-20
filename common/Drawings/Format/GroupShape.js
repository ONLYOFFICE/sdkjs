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
		var c_oAscSizeRelFromH = AscCommon.c_oAscSizeRelFromH;
		var c_oAscSizeRelFromV = AscCommon.c_oAscSizeRelFromV;
		var isRealObject = AscCommon.isRealObject;
		var History = AscCommon.History;
		var global_MatrixTransformer = AscCommon.global_MatrixTransformer;

		var CShape = AscFormat.CShape;


		AscDFH.changesFactory[AscDFH.historyitem_GroupShapeSetNvGrpSpPr] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_GroupShapeSetSpPr] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_GroupShapeSetParent] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_GroupShapeAddToSpTree] = AscDFH.CChangesDrawingsContent;
		AscDFH.changesFactory[AscDFH.historyitem_GroupShapeSetGroup] = AscDFH.CChangesDrawingsObject;
		AscDFH.changesFactory[AscDFH.historyitem_GroupShapeRemoveFromSpTree] = AscDFH.CChangesDrawingsContent;

		AscDFH.drawingsChangesMap[AscDFH.historyitem_GroupShapeSetNvGrpSpPr] = function (oClass, value) {
			oClass.nvGrpSpPr = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_GroupShapeSetSpPr] = function (oClass, value) {
			oClass.spPr = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_GroupShapeSetParent] = function (oClass, value) {
			oClass.oldParent = oClass.parent;
			oClass.parent = value;
		};
		AscDFH.drawingsChangesMap[AscDFH.historyitem_GroupShapeSetGroup] = function (oClass, value) {
			oClass.group = value;
		};

		AscDFH.drawingContentChanges[AscDFH.historyitem_GroupShapeAddToSpTree] = AscDFH.drawingContentChanges[AscDFH.historyitem_GroupShapeRemoveFromSpTree] = function (oClass) {
			return oClass.spTree;
		};

		/**
		 * @extends CGraphicObjectBase
		 * @constructor
		 */
		function CGroupShape() {
			AscFormat.CGraphicObjectBase.call(this);
			this.nvGrpSpPr = null;
			this.spTree = [];


			this.invertTransform = null;
			this.scaleCoefficients = {cx: 1, cy: 1};

			this.arrGraphicObjects = [];
			this.selectedObjects = [];


			this.selection =
				{
					groupSelection: null,
					chartSelection: null,
					textSelection: null
				};
		}

		AscFormat.InitClass(CGroupShape, AscFormat.CGraphicObjectBase, AscDFH.historyitem_type_GroupShape);

		CGroupShape.prototype.GetAllDrawingObjects = function (DrawingObjects) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].GetAllDrawingObjects) {
					this.spTree[i].GetAllDrawingObjects(DrawingObjects);
				}
			}
		};
		CGroupShape.prototype.getRelativePosition = function (obj) {
			let result = obj || {x: 0, y: 0};
			result.x += this.x;
			result.y += this.y;
			if (this.group) {
				this.group.getRelativePosition(result);
			}
			return result;
		};
		CGroupShape.prototype.hasSmartArt = function (bRetSmartArt) {
			let hasSmartArt = false;
			for (let i = 0; i < this.spTree.length; i += 1) {
				if (hasSmartArt) {
					return hasSmartArt;
				}
				if (this.spTree[i].hasSmartArt) {
					hasSmartArt = this.spTree[i].hasSmartArt(bRetSmartArt);
				}
			}
			return hasSmartArt;
		};

		CGroupShape.prototype.recalcTransformText = function () {
			this.spTree.forEach(function (oDrawing) {
				if (oDrawing.recalcTransformText) {
					oDrawing.recalcTransformText();
				}
			});
		};

		CGroupShape.prototype.documentGetAllFontNames = function (allFonts) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].documentGetAllFontNames)
					this.spTree[i].documentGetAllFontNames(allFonts);
			}
		};
		CGroupShape.prototype.getImageFromBulletsMap = function (oImages) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].getImageFromBulletsMap)
					this.spTree[i].getImageFromBulletsMap(oImages);
			}
		};
		CGroupShape.prototype.getDocContentsWithImageBullets = function (arrContents) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].getDocContentsWithImageBullets)
					this.spTree[i].getDocContentsWithImageBullets(arrContents);
			}
		};
		CGroupShape.prototype.handleAllContents = function (fCallback) {
			for (var i = 0; i < this.spTree.length; ++i) {
				this.spTree[i].handleAllContents(fCallback);
			}
		};
		CGroupShape.prototype.getAllDocContents = function (aDocContents) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].getAllDocContents)
					this.spTree[i].getAllDocContents(aDocContents);
			}
		};
		CGroupShape.prototype.checkRunContent = function (fCallback) {
			let aGraphics = this.getArrGraphicObjects();
			for (let nIdx = 0; nIdx < aGraphics.length; ++nIdx) {
				aGraphics[nIdx].checkRunContent(fCallback);
			}
		};

		CGroupShape.prototype.documentCreateFontMap = function (allFonts) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].documentCreateFontMap)
					this.spTree[i].documentCreateFontMap(allFonts);
			}
		};

		CGroupShape.prototype.setBDeleted2 = function (pr) {
			this.setBDeleted(pr);
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].setBDeleted2) {
					this.spTree[i].setBDeleted2(pr);
				} else {
					this.spTree[i].setBDeleted(pr);
				}
			}
		};

		CGroupShape.prototype.documentUpdateSelectionState = function () {
			if (this.selection.textSelection) {
				this.selection.textSelection.updateSelectionState();
			} else if (this.selection.groupSelection) {
				this.selection.groupSelection.documentUpdateSelectionState();
			} else if (this.selection.chartSelection) {
				this.selection.chartSelection.documentUpdateSelectionState();
			} else {
				this.getDrawingDocument().SelectClear();
				this.getDrawingDocument().TargetEnd();
			}
		};

		CGroupShape.prototype.drawSelectionPage = function (pageIndex) {
			var oMatrix = null;
			if (this.selection.textSelection) {
				if (this.selection.textSelection.transformText) {
					oMatrix = this.selection.textSelection.transformText.CreateDublicate();
				}
				this.getDrawingDocument().UpdateTargetTransform(oMatrix);
				this.selection.textSelection.getDocContent().DrawSelectionOnPage(pageIndex);
			} else if (this.selection.chartSelection && this.selection.chartSelection.selection.textSelection) {
				if (this.selection.chartSelection.selection.textSelection.transformText) {
					oMatrix = this.selection.chartSelection.selection.textSelection.transformText.CreateDublicate();
				}
				this.getDrawingDocument().UpdateTargetTransform(oMatrix);
				this.selection.chartSelection.selection.textSelection.getDocContent().DrawSelectionOnPage(pageIndex);
			}
		};

		CGroupShape.prototype.setNvGrpSpPr = function (pr) {
			History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_GroupShapeSetNvGrpSpPr, this.nvGrpSpPr, pr));
			this.nvGrpSpPr = pr;
		};

		CGroupShape.prototype.setSpPr = function (pr) {
			History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_GroupShapeSetSpPr, this.spPr, pr));
			this.spPr = pr;
			if (pr) {
				pr.setParent(this);
			}
		};
		CGroupShape.prototype.getSpCount = function () {
			return this.spTree.length;
		};
		CGroupShape.prototype.addToSpTree = function (pos, item) {
			if (!AscFormat.isRealNumber(pos))
				pos = this.spTree.length;
			History.Add(new AscDFH.CChangesDrawingsContent(this, AscDFH.historyitem_GroupShapeAddToSpTree, pos, [item], true));
			this.handleUpdateSpTree();
			if (item.group !== this) {
				item.setGroup(this);
			}
			this.spTree.splice(pos, 0, item);
		};

		CGroupShape.prototype.setParent = function (pr) {
			History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_GroupShapeSetParent, this.parent, pr));
			this.parent = pr;
		};

		CGroupShape.prototype.setGroup = function (group) {
			History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_GroupShapeSetGroup, this.group, group));
			this.group = group;
		};

		CGroupShape.prototype.getPosInSpTree = function (id) {
			for (var i = this.spTree.length - 1; i > -1; --i) {
				if (this.spTree[i].Get_Id() === id) {
					return i;
				}
			}
			return null;
		};

		CGroupShape.prototype.removeFromSpTree = function (id) {
			var nPos = this.getPosInSpTree(id);
			if (nPos !== null) {
				return this.removeFromSpTreeByPos(nPos);
			}
			return null;
		};

		CGroupShape.prototype.removeFromSpTreeByPos = function (pos) {
			var aSplicedShape = this.spTree.splice(pos, 1);
			History.Add(new AscDFH.CChangesDrawingsContent(this, AscDFH.historyitem_GroupShapeRemoveFromSpTree, pos, aSplicedShape, false));
			this.handleUpdateSpTree();
			return aSplicedShape[0];
		};

		CGroupShape.prototype.handleUpdateSpTree = function () {
			if (!this.group) {
				this.recalcInfo.recalculateArrGraphicObjects = true;
				this.recalcBounds();
				this.addToRecalculate();
			} else {
				this.recalcInfo.recalculateArrGraphicObjects = true;
				this.group.handleUpdateSpTree();
				this.recalcBounds();
			}
		};

		CGroupShape.prototype.handleUpdateExtents = function() {
			this.recalcTransform();
			for(let nSp = 0; nSp < this.spTree.length; ++nSp) {
				this.spTree[nSp].handleUpdateExtents();
			}
			if(!this.group) {
				if(this.addToRecalculate) {
					this.addToRecalculate();
				}
			}
		};

		CGroupShape.prototype.copy = function (oPr) {
			var copy = new CGroupShape();
			this.copy2(copy, oPr);
			return copy;
		};
		CGroupShape.prototype.copy2 = function (copy, oPr) {
			if (this.nvGrpSpPr) {
				copy.setNvGrpSpPr(this.nvGrpSpPr.createDuplicate());
			}
			if (this.spPr) {
				copy.setSpPr(this.spPr.createDuplicate());
				copy.spPr.setParent(copy);
			}
			for (var i = 0; i < this.spTree.length; ++i) {
				var _copy;
				if (this.spTree[i].isGroupObject()) {
					_copy = this.spTree[i].copy(oPr);
				} else {
					if (oPr && oPr.bSaveSourceFormatting) {
						_copy = this.spTree[i].getCopyWithSourceFormatting();
					} else {
						_copy = this.spTree[i].copy(oPr);
					}

				}
				if (oPr && AscCommon.isRealObject(oPr.idMap)) {
					oPr.idMap[this.spTree[i].Id] = _copy.Id;
				}
				copy.addToSpTree(copy.spTree.length, _copy);
				copy.spTree[copy.spTree.length - 1].setGroup(copy);
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

		CGroupShape.prototype.getAllImages = function (images) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (typeof this.spTree[i].getAllImages === "function") {
					this.spTree[i].getAllImages(images);
				}
			}
		};
		CGroupShape.prototype.clearLang = function () {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (typeof this.spTree[i].clearLang === "function") {
					this.spTree[i].clearLang();
				}
			}
		};

		CGroupShape.prototype.convertToWord = function (document) {
			this.setBDeleted(true);
			var c = new CGroupShape();
			c.setBDeleted(false);
			if (this.nvGrpSpPr) {
				c.setNvSpPr(this.nvGrpSpPr.createDuplicate());
			}
			if (this.spPr) {
				c.setSpPr(this.spPr.createDuplicate());
				c.spPr.setParent(c);
			}

			for (var i = 0; i < this.spTree.length; ++i) {
				c.addToSpTree(c.spTree.length, this.spTree[i].convertToWord(document));
				c.spTree[c.spTree.length - 1].setGroup(c);
			}
			c.removePlaceholder();
			return c;
		};

		CGroupShape.prototype.convertToPPTX = function (drawingDocument, worksheet) {
			var c = new CGroupShape();
			c.setBDeleted(false);
			c.setWorksheet(worksheet);
			if (this.nvGrpSpPr) {
				c.setNvSpPr(this.nvGrpSpPr.createDuplicate());
			}
			if (this.spPr) {
				c.setSpPr(this.spPr.createDuplicate());
				c.spPr.setParent(c);
			}

			for (var i = 0; i < this.spTree.length; ++i) {
				c.addToSpTree(c.spTree.length, this.spTree[i].convertToPPTX(drawingDocument, worksheet));
				c.spTree[c.spTree.length - 1].setGroup(c);
			}
			return c;
		};

		CGroupShape.prototype.getAllFonts = function (fonts) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (typeof this.spTree[i].getAllFonts === "function")
					this.spTree[i].getAllFonts(fonts);
			}
		};


		CGroupShape.prototype.isPlaceholder = function () {
			return !!(this.nvGrpSpPr != null && this.nvGrpSpPr.nvPr && this.nvGrpSpPr.nvPr.ph);
		};

		CGroupShape.prototype.getAllRasterImages = function (images) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].getAllRasterImages)
					this.spTree[i].getAllRasterImages(images);
			}
		};

		CGroupShape.prototype.getAllSlicerViews = function (aSlicerView) {
			for (var nSp = 0; nSp < this.spTree.length; ++nSp) {
				this.spTree[nSp].getAllSlicerViews(aSlicerView);
			}
		};

		CGroupShape.prototype.hit = function (x, y) {
			for (var i = this.spTree.length - 1; i > -1; --i) {
				if (this.spTree[i].hit(x, y)) {
					return true;
				}
			}
			return false;
		};

		CGroupShape.prototype.onMouseMove = function (e, x, y) {
			for (var i = this.spTree.length - 1; i > -1; --i) {
				if (this.spTree[i].onMouseMove(e, x, y)) {
					return true;
				}
			}
			return false;
		};

		/**
		 * @memberOf CGroupShape
		 */
		CGroupShape.prototype.draw = function (graphics) {
			if (this.checkNeedRecalculate && this.checkNeedRecalculate()) {
				return;
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
			for (var i = 0; i < this.spTree.length; ++i)
				this.spTree[i].draw(graphics);


			this.drawLocks(this.transform, graphics);
			if (oClipRect) {
				graphics.RestoreGrState();
			}
			graphics.reset();
			graphics.SetIntegerGrid(true);
		};

		CGroupShape.prototype.deselectObject = function (object) {
			for (var i = 0; i < this.selectedObjects.length; ++i) {
				if (this.selectedObjects[i] === object) {
					object.selected = false;
					this.selectedObjects.splice(i, 1);
					return;
				}
			}
		};
		CGroupShape.prototype.resetInternalSelectionObject = function (oSelectedObject) {
			if (this.selection.textSelection === oSelectedObject) {
				if (this.selection.textSelection.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
					if (this.selection.textSelection.graphicObject) {
						this.selection.textSelection.graphicObject.RemoveSelection();
					}
				} else {
					const oContent = this.selection.textSelection.getDocContent();
					oContent && oContent.RemoveSelection();
				}
			}
			if (this.selection.chartSelection === oSelectedObject) {
				this.selection.chartSelection.resetSelection();
				this.selection.chartSelection = null;
			}
		};
		CGroupShape.prototype.deselectInternal = function(oController) {
			const oMainGroup = this.getMainGroup();
			if (oMainGroup.selectedObjects.length) {
				const oSelectedObjects = {};
				for (let i = oMainGroup.selectedObjects.length - 1; i >= 0; i -= 1) {
					const oSelectedObject = oMainGroup.selectedObjects[i];
					if (oSelectedObject.bDeleted) {
						oMainGroup.deselectObject(oSelectedObject);
						oMainGroup.resetInternalSelectionObject(oSelectedObject);
					} else {
						oSelectedObjects[oSelectedObject.Get_Id()] = oSelectedObject;
					}
				}

				const arrGroups = [this];
				while (arrGroups.length) {
					const oGroup = arrGroups.pop();
					for (let i = 0; i < oGroup.spTree.length; i += 1) {
						const oShape = oGroup.spTree[i];
						if (oShape.isGroup()) {
							arrGroups.push(oShape);
						} else if (oSelectedObjects[oShape.Get_Id()]) {
							oMainGroup.deselectObject(oShape);
							oMainGroup.resetInternalSelectionObject(oShape);
						}
					}
				}
				if (!oMainGroup.selectedObjects.length) {
					oController.selection.groupSelection = null;
				}
			}
		};


		CGroupShape.prototype.getLocalTransform = function () {
			if (this.recalcInfo.recalculateTransform) {
				this.recalculateTransform();
				this.recalcInfo.recalculateTransform = false;
			}
			return this.localTransform;
		};

		CGroupShape.prototype.getArrGraphicObjects = function () {
			if (this.recalcInfo.recalculateArrGraphicObjects)
				this.recalculateArrGraphicObjects();
			return this.arrGraphicObjects;
		};

		CGroupShape.prototype.getInvertTransform = function () {
			return this.invertTransform;
		};

		CGroupShape.prototype.getResultScaleCoefficients = function () {
			if (this.recalcInfo.recalculateScaleCoefficients) {
				var cx, cy;
				if (this.spPr && this.spPr.xfrm && this.spPr.xfrm.isNotNullForGroup()) {

					var dExtX = this.spPr.xfrm.extX, dExtY = this.spPr.xfrm.extY;


					if (this.drawingBase && !this.group) {
						var metrics = this.drawingBase.getGraphicObjectMetrics();
						var rot = 0;
						if (this.spPr && this.spPr.xfrm) {
							if (AscFormat.isRealNumber(this.spPr.xfrm.rot)) {
								rot = AscFormat.normalizeRotate(this.spPr.xfrm.rot);
							}
						}

						var metricExtX, metricExtY;
						//  if(!(this instanceof AscFormat.CGroupShape))
						{
							if (AscFormat.checkNormalRotate(rot)) {
								dExtX = metrics.extX;
								dExtY = metrics.extY;
							} else {
								dExtX = metrics.extY;
								dExtY = metrics.extX;
							}
						}
					}

					if (this.spPr.xfrm.chExtX > 0)
						cx = dExtX / this.spPr.xfrm.chExtX;
					else
						cx = 1;

					if (this.spPr.xfrm.chExtY > 0)
						cy = dExtY / this.spPr.xfrm.chExtY;
					else
						cy = 1;
				} else {
					cx = 1;
					cy = 1;
				}
				if (isRealObject(this.group)) {
					var group_scale_coefficients = this.group.getResultScaleCoefficients();
					cx *= group_scale_coefficients.cx;
					cy *= group_scale_coefficients.cy;
				} else {
					let oParaDrawing = AscFormat.getParaDrawing(this);
					if (oParaDrawing) {
						let dScaleCoefficient = oParaDrawing.GetScaleCoefficient();
						cx *= dScaleCoefficient;
						cy *= dScaleCoefficient;
					}
				}

				this.scaleCoefficients.cx = cx;
				this.scaleCoefficients.cy = cy;
				this.recalcInfo.recalculateScaleCoefficients = false;
			}
			return this.scaleCoefficients;
		};

		CGroupShape.prototype.getCompiledTransparent = function () {
			return null;
		};

		CGroupShape.prototype.selectObject = function (object, pageIndex) {
			object.select(this, pageIndex);
		};
		CGroupShape.prototype.onChangeDrawingsSelection = function () {
		};

		CGroupShape.prototype.recalculate = function () {
			var recalcInfo = this.recalcInfo;

			if (recalcInfo.recalculateBrush) {
				this.recalculateBrush();
				recalcInfo.recalculateBrush = false;
			}
			if (recalcInfo.recalculatePen) {
				this.recalculatePen();
				recalcInfo.recalculatePen = false;
			}
			if (recalcInfo.recalculateScaleCoefficients) {
				this.getResultScaleCoefficients();
				recalcInfo.recalculateScaleCoefficients = false;
			}

			if (recalcInfo.recalculateTransform) {
				this.recalculateTransform();
				recalcInfo.recalculateTransform = false;
			}
			if (recalcInfo.recalculateArrGraphicObjects)
				this.recalculateArrGraphicObjects();
			for (var i = 0; i < this.spTree.length; ++i)
				this.spTree[i].recalculate();
			if (recalcInfo.recalculateBounds) {
				this.recalculateBounds();
				recalcInfo.recalculateBounds = false;
			}
		};

		CGroupShape.prototype.recalcTransform = function () {
			this.recalcInfo.recalculateTransform = true;
			this.recalcInfo.recalculateScaleCoefficients = true;
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].recalcTransform)
					this.spTree[i].recalcTransform();
				else {
					this.spTree[i].recalcInfo.recalculateTransform = true;
					this.spTree[i].recalcInfo.recalculateTransformText = true;
				}
			}
		};

		CGroupShape.prototype.canRotate = function () {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (!this.spTree[i].canRotate())
					return false;
			}
			return AscFormat.CGraphicObjectBase.prototype.canRotate.call(this);
		};

		CGroupShape.prototype.canResize = function () {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (!this.spTree[i].canResize())
					return false
			}
			return AscFormat.CGraphicObjectBase.prototype.canResize.call(this);
		};

		CGroupShape.prototype.drawAdjustments = function () {
		};

		CGroupShape.prototype.recalculateBrush = function () {
		};

		CGroupShape.prototype.recalculatePen = function () {
		};

		CGroupShape.prototype.recalculateArrGraphicObjects = function () {
			this.arrGraphicObjects.length = 0;
			for (let nSp = 0; nSp < this.spTree.length; ++nSp) {
				let oSp = this.spTree[nSp];
				if (oSp.isGroup() || oSp.isSmartArtObject()) {
					let aGraphicObjets = oSp.getArrGraphicObjects();
					for (let nGr = 0; nGr < aGraphicObjets.length; ++nGr) {
						this.arrGraphicObjects.push(aGraphicObjets[nGr]);
					}
				} else {
					this.arrGraphicObjects.push(oSp);
				}
			}
		};

		CGroupShape.prototype.paragraphAdd = function (paraItem, bRecalculate) {
			if (this.selection.textSelection) {
				this.selection.textSelection.paragraphAdd(paraItem, bRecalculate);
			} else if (this.selection.chartSelection) {
				this.selection.chartSelection.paragraphAdd(paraItem, bRecalculate);
			} else {
				var i;
				if (paraItem.Type === para_TextPr) {
					AscFormat.DrawingObjectsController.prototype.applyDocContentFunction.call(this, CDocumentContent.prototype.AddToParagraph, [paraItem, bRecalculate], CTable.prototype.AddToParagraph);
				} else if (this.selectedObjects.length === 1
					&& this.selectedObjects[0].getObjectType() === AscDFH.historyitem_type_Shape
					&& this.selectedObjects[0].canEditText()) {
					this.selection.textSelection = this.selectedObjects[0];
					this.selection.textSelection.paragraphAdd(paraItem, bRecalculate);
					if (AscFormat.isRealNumber(this.selection.textSelection.selectStartPage))
						this.selection.textSelection.select(this, this.selection.textSelection.selectStartPage);
				} else if (this.selectedObjects.length > 0) {
					if (this.parent) {
						this.parent.GoToText();
						this.resetSelection();
					}
				}
			}
		};

		CGroupShape.prototype.applyTextFunction = AscFormat.DrawingObjectsController.prototype.applyTextFunction;

		CGroupShape.prototype.applyDocContentFunction = AscFormat.DrawingObjectsController.prototype.applyDocContentFunction;

		CGroupShape.prototype.applyAllAlign = function (val) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (typeof this.spTree[i].applyAllAlign === "function") {
					this.spTree[i].applyAllAlign(val);
				}
			}
		};

		CGroupShape.prototype.applyAllSpacing = function (val) {

			for (var i = 0; i < this.spTree.length; ++i) {
				if (typeof this.spTree[i].applyAllSpacing === "function") {
					this.spTree[i].applyAllSpacing(val);
				}
			}
		};

		CGroupShape.prototype.applyAllNumbering = function (val) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (typeof this.spTree[i].applyAllNumbering === "function") {
					this.spTree[i].applyAllNumbering(val);
				}
			}
		};

		CGroupShape.prototype.applyAllIndent = function (val) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (typeof this.spTree[i].applyAllIndent === "function") {
					this.spTree[i].applyAllIndent(val);
				}
			}
		};

		CGroupShape.prototype.checkExtentsByDocContent = function (bForce, bNeedRecalc) {
			var bRet = false;
			for (var i = 0; i < this.spTree.length; ++i) {
				if (typeof this.spTree[i].checkExtentsByDocContent === "function") {
					if (this.spTree[i].checkExtentsByDocContent(bForce, bNeedRecalc)) {
						bRet = true;
					}
				}
			}
			return bRet;
		};

		CGroupShape.prototype.changeSize = function (kw, kh) {
			if (this.spPr && this.spPr.xfrm && this.spPr.xfrm.isNotNullForGroup()) {
				var xfrm = this.spPr.xfrm;
				xfrm.setOffX(xfrm.offX * kw);
				xfrm.setOffY(xfrm.offY * kh);
				xfrm.setExtX(xfrm.extX * kw);
				xfrm.setExtY(xfrm.extY * kh);
				xfrm.setChExtX(xfrm.chExtX * kw);
				xfrm.setChExtY(xfrm.chExtY * kh);
				xfrm.setChOffX(xfrm.chOffX * kw);
				xfrm.setChOffY(xfrm.chOffY * kh);
			}
			for (var i = 0; i < this.spTree.length; ++i) {
				this.spTree[i].changeSize(kw, kh);
			}
		};

		CGroupShape.prototype.recalculateTransform = function () {
			this.cachedImage = null;
			var xfrm;
			if (this.spPr.xfrm.isNotNullForGroup())
				xfrm = this.spPr.xfrm;
			else {
				xfrm = new AscFormat.CXfrm();
				xfrm.offX = 0;
				xfrm.offY = 0;
				xfrm.extX = 5;
				xfrm.extY = 5;
				xfrm.chOffX = 0;
				xfrm.chOffY = 0;
				xfrm.chExtX = 5;
				xfrm.chExtY = 5;
			}

			if (!isRealObject(this.group)) {
				this.x = xfrm.offX;
				this.y = xfrm.offY;
				this.extX = xfrm.extX;
				this.extY = xfrm.extY;
				this.rot = AscFormat.isRealNumber(xfrm.rot) ? xfrm.rot : 0;
				this.flipH = this.flipH === true;
				this.flipV = this.flipV === true;
			} else {
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

		CGroupShape.prototype.getTransformMatrix = function () {
			if (this.recalcInfo.recalculateTransform) {
				this.recalculateTransform();
			}
			return this.transform;
		};

		CGroupShape.prototype.getSnapArrays = function (snapX, snapY) {
			var transform = this.getTransformMatrix();
			snapX.push(transform.tx);
			snapX.push(transform.tx + this.extX * 0.5);
			snapX.push(transform.tx + this.extX);
			snapY.push(transform.ty);
			snapY.push(transform.ty + this.extY * 0.5);
			snapY.push(transform.ty + this.extY);
			for (var i = 0; i < this.arrGraphicObjects.length; ++i) {
				if (this.arrGraphicObjects[i].getSnapArrays) {
					this.arrGraphicObjects[i].getSnapArrays(snapX, snapY);
				}
			}
		};
		CGroupShape.prototype.getSelectedArray = function () {
			return this.selectedObjects;
		};
		CGroupShape.prototype.getSelectionState = function () {
			var selection_state = {};
			if (this.selection.textSelection) {
				selection_state.textObject = this.selection.textSelection;
				selection_state.selectStartPage = this.selection.textSelection.selectStartPage;
				selection_state.textSelection = this.selection.textSelection.getDocContent().GetSelectionState();
			} else if (this.selection.chartSelection) {
				selection_state.chartObject = this.selection.chartSelection;
				selection_state.selectStartPage = this.selection.chartSelection.selectStartPage;
				selection_state.chartSelection = this.selection.chartSelection.getSelectionState();
			} else {
				selection_state.selection = [];
				for (var i = 0; i < this.selectedObjects.length; ++i) {
					selection_state.selection.push({
						object: this.selectedObjects[i],
						pageIndex: this.selectedObjects[i].selectStartPage
					});
				}
			}
			return selection_state;
		};

		CGroupShape.prototype.setSelectionState = function (selection_state) {
			this.resetSelection(this);
			if (selection_state.textObject) {
				this.selectObject(selection_state.textObject, selection_state.selectStartPage);
				this.selection.textSelection = selection_state.textObject;
				selection_state.textObject.getDocContent().SetSelectionState(selection_state.textSelection, selection_state.textSelection.length - 1);
			} else if (selection_state.chartSelection) {
				this.selectObject(selection_state.chartObject, selection_state.selectStartPage);
				this.selection.chartSelection = selection_state.chartObject;
				selection_state.chartObject.setSelectionState(selection_state.chartSelection);
			} else {
				for (var i = 0; i < selection_state.selection.length; ++i) {
					this.selectObject(selection_state.selection[i].object, selection_state.selection[i].pageIndex);
				}
			}
		};

		CGroupShape.prototype.documentUpdateRulersState = function () {
			if (this.selectedObjects.length === 1 && this.selectedObjects[0].documentUpdateRulersState)
				this.selectedObjects[0].documentUpdateRulersState();
		};

		CGroupShape.prototype.CheckNeedRecalcAutoFit = function (oSectPr) {
			var bRet = false;
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].CheckNeedRecalcAutoFit && this.spTree[i].CheckNeedRecalcAutoFit(oSectPr)) {
					bRet = true;
				}
			}
			if (bRet) {
				this.recalcWrapPolygon();
			}
			return bRet;
		};

		CGroupShape.prototype.CheckGroupSizes = function () {
			var bRet = false;
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].checkExtentsByDocContent) {
					if (this.spTree[i].checkExtentsByDocContent(undefined, false)) {
						bRet = true;
					}
				}
			}
			if (bRet) {
				if (!this.group) {
					this.updateCoordinatesAfterInternalResize();
				}
				if (this.parent instanceof AscCommonWord.ParaDrawing) {
					this.parent.CheckWH();
				}
			}
			return bRet;
		};

		CGroupShape.prototype.GetRevisionsChangeElement = function (SearchEngine) {
			var i;
			if (this.selectedObjects.length === 0) {
				if (SearchEngine.GetDirection() > 0) {
					i = 0;
				} else {
					i = this.arrGraphicObjects.length - 1;
				}
			} else {
				if (SearchEngine.GetDirection() > 0) {
					for (i = 0; i < this.arrGraphicObjects.length; ++i) {
						if (this.arrGraphicObjects[i].selected) {
							break;
						}
					}
					if (i === this.arrGraphicObjects.length) {
						return;
					}
				} else {
					for (i = this.arrGraphicObjects.length - 1; i > -1; --i) {
						if (this.arrGraphicObjects[i].selected) {
							break;
						}
					}
					if (i === -1) {
						return;
					}
				}
			}
			while (!SearchEngine.IsFound()) {
				if (this.arrGraphicObjects[i].GetRevisionsChangeElement) {
					this.arrGraphicObjects[i].GetRevisionsChangeElement(SearchEngine);
				}
				if (SearchEngine.GetDirection() > 0) {
					if (i === this.arrGraphicObjects.length - 1) {
						break;
					}
					++i;
				} else {
					if (i === 0) {
						break;
					}
					--i;
				}
			}
		};

		CGroupShape.prototype.Search = function (SearchEngine, Type) {
			var Len = this.arrGraphicObjects.length;
			for (var i = 0; i < Len; ++i) {
				if (this.arrGraphicObjects[i].Search) {
					this.arrGraphicObjects[i].Search(SearchEngine, Type);
				}
			}
		};

		CGroupShape.prototype.GetSearchElementId = function (bNext, bCurrent) {
			var Current = -1;
			var Len = this.arrGraphicObjects.length;

			var Id = null;
			if (true === bCurrent) {
				for (var i = 0; i < Len; ++i) {
					if (this.arrGraphicObjects[i] === this.selection.textSelection) {
						Current = i;
						break;
					}
				}
			}

			if (true === bNext) {
				var Start = (-1 !== Current ? Current : 0);

				for (var i = Start; i < Len; i++) {
					if (this.arrGraphicObjects[i].GetSearchElementId) {
						Id = this.arrGraphicObjects[i].GetSearchElementId(true, i === Current ? true : false);
						if (null !== Id)
							return Id;
					}
				}
			} else {
				var Start = (-1 !== Current ? Current : Len - 1);

				for (var i = Start; i >= 0; i--) {
					if (this.arrGraphicObjects[i].GetSearchElementId) {
						Id = this.arrGraphicObjects[i].GetSearchElementId(false, i === Current ? true : false);
						if (null !== Id)
							return Id;
					}
				}
			}

			return null;
		};

		CGroupShape.prototype.FindNextFillingForm = function (isNext, isCurrent) {
			if (this.graphicObject)
				return this.graphicObject.FindNextFillingForm(isNext, isCurrent);

			var Current = -1;
			var Len = this.arrGraphicObjects.length;

			var Id = null;
			if (true === isCurrent) {
				for (var i = 0; i < Len; ++i) {
					if (this.arrGraphicObjects[i] === this.selection.textSelection) {
						Current = i;
						break;
					}
				}
			}

			if (true === isNext) {
				var Start = (-1 !== Current ? Current : 0);

				for (var i = Start; i < Len; i++) {
					if (this.arrGraphicObjects[i].FindNextFillingForm) {
						Id = this.arrGraphicObjects[i].FindNextFillingForm(true, i === Current, i === Current);
						if (Id)
							return Id;
					}
				}
			} else {
				var Start = (-1 !== Current ? Current : Len - 1);

				for (var i = Start; i >= 0; i--) {
					if (this.arrGraphicObjects[i].FindNextFillingForm) {
						Id = this.arrGraphicObjects[i].FindNextFillingForm(false, i === Current, i === Current);
						if (Id)
							return Id;
					}
				}
			}

			return null;
		};

		CGroupShape.prototype.getCompiledFill = function () {
			this.compiledFill = null;
			if (isRealObject(this.spPr) && isRealObject(this.spPr.Fill) && isRealObject(this.spPr.Fill.fill)) {
				this.compiledFill = this.spPr.Fill.createDuplicate();
				if (this.compiledFill && this.compiledFill.fill && this.compiledFill.fill.type === Asc.c_oAscFill.FILL_TYPE_GRP) {
					if (this.group) {
						var group_compiled_fill = this.group.getCompiledFill();
						if (isRealObject(group_compiled_fill) && isRealObject(group_compiled_fill.fill)) {
							this.compiledFill = group_compiled_fill.createDuplicate();
						} else {
							this.compiledFill = null;
						}
					} else {
						this.compiledFill = null;
					}
				}
			} else if (isRealObject(this.group)) {
				var group_compiled_fill = this.group.getCompiledFill();
				if (isRealObject(group_compiled_fill) && isRealObject(group_compiled_fill.fill)) {
					this.compiledFill = group_compiled_fill.createDuplicate();
				} else {
					var hierarchy = this.getHierarchy();
					for (var i = 0; i < hierarchy.length; ++i) {
						if (isRealObject(hierarchy[i]) && isRealObject(hierarchy[i].spPr) && isRealObject(hierarchy[i].spPr.Fill) && isRealObject(hierarchy[i].spPr.Fill.fill)) {
							this.compiledFill = hierarchy[i].spPr.Fill.createDuplicate();
							break;
						}
					}
				}
			} else {
				var hierarchy = this.getHierarchy();
				for (var i = 0; i < hierarchy.length; ++i) {
					if (isRealObject(hierarchy[i]) && isRealObject(hierarchy[i].spPr) && isRealObject(hierarchy[i].spPr.Fill) && isRealObject(hierarchy[i].spPr.Fill.fill)) {
						this.compiledFill = hierarchy[i].spPr.Fill.createDuplicate();
						break;
					}
				}
			}
			return this.compiledFill;
		};

		CGroupShape.prototype.getCompiledLine = function () {
			return null;
		};

		CGroupShape.prototype.setVerticalAlign = function (align) {
			for (var _shape_index = 0; _shape_index < this.spTree.length; ++_shape_index) {
				if (this.spTree[_shape_index].setVerticalAlign) {
					this.spTree[_shape_index].setVerticalAlign(align);
				}
			}
		};

		CGroupShape.prototype.setVert = function (vert) {
			for (var _shape_index = 0; _shape_index < this.spTree.length; ++_shape_index) {
				if (this.spTree[_shape_index].setVert) {
					this.spTree[_shape_index].setVert(vert);
				}
			}
		};

		CGroupShape.prototype.setPaddings = function (paddings) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].setPaddings) {
					this.spTree[i].setPaddings(paddings);
				}
			}
		};
		CGroupShape.prototype.setTextFitType = function (type) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].setTextFitType) {
					this.spTree[i].setTextFitType(type);
				}
			}
		};

		CGroupShape.prototype.setVertOverflowType = function (type) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].setVertOverflowType) {
					this.spTree[i].setVertOverflowType(type);
				}
			}
		};

		CGroupShape.prototype.setColumnNumber = function (num) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].setColumnNumber) {
					this.spTree[i].setColumnNumber(num);
				}
			}
		};

		CGroupShape.prototype.setColumnSpace = function (spcCol) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].setColumnSpace) {
					this.spTree[i].setColumnSpace(spcCol);
				}
			}
		};

		CGroupShape.prototype.changePresetGeom = function (preset) {
			for (var _shape_index = 0; _shape_index < this.spTree.length; ++_shape_index) {
				if (this.spTree[_shape_index].getObjectType() === AscDFH.historyitem_type_Shape) {
					this.spTree[_shape_index].changePresetGeom(preset);
				}
			}
		};

		CGroupShape.prototype.changeShadow = function (oShadow) {

			for (var _shape_index = 0; _shape_index < this.spTree.length; ++_shape_index) {
				if (this.spTree[_shape_index].changeShadow) {
					this.spTree[_shape_index].changeShadow(oShadow);
				}
			}
		};

		CGroupShape.prototype.changeFill = function (fill) {
			for (var _shape_index = 0; _shape_index < this.spTree.length; ++_shape_index) {
				if (this.spTree[_shape_index].changeFill) {
					this.spTree[_shape_index].changeFill(fill);
				}
			}
		};

		CGroupShape.prototype.changeLine = function (line) {
			for (var _shape_index = 0; _shape_index < this.spTree.length; ++_shape_index) {
				if (this.spTree[_shape_index].changeLine) {
					this.spTree[_shape_index].changeLine(line);
				}
			}
		};
		CGroupShape.prototype.isGroup = function () {
			return true;
		};
		CGroupShape.prototype.normalize = function () {
			for (var i = 0; i < this.spTree.length; ++i) {
				this.spTree[i].normalize();
			}
			AscFormat.CGraphicObjectBase.prototype.normalize.call(this);
			var xfrm = this.spPr.xfrm;
			xfrm.setChExtX(xfrm.extX);
			xfrm.setChExtY(xfrm.extY);
			xfrm.setChOffX(0);
			xfrm.setChOffY(0);
		};

		CGroupShape.prototype.updateCoordinatesAfterInternalResize = function () {
			this.normalize();
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].isGroup())
					this.spTree[i].updateCoordinatesAfterInternalResize();
			}

			var sp_tree = this.spTree;

			var min_x, max_x, min_y, max_y;
			var sp = sp_tree[0];
			var xfrm = sp.spPr.xfrm;
			var rot = xfrm.rot == null ? 0 : xfrm.rot;
			var xc, yc;
			if (AscFormat.checkNormalRotate(rot)) {
				min_x = xfrm.offX;
				min_y = xfrm.offY;
				max_x = xfrm.offX + xfrm.extX;
				max_y = xfrm.offY + xfrm.extY;
			} else {
				xc = xfrm.offX + xfrm.extX * 0.5;
				yc = xfrm.offY + xfrm.extY * 0.5;
				min_x = xc - xfrm.extY * 0.5;
				min_y = yc - xfrm.extX * 0.5;
				max_x = xc + xfrm.extY * 0.5;
				max_y = yc + xfrm.extX * 0.5;
			}
			var cur_max_x, cur_min_x, cur_max_y, cur_min_y;
			for (i = 1; i < sp_tree.length; ++i) {
				sp = sp_tree[i];
				xfrm = sp.spPr.xfrm;
				rot = xfrm.rot == null ? 0 : xfrm.rot;

				if (AscFormat.checkNormalRotate(rot)) //  || (this.getName && this.getName() === 'Drawing')
				{
					cur_min_x = xfrm.offX;
					cur_min_y = xfrm.offY;
					cur_max_x = xfrm.offX + xfrm.extX;
					cur_max_y = xfrm.offY + xfrm.extY;
				} else {
					xc = xfrm.offX + xfrm.extX * 0.5;
					yc = xfrm.offY + xfrm.extY * 0.5;
					cur_min_x = xc - xfrm.extY * 0.5;
					cur_min_y = yc - xfrm.extX * 0.5;
					cur_max_x = xc + xfrm.extY * 0.5;
					cur_max_y = yc + xfrm.extX * 0.5;
				}
				if (cur_max_x > max_x)
					max_x = cur_max_x;
				if (cur_min_x < min_x)
					min_x = cur_min_x;

				if (cur_max_y > max_y)
					max_y = cur_max_y;
				if (cur_min_y < min_y)
					min_y = cur_min_y;
			}

			var temp;
			var x_min_clear = min_x;
			var y_min_clear = min_y;
			if (this.spPr.xfrm.flipH === true) {
				temp = max_x;
				max_x = this.spPr.xfrm.extX - min_x;
				min_x = this.spPr.xfrm.extX - temp;
			}

			if (this.spPr.xfrm.flipV === true) {
				temp = max_y;
				max_y = this.spPr.xfrm.extY - min_y;
				min_y = this.spPr.xfrm.extY - temp;
			}

			var old_x0, old_y0;
			var xfrm = this.spPr.xfrm;
			var rot = xfrm.rot == null ? 0 : xfrm.rot;
			var hc = xfrm.extX * 0.5;
			var vc = xfrm.extY * 0.5;
			old_x0 = this.spPr.xfrm.offX + hc - (hc * Math.cos(rot) - vc * Math.sin(rot));
			old_y0 = this.spPr.xfrm.offY + vc - (hc * Math.sin(rot) + vc * Math.cos(rot));
			var t_dx = min_x * Math.cos(rot) - min_y * Math.sin(rot);
			var t_dy = min_x * Math.sin(rot) + min_y * Math.cos(rot);
			var new_x0, new_y0;
			new_x0 = old_x0 + t_dx;
			new_y0 = old_y0 + t_dy;
			var new_hc = Math.abs(max_x - min_x) * 0.5;
			var new_vc = Math.abs(max_y - min_y) * 0.5;
			var new_xc = new_x0 + (new_hc * Math.cos(rot) - new_vc * Math.sin(rot));
			var new_yc = new_y0 + (new_hc * Math.sin(rot) + new_vc * Math.cos(rot));

			var pos_x, pos_y;
			pos_x = new_xc - new_hc;
			pos_y = new_yc - new_vc;

			var xfrm = this.spPr.xfrm;
			if (this.group || !(editor && editor.isDocumentEditor)) {
				xfrm.setOffX(pos_x);
				xfrm.setOffY(pos_y);
			}
			xfrm.setExtX(Math.abs(max_x - min_x));
			xfrm.setExtY(Math.abs(max_y - min_y));
			xfrm.setChExtX(Math.abs(max_x - min_x));
			xfrm.setChExtY(Math.abs(max_y - min_y));
			xfrm.setChOffX(0);
			xfrm.setChOffY(0);
			for (i = 0; i < sp_tree.length; ++i) {
				sp_tree[i].spPr.xfrm.setOffX(sp_tree[i].spPr.xfrm.offX - x_min_clear);
				sp_tree[i].spPr.xfrm.setOffY(sp_tree[i].spPr.xfrm.offY - y_min_clear);
			}
			this.checkDrawingBaseCoords();
			return {posX: pos_x, posY: pos_y};
		};


		CGroupShape.prototype.getParentObjects = function () {
			var parents = {slide: null, layout: null, master: null, theme: null};
			return parents;
		};

		CGroupShape.prototype.applyTextArtForm = function (sPreset) {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].applyTextArtForm) {
					this.spTree[i].applyTextArtForm(sPreset);
				}
			}
		};

		CGroupShape.prototype.createRotateTrack = function () {
			return new AscFormat.RotateTrackGroup(this);
		};

		CGroupShape.prototype.createMoveTrack = function () {
			return new AscFormat.MoveGroupTrack(this);
		};

		CGroupShape.prototype.createResizeTrack = function (cardDirection) {
			return new AscFormat.ResizeTrackGroup(this, cardDirection);
		};

		CGroupShape.prototype.resetSelection = function (graphicObjects) {
			this.selection.textSelection = null;
			if (this.selection.chartSelection) {
				this.selection.chartSelection.resetSelection();
			}
			this.selection.chartSelection = null;

			for (var i = this.selectedObjects.length - 1; i > -1; --i) {
				var old_gr = this.selectedObjects[i].group;
				var obj = this.selectedObjects[i];
				obj.group = this;
				obj.deselect(graphicObjects);
				obj.group = old_gr;
			}
		};

		CGroupShape.prototype.resetInternalSelection = AscFormat.DrawingObjectsController.prototype.resetInternalSelection;

		CGroupShape.prototype.recalculateCurPos = AscFormat.DrawingObjectsController.prototype.recalculateCurPos;

		CGroupShape.prototype.loadDocumentStateAfterLoadChanges = AscFormat.DrawingObjectsController.prototype.loadDocumentStateAfterLoadChanges;
		CGroupShape.prototype.getAllConnectors = AscFormat.DrawingObjectsController.prototype.getAllConnectors;
		CGroupShape.prototype.getAllShapes = AscFormat.DrawingObjectsController.prototype.getAllShapes;


		CGroupShape.prototype.calculateSnapArrays = function (snapArrayX, snapArrayY) {
			var sp;
			for (var i = 0; i < this.arrGraphicObjects.length; ++i) {
				sp = this.arrGraphicObjects[i];
				sp.calculateSnapArrays(snapArrayX, snapArrayY);
				sp.recalculateSnapArrays();
			}
		};

		CGroupShape.prototype.setNvSpPr = function (pr) {
			History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_GroupShapeSetNvGrpSpPr, this.nvGrpSpPr, pr));
			this.nvGrpSpPr = pr;
		};

		CGroupShape.prototype.RestartSpellCheck = function () {
			for (var i = 0; i < this.spTree.length; ++i) {
				this.spTree[i].RestartSpellCheck && this.spTree[i].RestartSpellCheck();
			}
		};

		CGroupShape.prototype.recalculateLocalTransform = CShape.prototype.recalculateLocalTransform;

		CGroupShape.prototype.bringToFront = function ()//перемещаем заселекченые объекты наверх
		{
			var i;
			var arrDrawings = [];
			for (i = this.spTree.length - 1; i > -1; --i) {
				if (this.spTree[i].isGroupObject()) {
					this.spTree[i].bringToFront();
				} else if (this.spTree[i].selected) {
					arrDrawings.push(this.removeFromSpTreeByPos(i));
				}
			}
			for (i = arrDrawings.length - 1; i > -1; --i) {
				this.addToSpTree(null, arrDrawings[i]);
			}
		};

		CGroupShape.prototype.bringForward = function () {
			var i;
			for (i = this.spTree.length - 1; i > -1; --i) {
				if (this.spTree[i].isGroupObject()) {
					this.spTree[i].bringForward();
				} else if (i < this.spTree.length - 1 && this.spTree[i].selected && !this.spTree[i + 1].selected) {
					var item = this.removeFromSpTreeByPos(i);
					this.addToSpTree(i + 1, item);
				}
			}
		};

		CGroupShape.prototype.sendToBack = function () {
			var i, arrDrawings = [];
			for (i = this.spTree.length - 1; i > -1; --i) {
				if (this.spTree[i].isGroupObject()) {
					this.spTree[i].sendToBack();
				} else if (this.spTree[i].selected) {
					arrDrawings.push(this.removeFromSpTreeByPos(i));
				}
			}
			arrDrawings.reverse();
			for (i = 0; i < arrDrawings.length; ++i) {
				this.addToSpTree(i, arrDrawings[i]);
			}
		};

		CGroupShape.prototype.bringBackward = function () {
			var i;
			for (i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].isGroupObject()) {
					this.spTree[i].bringBackward();
				} else if (i > 0 && this.spTree[i].selected && !this.spTree[i - 1].selected) {
					this.addToSpTree(i - 1, this.removeFromSpTreeByPos(i));
				}
			}
		};
		CGroupShape.prototype.recalcSmartArtConnections = function () {};
		CGroupShape.prototype.Refresh_RecalcData = function (oData) {
			if (oData) {
				switch (oData.Type) {

					case AscDFH.historyitem_ShapeSetBDeleted: {
						if (!this.bDeleted) {
							this.addToRecalculate();
						}
						break;
					}
					case AscDFH.historyitem_GroupShapeAddToSpTree:
					case AscDFH.historyitem_GroupShapeRemoveFromSpTree: {
						if (!this.bDeleted) {
							this.handleUpdateSpTree();
							this.recalcSmartArtConnections();
						}
						break;
					}
					case AscDFH.historyitem_SmartArtDrawing: {
						if (!this.bDeleted) {
							this.addToRecalculate();
						}
						break;
					}
				}
			}
		};

		CGroupShape.prototype.checkTypeCorrect = function () {
			if (!this.spPr) {
				return false;
			}
			if (this.spTree.length === 0) {
				return false;
			}
			return true;
		};

		CGroupShape.prototype.resetGroups = function () {
			for (var i = 0; i < this.spTree.length; ++i) {
				if (this.spTree[i].resetGroups) {
					this.spTree[i].resetGroups();
				}
				if (this.spTree[i].group !== this) {
					this.spTree[i].setGroup(this);
				}
			}
		};

		CGroupShape.prototype.findConnector = function (x, y) {
			for (var i = this.spTree.length - 1; i > -1; --i) {
				var oConInfo = this.spTree[i].findConnector(x, y);
				if (oConInfo) {
					return oConInfo;
				}
			}
			return null;
		};

		CGroupShape.prototype.findConnectionShape = function (x, y) {
			for (var i = this.spTree.length - 1; i > -1; --i) {
				var _ret = this.spTree[i].findConnectionShape(x, y);
				if (_ret) {
					return _ret;
				}
			}
			return null;
		};

		CGroupShape.prototype.GetAllContentControls = function (arrContentControls) {
			for (var i = 0; i < this.spTree.length; ++i) {
				this.spTree[i].GetAllContentControls(arrContentControls);
			}
		};

		CGroupShape.prototype.getCopyWithSourceFormatting = function (oIdMap) {
			var oPr = new AscFormat.CCopyObjectProperties();
			oPr.idMap = oIdMap;
			oPr.bSaveSourceFormatting = true;
			return this.copy(oPr);
		};

		CGroupShape.prototype.GetAllFields = function (isUseSelection, arrFields) {
			var _arrFields = arrFields ? arrFields : [], i;
			if (isUseSelection) {
				for (i = 0; i < this.selectedObjects.length; ++i) {
					this.selectedObjects[i].GetAllFields(isUseSelection, _arrFields);
				}
			} else {
				for (i = 0; i < this.spTree.length; ++i) {
					this.spTree[i].GetAllFields(isUseSelection, _arrFields);
				}
			}
			return _arrFields;
		};
		CGroupShape.prototype.GetAllSeqFieldsByType = function (sType, aFields) {
			for (var i = 0; i < this.spTree.length; ++i) {
				this.spTree[i].GetAllSeqFieldsByType(sType, aFields)
			}
		};
		CGroupShape.prototype.createPlaceholderControl = function (aControls) {
			for (var i = 0; i < this.spTree.length; ++i) {
				this.spTree[i].createPlaceholderControl(aControls);
			}
		};
		CGroupShape.prototype.onSlicerUpdate = function (sName) {
			var bRet = false;
			for (var i = 0; i < this.spTree.length; ++i) {
				bRet = bRet || this.spTree[i].onSlicerUpdate(sName);
			}
			return bRet;
		};
		CGroupShape.prototype.onSlicerDelete = function (sName) {
			var bRet = false;
			for (var i = 0; i < this.spTree.length; ++i) {
				bRet = bRet || this.spTree[i].onSlicerDelete(sName);
			}
			return bRet;
		};
		CGroupShape.prototype.onTimeSlicerDelete = function (sName) {
			var bRet = false;
			for (var i = 0; i < this.spTree.length; ++i) {
				bRet = bRet || this.spTree[i].onTimeSlicerDelete(sName);
			}
			return bRet;
		};
		CGroupShape.prototype.onSlicerLock = function (sName, bLock) {
			for (var i = 0; i < this.spTree.length; ++i) {
				this.spTree[i].onSlicerLock(sName, bLock);
			}
		};
		CGroupShape.prototype.onSlicerChangeName = function (sName, sNewName) {
			for (var i = 0; i < this.spTree.length; ++i) {
				this.spTree[i].onSlicerChangeName(sName, sNewName);
			}
		};
		CGroupShape.prototype.getSlicerViewByName = function (name) {
			var res = null;
			for (var i = 0; i < this.spTree.length; ++i) {
				res = this.spTree[i].getSlicerViewByName(name);
				if (res) {
					return res;
				}
			}
			return res;
		};

		CGroupShape.prototype.handleObject = function (fCallback) {
			fCallback(this);
			for (var nSp = 0; nSp < this.spTree.length; ++nSp) {
				this.spTree[nSp].handleObject(fCallback);
			}
		};

		//for bug 52775. remove in the next version
		CGroupShape.prototype.applySmartArtTextStyle = function () {
			for (var nSp = 0; nSp < this.spTree.length; ++nSp) {
				this.spTree[nSp].applySmartArtTextStyle();
			}
		};
		CGroupShape.prototype.getTypeName = function () {
			return AscCommon.translateManager.getValue("Group");
		};
		CGroupShape.prototype.GetAllOleObjects = function (sPluginId, arrObjects) {
			for (let nSp = 0; nSp < this.spTree.length; ++nSp) {
				this.spTree[nSp].GetAllOleObjects(sPluginId, arrObjects);
			}
		};


		CGroupShape.prototype.checkXfrm = function () {
			if (!this.spPr) {
				return;
			}
			if (!this.spPr.xfrm && this.spTree.length > 0) {
				var oXfrm = new AscFormat.CXfrm();
				oXfrm.setOffX(0);
				oXfrm.setOffY(0);
				oXfrm.setChOffX(0);
				oXfrm.setChOffY(0);
				oXfrm.setExtX(50);
				oXfrm.setExtY(50);
				oXfrm.setChExtX(50);
				oXfrm.setChExtY(50);
				this.spPr.setXfrm(oXfrm);
				this.updateCoordinatesAfterInternalResize();
				this.spPr.xfrm.setParent(this.spPr);
			}
		};
		CGroupShape.prototype.clearChartDataCache = function () {
			for (let nSp = 0; nSp < this.spTree.length; ++nSp) {
				this.spTree[nSp].clearChartDataCache();
			}
		};
		CGroupShape.prototype.pasteFormatting = function (oFormatData) {
			if(!oFormatData) {
				return;
			}
			for (let nSp = 0; nSp < this.spTree.length; ++nSp) {
				this.spTree[nSp].pasteFormatting(oFormatData);
			}
		};

		CGroupShape.prototype.compareForMorph = function(oDrawingToCheck, oCurCandidate, oMapPaired) {
			if(this.getObjectType() !== oDrawingToCheck.getObjectType()) {
				return oCurCandidate;
			}
			const sName = this.getOwnName();
			if(sName && sName.startsWith(AscFormat.OBJECT_MORPH_MARKER)) {
				const sCheckName = oDrawingToCheck.getOwnName();
				if(sName !== sCheckName) {
					return oCurCandidate;
				}
				return oDrawingToCheck;
			}
			if(this.spTree.length !== oDrawingToCheck.spTree.length) {
				return oCurCandidate;
			}
			for(let nSp = 0; nSp < this.spTree.length; ++nSp) {
				let oSp = this.spTree[nSp];
				let oSpCheck = oDrawingToCheck.spTree[nSp];
				if(!oSp.compareForMorph(oSpCheck, null)) {
					return oCurCandidate;
				}
			}
			if(!oCurCandidate) {
				if(oMapPaired && oMapPaired[oDrawingToCheck.Id]) {
					let oParedDrawing = oMapPaired[oDrawingToCheck.Id].drawing;
					if(oParedDrawing.getOwnName() === oDrawingToCheck.getOwnName()) {
						return oCurCandidate;
					}

					let dDistMCheck = oParedDrawing.getDistanceL1(oDrawingToCheck);
					let dDistMCur = this.getDistanceL1(oDrawingToCheck);
					if(dDistMCheck < dDistMCur) {
						return oCurCandidate;
					}

					let dSizeMCandidate = Math.abs(oParedDrawing.extX - oDrawingToCheck.extX) + Math.abs(oParedDrawing.extY - oDrawingToCheck.extY);
					let dSizeMCheck = Math.abs(oDrawingToCheck.extX - this.extX) + Math.abs(oDrawingToCheck.extY - this.extY);
					if(dSizeMCandidate < dSizeMCheck) {
						return oCurCandidate;
					}
				}
				return oDrawingToCheck;
			}
			const dDistCheck = this.getDistanceL1(oDrawingToCheck);
			const dDistCur = this.getDistanceL1(oCurCandidate);
			let dSizeMCandidate = Math.abs(oCurCandidate.extX - this.extX) + Math.abs(oCurCandidate.extY - this.extY);
			let dSizeMCheck = Math.abs(oDrawingToCheck.extX - this.extX) + Math.abs(oDrawingToCheck.extY - this.extY);
			if(dDistCur < dDistCheck) {
				return  oCurCandidate;
			}
			else {
				if(dSizeMCandidate < dSizeMCheck) {
					return  oCurCandidate;
				}
			}
			if(!oMapPaired || !oMapPaired[oDrawingToCheck.Id]) {
				return oDrawingToCheck;
			}
			else {
				let oParedDrawing = oMapPaired[oDrawingToCheck.Id].drawing;
				if(oParedDrawing.getOwnName() === oDrawingToCheck.getOwnName()) {
					return oCurCandidate;
				}
				else {
					return oDrawingToCheck;
				}
			}
			return oCurCandidate;
		};
		CGroupShape.prototype.checkDrawingPartWithHistory = function() {};
		CGroupShape.prototype.generateSmartArtDrawingPart = function () {
			for (let i = 0; i < this.spTree.length; i += 1) {
				this.spTree[i].generateSmartArtDrawingPart();
			}
		};
		CGroupShape.prototype.isHaveOnlyInks = function () {
			if (!this.spTree.length) {
				return false;
			}
			for (let i = 0; i < this.spTree.length; i++) {
				const oDrawing = this.spTree[i];
				if (!(oDrawing.isInk() || oDrawing.isHaveOnlyInks())) {
					return false;
				}
			}
			return true;
		};

		CGroupShape.prototype.getAllInks = function (arrInks) {
			arrInks = arrInks || [];
			for (let i = this.spTree.length - 1; i >= 0; i -= 1) {
				const oDrawing = this.spTree[i];
				if (oDrawing.isInk() || oDrawing.isHaveOnlyInks()) {
					arrInks.push(oDrawing);
				} else {
					oDrawing.getAllInks(arrInks);
				}
			}
			return arrInks;
		};

		//--------------------------------------------------------export----------------------------------------------------
		window['AscFormat'] = window['AscFormat'] || {};
		window['AscFormat'].CGroupShape = CGroupShape;
	})(window);
