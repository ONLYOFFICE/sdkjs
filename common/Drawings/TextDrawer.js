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

(function (window, undefined) {

// Import
	var g_fontApplication = AscFonts.g_fontApplication;

	var g_oTextMeasurer = AscCommon.g_oTextMeasurer;

	var Geometry = AscFormat.Geometry;
	var EPSILON_TEXT_AUTOFIT = AscFormat.EPSILON_TEXT_AUTOFIT;
	var ObjectToDraw = AscFormat.ObjectToDraw;

	var PATH_DIV_EPSILON = 0.1;
	var UNDERLINE_DIV_EPSILON = 3;


	function ParaDrawingStruct(nPageIndex, pDrawing) {
		this.oDrawing = pDrawing;
		this.nPageIndex = nPageIndex;
	}

	ParaDrawingStruct.prototype.Draw = function (pGraphics) {
		if (this.oDrawing) {
			this.oDrawing.Draw(0, 0, pGraphics, this.nPageIndex, 0)
		}
	};

	function CDocContentStructure() {
		this.m_nType = DRAW_COMMAND_CONTENT;
		this.m_aContent = [];
		this.m_aByLines = null;
		this.m_aDrawingsStruct = [];
		this.m_aBackgrounds = [];
		this.m_aBorders = [];
		this.m_aParagraphBackgrounds = [];
		this.m_oBoundsRect1 = null;
		this.m_oBoundsRect2 = null;
		this.m_aComments = [];
	}

	CDocContentStructure.prototype.forEachAnimationObjectToDraw = function (fCallback) {
		for (let i = 0; i < this.m_aContent.length; i++) {
			const oParagraph = this.m_aContent[i];
			for (let j = 0; j < oParagraph.m_aSplitBackgroundsByIterationType.length; j += 1) {
				const aWord = oParagraph.m_aSplitBackgroundsByIterationType[j];
				for (let k = 0; k < aWord.length; k += 1) {
					const oObjectToDraw = aWord[k];
					fCallback(oObjectToDraw);
				}
			}
			for (let j = 0; j < oParagraph.m_aWords.length; j += 1) {
				const aWord = oParagraph.m_aWords[j];
				for (let k = 0; k < aWord.length; k += 1) {
					const oObjectToDraw = aWord[k];
					fCallback(oObjectToDraw);
				}
			}
			for (let j = 0; j < oParagraph.m_aSplitForegroundsByIterationType.length; j += 1) {
				const aWord = oParagraph.m_aSplitForegroundsByIterationType[j];
				for (let k = 0; k < aWord.length; k += 1) {
					const oObjectToDraw = aWord[k];
					fCallback(oObjectToDraw);
				}
			}
		}
	};
	CDocContentStructure.prototype.Recalculate = function (oTheme, oColorMap, dWidth, dHeight, oShape) {
		for (var i = 0; i < this.m_aContent.length; ++i) {
			this.m_aContent[i].Recalculate(oTheme, oColorMap, dWidth, dHeight, oShape);
		}
	};
	CDocContentStructure.prototype.CheckContentStructs = function (aContentStructs) {
		for (let nElement = 0; nElement < this.m_aContent.length; ++nElement) {
			let oElement = this.m_aContent[nElement];
			oElement.CheckContentStructs(aContentStructs);
		}
	};
	CDocContentStructure.prototype.draw = function (graphics, transform, oTheme, oColorMap) {
		var i;
		for (i = 0; i < this.m_aDrawingsStruct.length; ++i) {
			this.m_aDrawingsStruct[i].Draw(graphics);
		}

		for (i = 0; i < this.m_aParagraphBackgrounds.length; ++i) {
			this.m_aParagraphBackgrounds[i].draw(graphics, undefined, transform, oTheme, oColorMap);
		}
		for (i = 0; i < this.m_aBorders.length; ++i) {
			this.m_aBorders[i].draw(graphics);
		}
		for (i = 0; i < this.m_aBackgrounds.length; ++i) {
			this.m_aBackgrounds[i].draw(graphics, undefined, transform, oTheme, oColorMap);
		}
		for (i = 0; i < this.m_aContent.length; ++i) {
			this.m_aContent[i].draw(graphics, transform, oTheme, oColorMap);
		}
	};
	CDocContentStructure.prototype.drawComments = function (graphics, transform) {
		var i;
		for (i = 0; i < this.m_aComments.length; ++i) {
			this.m_aComments[i].drawComment2(graphics, undefined, transform);
		}
	};

	CDocContentStructure.prototype.checkByWarpStruct = function (oWarpStruct, dWidth, dHeight, oTheme, oColorMap, oShape, dOneLineWidth, XLimit, dContentHeight, dKoeff) {
		var i, j, t, aByPaths, aWarpedObjects = [];
		var bOddPaths = oWarpStruct.pathLst.length / 2 - ((oWarpStruct.pathLst.length / 2) >> 0) > 0;
		var nDivCount = bOddPaths ? oWarpStruct.pathLst.length : oWarpStruct.pathLst.length >> 1, oNextPointOnPolygon,
			oObjectToDrawNext, oMatrix = new AscCommon.CMatrix();
		aByPaths = oWarpStruct.getArrayPolygonsByPaths(PATH_DIV_EPSILON);
		var nLastIndex = 0, dTmp, oBoundsChecker, oTemp, nIndex, aWarpedObjects2 = [];
		oBoundsChecker = new AscFormat.CSlideBoundsChecker();
		oBoundsChecker.init(100, 100, 100, 100);
		var dMinX, dMaxX;
		for (j = 0; j < this.m_aByLines.length; ++j) {
			oTemp = this.m_aByLines[j];
			for (t = 0; t < oTemp.length; ++t) {
				oTemp[t].GetAllWarped(aWarpedObjects2);
				oTemp[t].CheckBoundsWarped(oBoundsChecker);
			}
		}
		dMinX = oBoundsChecker.Bounds.min_x;
		dMaxX = oBoundsChecker.Bounds.max_x;
		for (i = 0; i < nDivCount; ++i) {
			dTmp = (this.m_aByLines.length - nLastIndex) / (nDivCount - i);
			nIndex = nLastIndex + (dTmp >> 0) + ((dTmp - (dTmp >> 0)) > 0 ? 1 : 0);
			aWarpedObjects.length = 0;
			oBoundsChecker.Bounds.ClearNoAttack();
			for (j = nLastIndex; j < nIndex; ++j) {
				oTemp = this.m_aByLines[j];
				for (t = 0; t < oTemp.length; ++t) {
					oTemp[t].GetAllWarped(aWarpedObjects);
					oTemp[t].CheckBoundsWarped(oBoundsChecker);
				}
			}
			if (oBoundsChecker.Bounds.min_x > dMinX) {
				oBoundsChecker.Bounds.min_x = dMinX;
			}
			if (oBoundsChecker.Bounds.max_x < dMaxX) {
				oBoundsChecker.Bounds.max_x = dMaxX;
			}
			if (!bOddPaths) {
				ObjectsToDrawBetweenTwoPolygons(aWarpedObjects, oBoundsChecker.Bounds, new PolygonWrapper(aByPaths[i << 1]), new PolygonWrapper(aByPaths[(i << 1) + 1]));
			}
			else {
				oNextPointOnPolygon = null;
				oObjectToDrawNext = null;
				var oPolygon = new PolygonWrapper(aByPaths[i]);
				for (t = 0; t < aWarpedObjects.length; ++t) {
					var oWarpedObject = aWarpedObjects[t];
					var bArcDown = "textArchDown" !== oWarpStruct.preset && i < 1;
					if (!AscFormat.isRealNumber(oWarpedObject.x) || !AscFormat.isRealNumber(oWarpedObject.y)) {
						oWarpedObject.geometry.checkByPolygon(oPolygon, bArcDown, XLimit * dKoeff, dContentHeight, dKoeff, nDivCount > 1 ? oBoundsChecker.Bounds : null);
					}
					else {
						oNextPointOnPolygon = this.checkTransformByOddPath(oMatrix, oWarpedObject, aWarpedObjects[t + 1], oNextPointOnPolygon, dContentHeight, dKoeff, bArcDown, oPolygon, XLimit);
						oWarpedObject.geometry.transform(oMatrix, dKoeff);
					}
				}
			}
			nLastIndex = nIndex;
		}
		this.checkUnionPaths(aWarpedObjects2);
	};
	CDocContentStructure.prototype.checkContentReduct = function (oWarpStruct, dWidth, dHeight, oTheme, oColorMap, oShape, dOneLineWidth, XLimit, dContentHeight, dKoeff) {
		var i, j, t, aByPaths, aWarpedObjects = [];
		var bOddPaths = oWarpStruct.pathLst.length / 2 - ((oWarpStruct.pathLst.length / 2) >> 0) > 0;
		var nDivCount = bOddPaths ? oWarpStruct.pathLst.length : oWarpStruct.pathLst.length >> 1, oNextPointOnPolygon,
			oObjectToDrawNext, oMatrix;
		aByPaths = oWarpStruct.getArrayPolygonsByPaths(PATH_DIV_EPSILON);
		var nLastIndex = 0, dTmp, oBoundsChecker, oTemp, nIndex, aWarpedObjects2 = [];
		oBoundsChecker = new AscFormat.CSlideBoundsChecker();
		oBoundsChecker.init(100, 100, 100, 100);
		for (j = 0; j < this.m_aByLines.length; ++j) {
			oTemp = this.m_aByLines[j];
			for (t = 0; t < oTemp.length; ++t) {
				oTemp[t].GetAllWarped(aWarpedObjects2);
				oTemp[t].CheckBoundsWarped(oBoundsChecker);
			}
		}
		var aBounds = [];
		var oRet = {maxDx: 0, maxDy: 0};
		for (i = 0; i < nDivCount; ++i) {
			dTmp = (this.m_aByLines.length - nLastIndex) / (nDivCount - i);
			nIndex = nLastIndex + (dTmp >> 0) + ((dTmp - (dTmp >> 0)) > 0 ? 1 : 0);
			aWarpedObjects.length = 0;
			for (j = nLastIndex; j < nIndex; ++j) {
				oTemp = this.m_aByLines[j];
				for (t = 0; t < oTemp.length; ++t) {
					oTemp[t].GetAllWarped(aWarpedObjects);
				}
			}
			var bArcDown = "textArchDown" !== oWarpStruct.preset && i < 1;
			if (bOddPaths) {
				oNextPointOnPolygon = null;
				oObjectToDrawNext = null;
				var oPolygon = new PolygonWrapper(aByPaths[i]);
				for (t = 0; t < aWarpedObjects.length; ++t) {
					var oWarpedObject = aWarpedObjects[t];
					if (AscFormat.isRealNumber(oWarpedObject.x) && AscFormat.isRealNumber(oWarpedObject.y)) {
						oMatrix = new AscCommon.CMatrix();
						oBoundsChecker.Bounds.ClearNoAttack();
						oNextPointOnPolygon = this.checkTransformByOddPath(oMatrix, oWarpedObject, aWarpedObjects[t + 1], oNextPointOnPolygon, dContentHeight, dKoeff, bArcDown, oPolygon, XLimit);
						oWarpedObject.draw(oBoundsChecker);

						var oBounds = new AscFormat.CBoundsController();
//                    oBounds.fromBounds(oBoundsChecker.Bounds);
						for (var k = 0; k < aBounds.length; ++k) {
							var oIntersection = this.checkIntersectionBounds(aBounds[k].Bounds, aBounds[k].Transform, oBounds, oMatrix, aBounds.length - k);
							if (oIntersection.DX > oRet.maxDx) {
								oRet.maxDx = oIntersection.DX;
							}
							if (oIntersection.DY < oRet.maxDy) {
								oRet.maxDy = oIntersection.DY;
							}
						}
						aBounds.push({Bounds: oBounds, Transform: oMatrix});
					}
				}
			}
			nLastIndex = nIndex;
		}
		return oRet;
	};

	CDocContentStructure.prototype.getAllBackgroundsBorders = function (aParaBackgrounds, aBackgrounds, aBorders, aComments) {
		for (var i = 0; i < this.m_aContent.length; ++i) {
			this.m_aContent[i].getAllBackgroundsBorders(aParaBackgrounds, aBackgrounds, aBorders, aComments);
		}
	};
	CDocContentStructure.prototype.getCombinedGeometry = function (transformMatrix) {
		const paragraphStructures = this.m_aContent;
		const lineStructures = paragraphStructures.reduce(function (acc, paragraphStructure) {
			return acc.concat(paragraphStructure.m_aContent)
		}, []);
		const objectsToDraw = lineStructures.reduce(function (acc, lineStructure) {
			return acc.concat(lineStructure.m_aContent)
		}, []);

		// allPaths - array of Path2 instances
		const allPaths = objectsToDraw.reduce(function (acc, objectToDraw) {
			return acc.concat(objectToDraw.geometry.pathLst)
		}, []);

		// allFormatPaths - array of Path instances
		const allFormatPaths = allPaths.map(function (path) {
			const formatPath = new AscFormat.Path();
			path.convertToBezierCurves(formatPath, transformMatrix || new AscCommon.CMatrix());

			formatPath.ArrPathCommandInfo = formatPath.ArrPathCommand.map(function (pathCommand) {
				switch (pathCommand.id) {
					case AscFormat.moveTo:
					case AscFormat.lineTo:
						return {id: pathCommand.id, X: '' + pathCommand.X * 36000, Y: '' + pathCommand.Y * 36000};
					case AscFormat.bezier4:
						return {
							id: pathCommand.id,
							X0: '' + pathCommand.X0 * 36000,
							X1: '' + pathCommand.X1 * 36000,
							X2: '' + pathCommand.X2 * 36000,
							Y0: '' + pathCommand.Y0 * 36000,
							Y1: '' + pathCommand.Y1 * 36000,
							Y2: '' + pathCommand.Y2 * 36000
						};
					default:
						return {id: pathCommand.id};
				}
			});

			const properties = ['extrusionOk', 'fill', 'parent', 'pathH', 'pathW', 'stroke'];
			properties.forEach(function (property) {
				formatPath[property] = path[property];
			});

			return formatPath;
		});

		const combinedGeometry = new AscFormat.Geometry();
		allFormatPaths.forEach(function (path) {
			if (path.ArrPathCommand.length > 0) {
				combinedGeometry.AddPath(path);
			}
		});

		return combinedGeometry;
	};


	function CheckIntervalIntersection(X1, Y1, X2, Y2, X3, Y3, X4, Y4) {
		return (((X3 - X1) * (Y2 - Y1) - (Y3 - Y1) * (X2 - X1)) * ((X4 - X1) * (Y2 - Y1) - (Y4 - Y1) * (X2 - X1)) >= 0) && (((X1 - X3) * (Y4 - Y3) - (Y1 - Y3) * (X4 - X3)) * ((X2 - X3) * (Y4 - Y3) - (Y2 - Y3) * (X4 - X3)) >= 0);
	}

	function BoundsRect() {
		this.arrBounds = [{X0: 0, Y0: 0, X1: 0, Y1: 0}, {X0: 0, Y0: 0, X1: 0, Y1: 0}, {
			X0: 0,
			Y0: 0,
			X1: 0,
			Y1: 0
		}, {X0: 0, Y0: 0, X1: 0, Y1: 0}]
	}

	BoundsRect.prototype.fromBoundsAndTransform = function (oBounds, oTransform) {
		this.arrBounds[0].X0 = oTransform.TransformPointX(oBounds.min_x, oBounds.min_y);
		this.arrBounds[0].Y0 = oTransform.TransformPointY(oBounds.min_x, oBounds.min_y);
		this.arrBounds[0].X1 = oTransform.TransformPointX(oBounds.max_x, oBounds.min_y);
		this.arrBounds[0].Y1 = oTransform.TransformPointY(oBounds.max_x, oBounds.min_y);

		this.arrBounds[1].X0 = oTransform.TransformPointX(oBounds.max_x, oBounds.max_y);
		this.arrBounds[1].Y0 = oTransform.TransformPointY(oBounds.max_x, oBounds.max_y);
		this.arrBounds[1].X1 = oTransform.TransformPointX(oBounds.max_x, oBounds.min_y);
		this.arrBounds[1].Y1 = oTransform.TransformPointY(oBounds.max_x, oBounds.min_y);

		this.arrBounds[2].X0 = oTransform.TransformPointX(oBounds.max_x, oBounds.max_y);
		this.arrBounds[2].Y0 = oTransform.TransformPointY(oBounds.max_x, oBounds.max_y);
		this.arrBounds[2].X1 = oTransform.TransformPointX(oBounds.min_x, oBounds.max_y);
		this.arrBounds[2].Y1 = oTransform.TransformPointY(oBounds.min_x, oBounds.max_y);

		this.arrBounds[3].X0 = oTransform.TransformPointX(oBounds.min_x, oBounds.max_y);
		this.arrBounds[3].Y0 = oTransform.TransformPointY(oBounds.min_x, oBounds.max_y);
		this.arrBounds[3].X1 = oTransform.TransformPointX(oBounds.min_x, oBounds.min_y);
		this.arrBounds[3].Y1 = oTransform.TransformPointY(oBounds.min_x, oBounds.min_y);
	};


	BoundsRect.prototype.checkIntersection = function (oBoundsRect) {
		var i, j, oCurBound1, oCurBound2;
		for (i = 0; i < 4; ++i) {
			oCurBound1 = this.arrBounds[i];
			for (j = 0; j < 4; ++j) {
				oCurBound2 = oBoundsRect[j];
				if (CheckIntervalIntersection(oCurBound1.X0, oCurBound1.Y0, oCurBound1.X1, oCurBound1.Y1, oCurBound2.X0, oCurBound2.Y0, oCurBound2.X1, oCurBound2.Y1)) {
					return true;
				}
			}
		}
		return false;
	};

	BoundsRect.prototype.getParamValueProjection = function (x0, y0, x1, y1, x, y) {
		var dX = (x1 - x0), dY = (y1 - y0);

		if (Math.abs(dX) > 0) {
			var dXSq = dX * dX;
			var xi = (x * dXSq + dX * dY * y - dY * (x1 * y0 - x0 * y1)) / (dXSq + dY * dY);
			return (xi - x1) / dX;
		}
		else if (Math.abs(dY) > 0) {
			var dYSq = dY * dY;
			var yi = (y * dYSq + dX * dY * x + dX * (x1 * y0 - x0 * y1)) / (dX * dX + dYSq);
			return (yi - y1) / dY;
		}
		return -1;
	};

	BoundsRect.prototype.checkIntervalProjection = function (oInterval1, oInterval2) {
		var param0 = this.getParamValueProjection(oInterval1.X0, oInterval1.Y0, oInterval1.X1, oInterval1.Y1, oInterval2.X0, oInterval2.Y0);
		var param1 = this.getParamValueProjection(oInterval1.X0, oInterval1.Y0, oInterval1.X1, oInterval1.Y1, oInterval2.X1, oInterval2.Y1);
		if (Math.abs(param0 - 0.5) <= 0.5 || Math.abs(param1 - 0.5) <= 0.5) {
			var dX = (oInterval1.X0 - oInterval1.X1), dY = (oInterval1.Y0 - oInterval1.Y1);
			var dLen = Math.sqrt(dX * dX + dY * dY);
			return ((param1 > 1 ? 1 : (param1 < 0 ? 0 : param1)) - (param0 > 1 ? 1 : (param0 < 0 ? 0 : param0))) * dLen;
		}
		else {
			return 0;
		}
	};


	BoundsRect.prototype.checkBoundsRectProjection = function (oBoundsRect) {
		var dX = 0, dY = 0, dLTemp;

		dLTemp = this.checkIntervalProjection(this.arrBounds[1], oBoundsRect.arrBounds[1]);
		if (dLTemp < dY) {
			dY = dLTemp;
		}
		dLTemp = this.checkIntervalProjection(this.arrBounds[1], oBoundsRect.arrBounds[3]);
		if (dLTemp < dY) {
			dY = dLTemp;
		}
		dLTemp = this.checkIntervalProjection(this.arrBounds[3], oBoundsRect.arrBounds[1]);
		if (dLTemp < dY) {
			dY = dLTemp;
		}
		dLTemp = this.checkIntervalProjection(this.arrBounds[3], oBoundsRect.arrBounds[3]);
		if (dLTemp < dY) {
			dY = dLTemp;
		}

		dLTemp = this.checkIntervalProjection(this.arrBounds[2], oBoundsRect.arrBounds[1]);
		if (Math.abs(dLTemp) > dX) {
			dX = Math.abs(dLTemp);
		}

		dLTemp = this.checkIntervalProjection(this.arrBounds[2], oBoundsRect.arrBounds[3]);
		if (Math.abs(dLTemp) > dX) {
			dX = Math.abs(dLTemp);
		}

		dLTemp = this.checkIntervalProjection(this.arrBounds[0], oBoundsRect.arrBounds[1]);
		if (Math.abs(dLTemp) > dX) {
			dX = Math.abs(dLTemp);
		}

		dLTemp = this.checkIntervalProjection(this.arrBounds[0], oBoundsRect.arrBounds[3]);
		if (Math.abs(dLTemp) > dX) {
			dX = Math.abs(dLTemp);
		}

		return {DX: dX, DY: dY};
	};

	CDocContentStructure.prototype.checkIntersectionBounds = function (oBounds1, oTransform1, oBounds2, oTransform2, nDist) {
		if (!this.m_oBoundsRect1) {
			this.m_oBoundsRect1 = new BoundsRect();
			this.m_oBoundsRect2 = new BoundsRect();
		}
		this.m_oBoundsRect1.fromBoundsAndTransform(oBounds1, oTransform1);
		this.m_oBoundsRect2.fromBoundsAndTransform(oBounds2, oTransform2);
		if (this.m_oBoundsRect1.checkIntersection(this.m_oBoundsRect2.arrBounds)) {
			var oBounds = this.m_oBoundsRect1.checkBoundsRectProjection(this.m_oBoundsRect2);
			oBounds.DX /= nDist;
			return oBounds;
		}
		return {DX: 0, DY: 0};
	};

	CDocContentStructure.prototype.checkTransformByOddPath = function (oMatrix, oWarpedObject, oObjectToDrawNext, oNextPointOnPolygon, dContentHeight, dKoeff, bArcDown, oPolygon, XLimit) {
		var x0, y0, x1, y1, x0t, y0t, x1t, y1t, cX, cX2, dX, dY, dX2, dY2, dNorm, oPointOnPolygon;
		x0 = oWarpedObject.x * dKoeff;
		y0 = oWarpedObject.y * dKoeff;
		x1 = x0;
		if (bArcDown) {
			y1 = 0;
		}
		else {
			y1 = dContentHeight * dKoeff;
		}
		cX = oWarpedObject.x / XLimit;
		var oRet = null;
		if (oNextPointOnPolygon) {
			oPointOnPolygon = oNextPointOnPolygon;
		}
		else {
			oPointOnPolygon = oPolygon.getPointOnPolygon(cX, true);
		}
		x1t = oPointOnPolygon.x;
		y1t = oPointOnPolygon.y;
		dX = oPointOnPolygon.oP2.x - oPointOnPolygon.oP1.x;
		dY = oPointOnPolygon.oP2.y - oPointOnPolygon.oP1.y;

		if (bArcDown) {
			dX = -dX;
			dY = -dY;
		}
		if (oObjectToDrawNext && AscFormat.isRealNumber(oObjectToDrawNext.x) && AscFormat.isRealNumber(oObjectToDrawNext.y) && oObjectToDrawNext.x > oWarpedObject.x) {
			cX2 = (oObjectToDrawNext.x) / XLimit;
			oRet = oPolygon.getPointOnPolygon(cX2, true);
			dX2 = oRet.oP2.x - oRet.oP1.x;
			dY2 = oRet.oP2.y - oRet.oP1.y;
			if (bArcDown) {
				dX2 = -dX2;
				dY2 = -dY2;
			}
			dX = dX + dX2;
			dY = dY + dY2;
		}
		else {
			oRet = null;
		}
		dNorm = Math.sqrt(dX * dX + dY * dY);

		if (bArcDown) {
			x0t = x1t + (dY / dNorm) * (y0);
			y0t = y1t - (dX / dNorm) * (y0);
		}
		else {
			x0t = x1t + (dY / dNorm) * (dContentHeight * dKoeff - y0);
			y0t = y1t - (dX / dNorm) * (dContentHeight * dKoeff - y0);
		}
		oMatrix.shx = (x1t - x0t) / (y1 - y0);
		oMatrix.sy = (y1t - y0t) / (y1 - y0);
		oMatrix.sx = oMatrix.sy;
		oMatrix.shy = -oMatrix.shx;
		oMatrix.tx = x0t - x0 * oMatrix.sx - y0 * oMatrix.shx;
		oMatrix.ty = y0t - x0 * oMatrix.shy - y0 * oMatrix.sy;
		return oRet;
	};

	CDocContentStructure.prototype.checkUnionPaths = function (aWarpedObjects) {
		var aToCheck = [];
		var aWarpedObjects2, i, j, k;
		if (Array.isArray(aWarpedObjects)) {
			aToCheck.push(aWarpedObjects);
		}
		else {
			for (j = 0; j < this.m_aByLines.length; ++j) {
				aWarpedObjects2 = [];
				var oTemp = this.m_aByLines[j];
				for (var t = 0; t < oTemp.length; ++t) {
					oTemp[t].GetAllWarped(aWarpedObjects2);
				}
				aToCheck.push(aWarpedObjects2);
			}
		}
		for (j = 0; j < aToCheck.length; ++j) {
			var aByProps = [], oCurObjectToDraw;
			aWarpedObjects2 = aToCheck[j];
			for (i = 0; i < aWarpedObjects2.length; ++i) {
				oCurObjectToDraw = aWarpedObjects2[i];
				for (k = 0; k < aByProps.length; ++k) {
					var oObjToDraw = aByProps[k];
					if (CompareBrushes(oObjToDraw.brush, oCurObjectToDraw.brush) && ComparePens(oObjToDraw.pen, oCurObjectToDraw.pen) && !oCurObjectToDraw.Comment && !oObjToDraw.Comment && !oCurObjectToDraw.geometry.bDrawSmart && !oObjToDraw.geometry.bDrawSmart) {
						oObjToDraw.geometry.pathLst = oObjToDraw.geometry.pathLst.concat(oCurObjectToDraw.geometry.pathLst);
						oCurObjectToDraw.geometry.pathLst.length = 0;
						break;
					}
				}
				if (k === aByProps.length) {
					aByProps.push(oCurObjectToDraw);
				}
			}
		}
		this.m_aBackgrounds.length = 0;
		this.m_aBorders.length = 0;
		this.getAllBackgroundsBorders(this.m_aParagraphBackgrounds, this.m_aBackgrounds, this.m_aBorders, this.m_aComments);
	};
	CDocContentStructure.prototype.getParagraphStructures = function () {
		return this.m_aContent;
	};


	CDocContentStructure.prototype.Write_ToBinary = function (writer) {
		writer.WriteLong(this.m_aContent.length);
        for(let idx = 0; idx < this.m_aContent.length; ++idx) {
            this.m_aContent[idx].Write_ToBinary(writer);
        }
	};
	CDocContentStructure.prototype.Read_FromBinary = function (reader) {
		let count = reader.GetLong();
        for(let idx = 0; idx < count; ++idx) {
            let paragraph = new CParagraphStructure();
            paragraph.Read_FromBinary(reader);
            this.m_aContent.push(paragraph);
        }
	};

	function CAdditionalElementPosition() {
		this.line = 0;
		this.posInLine = -1;
		this.charIndex = 0;
	}

	CAdditionalElementPosition.prototype.reset = function () {
		this.line = 0;
		this.posInLine = -1;
		this.charIndex = 0;
	};

	function CParagraphStructure(oParagraph) {
		this.m_nType = DRAW_COMMAND_PARAGRAPH;
		this.m_aContent = [];
		this.m_aWords = [];
		this.m_aSplitBackgroundsByIterationType = null;
		this.m_aSplitForegroundsByIterationType = null;
		this.m_oParagraph = oParagraph;
		this.m_oLastWordStart = new CAdditionalElementPosition();
		this.m_aWrapperElementsCache = null;
		this.m_nSplitParagraphType = null;
	}

	CParagraphStructure.prototype.generateWrappersBySplit = function (nSplitParagraphType, oTransform, oTheme, oColorMap, oDrawing) {
		this.m_nSplitParagraphType = nSplitParagraphType;
		switch (this.m_nSplitParagraphType) {
			case AscFormat.ITERATEDATA_TYPE_WORD: {
				this.m_aWrapperElementsCache = this.getCombinedWordWrappers(oTransform, oTheme, oColorMap, oDrawing);
				break;
			}
			case AscFormat.ITERATEDATA_TYPE_LETTER: {
				this.m_aWrapperElementsCache = this.getCombinedLetterWrappers(oTransform, oTheme, oColorMap, oDrawing);
				break;
			}
			default: {
				this.m_aWrapperElementsCache = this.getAllDrawingObjectsWrapper(oTransform, oTheme, oColorMap, oDrawing);
				break;
			}
		}
	};
	CParagraphStructure.prototype.getParagraphWrappersByIterationType = function (nIterationType) {
		if (this.m_nSplitParagraphType === nIterationType) {
			return this.m_aWrapperElementsCache;
		}

		switch (nIterationType) {
			case AscFormat.ITERATEDATA_TYPE_WORD: {
				let nCacheIndex = 0;
				const arrWords = [];
				for (let i = 0; i < this.m_aWords.length; i += 1) {
					const oWord = this.m_aWords[i];
					const nWordLength = oWord.length;
					const arrWordWrappers = this.m_aWrapperElementsCache.slice(nCacheIndex, nWordLength);
					arrWords.push(new AscCommonSlide.CWrapperDrawer(arrWordWrappers));
					nCacheIndex += nWordLength;
				}
				return arrWords;
			}
			case AscFormat.ITERATEDATA_TYPE_LETTER: {
				return [];
			}
			default: {
				return [new AscCommonSlide.CWrapperDrawer(this.m_aWrapperElementsCache)];
			}
		}
	};
	CParagraphStructure.prototype.getCharWrappersByIterationType = function (nIterationType, arrDrawingObjects, oCachedWrapperDrawer, nStartIndex, nEndIndex) {
		switch (nIterationType) {
			case AscFormat.ITERATEDATA_TYPE_WORD: {
				let nCacheIndex = 0;
				for (let i = 0; i < this.m_aWords.length; i += 1) {
					const arrWord = this.m_aWords[i];
					const arrCharWord = [];
					for (let j = 0; j < arrWord.length; j += 1) {
						const oObjectToDraw = arrWord[j];
						if (oObjectToDraw.charIndex >= nStartIndex) {
							if (oObjectToDraw.charIndex <= nEndIndex) {
								arrCharWord.push(this.m_aWrapperElementsCache[nCacheIndex]);
							}
							else {
								return true;
							}
						}
						nCacheIndex += 1;
					}
					if (arrCharWord.length) {
						arrDrawingObjects.push(new AscCommonSlide.CWrapperDrawer(arrCharWord));
					}
				}
				break;
			}
			case AscFormat.ITERATEDATA_TYPE_LETTER: {
				let nCacheIndex = 0;
				for (let i = 0; i < this.m_aWords.length; i += 1) {
					const arrWord = this.m_aWords[i];
					for (let j = 0; j < arrWord.length; j += 1) {
						const oObjectToDraw = arrWord[j];
						if (oObjectToDraw.charIndex >= nStartIndex) {
							if (oObjectToDraw.charIndex <= nEndIndex) {
								arrDrawingObjects.push(this.m_aWrapperElementsCache[nCacheIndex]);
							}
							else {
								return true;
							}
						}
						nCacheIndex += 1;
					}
				}
				break;
			}
			default: {
				let nCacheIndex = 0;
				for (let i = 0; i < this.m_aWords.length; i += 1) {
					const arrWord = this.m_aWords[i];
					for (let j = 0; j < arrWord.length; j += 1) {
						const oObjectToDraw = arrWord[j];
						if (oObjectToDraw.charIndex >= nStartIndex) {
							if (oObjectToDraw.charIndex <= nEndIndex) {
								oCachedWrapperDrawer.addElement(this.m_aWrapperElementsCache[nCacheIndex]);
							}
							else {
								return true;
							}
						}
						nCacheIndex += 1;
					}
				}
				break;
			}
		}
		return false;
	};
	CParagraphStructure.prototype.getWrappers = function () {
		return this.m_aWrapperElementsCache;
	};

	function isPresentationBulletWord(arrWord) {
		for (let i = 0; i < arrWord.length; i++) {
			if (!arrWord[i].isBulletSymbol) {
				return false;
			}
		}
		return true;
	}

	CParagraphStructure.prototype.getObjectToDrawIterator = function (nDrawingType) {
		const oThis = this;
		let nLineIndex = -1;
		let oLine, arrLineContent, nObjectToDrawIndex = 0, oObjectToDraw;
		return function () {
			nObjectToDrawIndex -= 1;
			oObjectToDraw = arrLineContent && arrLineContent[nObjectToDrawIndex];
			if (nObjectToDrawIndex < 0) {
				nLineIndex += 1;
				oLine = oThis.m_aContent[nLineIndex];
				if (oLine) {
					arrLineContent = oLine.getContentByDrawingType(nDrawingType);
					nObjectToDrawIndex = arrLineContent.length - 1;
					oObjectToDraw = arrLineContent[nObjectToDrawIndex];
				}
			}
			return oObjectToDraw;
		}
	};
	CParagraphStructure.prototype.getLetterElementsMap = function (arrRes) {
		const oMapWords = {};
		let nLetterCounter = 0;
		for (let i = 0; i < this.m_aWords.length; i++) {
			const arrWord = this.m_aWords[i];
			for (let j = 0; j < arrWord.length; j++) {
				const oTextObjectToDraw = arrWord[j];
				if (!arrRes[nLetterCounter]) {
					arrRes.push([]);
				}
				const arrResWord = arrRes[nLetterCounter];
				if (oTextObjectToDraw.TextElement) {
					oMapWords[oTextObjectToDraw.TextElement.textDrawerId] = arrResWord;
				}
				for (let k = 0; k < oTextObjectToDraw.SpaceTextElements.length; k += 1) {
					const oSpaceElement = oTextObjectToDraw.SpaceTextElements[k];
					oMapWords[oSpaceElement.textDrawerId] = arrResWord;
				}
				nLetterCounter += 1;
			}
		}
		return oMapWords;
	};
	CParagraphStructure.prototype.getWordElementsMap = function (arrRes) {
		const oMapWords = {};
		for (let i = 0; i < this.m_aWords.length; i++) {
			const arrWord = this.m_aWords[i];
			if (!arrRes[i]) {
				arrRes.push([]);
			}
			const arrResWord = arrRes[i];
			for (let j = 0; j < arrWord.length; j++) {
				const oTextObjectToDraw = arrWord[j];
				if (oTextObjectToDraw.TextElement) {
					oMapWords[oTextObjectToDraw.TextElement.textDrawerId] = arrResWord;
				}
				for (let k = 0; k < oTextObjectToDraw.SpaceTextElements.length; k += 1) {
					const oSpaceElement = oTextObjectToDraw.SpaceTextElements[k];
					oMapWords[oSpaceElement.textDrawerId] = arrResWord;
				}
			}
		}
		return oMapWords;
	};
	CParagraphStructure.prototype.getCombinedGroundElements = function (nDrawingType, oLettersMap) {
		const fGetObjectToDraw = this.getObjectToDrawIterator(nDrawingType);
		let oObjectToDraw;
		let arrAddElements = [];
		while (oObjectToDraw = fGetObjectToDraw()) {
			const oTextElement = oObjectToDraw.TextElement;
			if (oTextElement) {
				const arrLetter = oLettersMap[oTextElement.textDrawerId];
				if (arrLetter) {
					arrLetter.push(oObjectToDraw);
					arrLetter.push.apply(arrLetter, arrAddElements);
					arrAddElements = [];
				}
				else {
					arrAddElements.push(oObjectToDraw);
				}
			}
			else {
				arrAddElements.push(oObjectToDraw);
			}
		}
	};
	CParagraphStructure.prototype.getBackgroundElementsByWord = function () {
		if (!this.m_aSplitBackgroundsByIterationType) {
			this.m_aSplitBackgroundsByIterationType = [];
			const oWordsMap = this.getWordElementsMap(this.m_aSplitBackgroundsByIterationType);
			this.getCombinedGroundElements(CLineStructure_DrawType_Highlights, oWordsMap);
		}
		return this.m_aSplitBackgroundsByIterationType;
	};
	CParagraphStructure.prototype.getForegroundElementsByWord = function () {
		if (!this.m_aSplitForegroundsByIterationType) {
			this.m_aSplitForegroundsByIterationType = [];
			const oWordsMap = this.getWordElementsMap(this.m_aSplitForegroundsByIterationType);
			this.getCombinedGroundElements(CLineStructure_DrawType_Underlines, oWordsMap);
			this.getCombinedGroundElements(CLineStructure_DrawType_DUnderlines, oWordsMap);
			this.getCombinedGroundElements(CLineStructure_DrawType_Strikeouts, oWordsMap);
			this.getCombinedGroundElements(CLineStructure_DrawType_DStrikeouts, oWordsMap);
		}
		return this.m_aSplitForegroundsByIterationType;
	};
	CParagraphStructure.prototype.getBackgroundElementsByLetter = function () {
		if (!this.m_aSplitBackgroundsByIterationType) {
			this.m_aSplitBackgroundsByIterationType = [];
			const oLetterMap = this.getLetterElementsMap(this.m_aSplitBackgroundsByIterationType);
			this.getCombinedGroundElements(CLineStructure_DrawType_Highlights, oLetterMap);
		}
		return this.m_aSplitBackgroundsByIterationType;
	};
	CParagraphStructure.prototype.getForegroundElementsByLetter = function () {
		if (!this.m_aSplitForegroundsByIterationType) {
			this.m_aSplitForegroundsByIterationType = [];
			const oLetterMap = this.getLetterElementsMap(this.m_aSplitForegroundsByIterationType);
			this.getCombinedGroundElements(CLineStructure_DrawType_Underlines, oLetterMap);
			this.getCombinedGroundElements(CLineStructure_DrawType_DUnderlines, oLetterMap);
			this.getCombinedGroundElements(CLineStructure_DrawType_Strikeouts, oLetterMap);
			this.getCombinedGroundElements(CLineStructure_DrawType_DStrikeouts, oLetterMap);
		}
		return this.m_aSplitForegroundsByIterationType;
	};
	CParagraphStructure.prototype.getBackgroundElementsByAll = function () {
		if (!this.m_aSplitBackgroundsByIterationType) {
			this.m_aSplitBackgroundsByIterationType = [];
			const arrElements = [];
			for (let i = 0; i < this.m_aContent.length; i++) {
				const oLine = this.m_aContent[i];
				arrElements.push.apply(arrElements, oLine.m_aHighlights);
			}
			this.m_aSplitBackgroundsByIterationType.push(arrElements);
		}
		return this.m_aSplitBackgroundsByIterationType;
	};
	CParagraphStructure.prototype.getForegroundElementsByAll = function () {
		if (!this.m_aSplitForegroundsByIterationType) {
			this.m_aSplitForegroundsByIterationType = [];
			const arrElements = [];
			for (let i = 0; i < this.m_aContent.length; i++) {
				const oLine = this.m_aContent[i];
				arrElements.push.apply(arrElements, oLine.m_aUnderlines);
				arrElements.push.apply(arrElements, oLine.m_aDUnderlines);
				arrElements.push.apply(arrElements, oLine.m_aStrikeouts);
				arrElements.push.apply(arrElements, oLine.m_aDStrikeouts);
			}
			this.m_aSplitForegroundsByIterationType.push(arrElements);
		}
		return this.m_aSplitForegroundsByIterationType;
	};
	CParagraphStructure.prototype.getCombinedWordWrappers = function (oTransform, oTheme, oColorMap, oDrawing) {
		const arrRes = [];
		let arrWordSymbols = [];
		const arrBackgroundWords = this.getBackgroundElementsByWord();
		const arrForegroundWords = this.getForegroundElementsByWord();
		for (let i = 0; i < this.m_aWords.length; i += 1) {
			const aWord = this.m_aWords[i];
			arrWordSymbols.push.apply(arrWordSymbols, aWord);
			if (isPresentationBulletWord(aWord)) {
				continue;
			}
			const arrBackgroundWord = arrBackgroundWords[i];
			const arrForegroundWord = arrForegroundWords[i];
			const oWrapperWord = new AscCommonSlide.CObjectForDrawArrayWrapper(arrWordSymbols, oTransform, oTheme, oColorMap, oDrawing, arrForegroundWord, arrBackgroundWord);
			arrRes.push(oWrapperWord);
			arrWordSymbols = [];
		}
		if (arrWordSymbols.length) {
			const oWrapperWord = new AscCommonSlide.CObjectForDrawArrayWrapper(arrWordSymbols, oTransform, oTheme, oColorMap, oDrawing);
			arrRes.push(oWrapperWord);
		}
		return arrRes;
	};
	CParagraphStructure.prototype.getCombinedLetterWrappers = function (oTransform, oTheme, oColorMap, oDrawing) {
		const arrRes = [];
		let nAdditionalCounter = 0;
		let arrLetterSymbols = [];
		const arrBackgroundLetters = this.getBackgroundElementsByLetter();
		const arrForegroundLetters = this.getForegroundElementsByLetter();
		for (let i = 0; i < this.m_aWords.length; i += 1) {
			const aWord = this.m_aWords[i];
			for (let j = 0; j < aWord.length; j += 1) {
				const oWordLetter = aWord[j];
				arrLetterSymbols.push(oWordLetter);
				if (oWordLetter.isBulletSymbol) {
					nAdditionalCounter += 1;
					continue;
				}
				const oBackgroundLetter = arrBackgroundLetters[nAdditionalCounter];
				const oForegroundLetter = arrForegroundLetters[nAdditionalCounter];
				const oWrapperWord = new AscCommonSlide.CObjectForDrawArrayWrapper(arrLetterSymbols, oTransform, oTheme, oColorMap, oDrawing, oForegroundLetter, oBackgroundLetter);
				arrRes.push(oWrapperWord);
				nAdditionalCounter += 1;
				arrLetterSymbols = [];
			}
		}
		if (arrLetterSymbols.length) {
			const oWrapperWord = new AscCommonSlide.CObjectForDrawArrayWrapper(arrLetterSymbols, oTransform, oTheme, oColorMap, oDrawing);
			arrRes.push(oWrapperWord);
		}
		return arrRes;
	};
	CParagraphStructure.prototype.getAllDrawingObjectsWrapper = function (oTransform, oTheme, oColorMap, oDrawing) {
		let arrContentObjects = [];
		arrContentObjects = arrContentObjects.concat.apply(arrContentObjects, this.m_aWords);
		let arrBackgroundObjects = [];
		arrBackgroundObjects = arrBackgroundObjects.concat.apply(arrBackgroundObjects, this.getBackgroundElementsByAll());
		let arrForegroundObjects = [];
		arrForegroundObjects = arrForegroundObjects.concat.apply(arrForegroundObjects, this.getForegroundElementsByAll());
		return [new AscCommonSlide.CObjectForDrawArrayWrapper(arrContentObjects, oTransform, oTheme, oColorMap, oDrawing, arrForegroundObjects, arrBackgroundObjects)];
	};

	CParagraphStructure.prototype.Recalculate = function (oTheme, oColorMap, dWidth, dHeight, oShape) {
		var i;
		for (i = 0; i < this.m_aContent.length; ++i) {
			this.m_aContent[i].Recalculate(oTheme, oColorMap, dWidth, dHeight, oShape);
		}
	};

	CParagraphStructure.prototype.CheckContentStructs = function (aContentStructs) {
		var i;
		for (i = 0; i < this.m_aContent.length; ++i) {
			this.m_aContent[i].CheckContentStructs(aContentStructs);
		}
	};
	CParagraphStructure.prototype.draw = function (graphics, transform, oTheme, oColorMap) {
		var i;
		for (i = 0; i < this.m_aContent.length; ++i) {
			this.m_aContent[i].draw(graphics, transform, oTheme, oColorMap);
		}
	};

	CParagraphStructure.prototype.getAllBackgroundsBorders = function (aParaBackgrounds, aBackgrounds, aBorders, aComments) {
		var i;
		for (i = 0; i < this.m_aContent.length; ++i) {
			this.m_aContent[i].getAllBackgroundsBorders(aParaBackgrounds, aBackgrounds, aBorders, aComments);
		}
	};
	CParagraphStructure.prototype.getTextStructures = function () {
		const aTextStructures = [];
		for (let nContent = 0; nContent < this.m_aContent.length; ++nContent) {
			this.m_aContent[nContent].getTextStructures(aTextStructures);
		}
		return aTextStructures;
	};
	CParagraphStructure.prototype.checkWord = function (oSpaceElement, bCheckGeometry) {
		const oWordPos = this.m_oLastWordStart;
		let aWord = [];
		let nWordLineCount = -1;
		let oObjectToDraw;
		for (let nLine = oWordPos.line; nLine < this.m_aContent.length; ++nLine) {
			let oLine = this.m_aContent[nLine];
			let aContent = oLine.m_aContent;
			if (aContent.length === 0) {
				break;
			}
			if (oWordPos.line < nLine) {
				oWordPos.posInLine = -1;
			}
			for (let nPosInLine = oWordPos.posInLine + 1; nPosInLine < aContent.length; ++nPosInLine) {
				oObjectToDraw = aContent[nPosInLine];
				if (bCheckGeometry && !oObjectToDraw.geometry.isEmpty() || oObjectToDraw.TextElement) {
					aWord.push(oObjectToDraw);
					oWordPos.posInLine = nPosInLine;
				}
				oObjectToDraw.charIndex = oWordPos.charIndex++;
			}
			oWordPos.charIndex++;
			oWordPos.line = nLine;
			if (aWord.length) {
				nWordLineCount += 1;
			}
		}
		if (oObjectToDraw && oSpaceElement) {
			oObjectToDraw.addSpaceTextElement(oSpaceElement);
		}

		if (aWord.length > 0) {
			this.m_aWords.push(aWord);
		}
	};
    CParagraphStructure.prototype.Write_ToBinary = function (writer) {
        writer.WriteLong(this.m_aContent.length);
        for(let idx = 0; idx < this.m_aContent.length; ++idx) {
            this.m_aContent[idx].Write_ToBinary(writer);
        }
    };
    CParagraphStructure.prototype.Read_FromBinary = function (reader) {
        let count = reader.GetLong();
        for(let idx = 0; idx < count; ++idx) {
            let paragraph = new CLineStructure();
            paragraph.Read_FromBinary(reader);
            this.m_aContent.push(paragraph);
        }
    };
	function CShapeStructure(memory) {
		this.memory = memory;
		this.m_aContent = [];
		this.m_aBorders = [];
		this.m_nType = DRAW_COMMAND_SHAPE;
	}

	CShapeStructure.prototype.Recalculate = function (oTheme, oColorMap, dWidth, dHeight, oShape) {
		for (let i = 0; i < this.m_aContent.length; ++i) this.m_aContent[i].Recalculate(oTheme, oColorMap, dWidth, dHeight, oShape, true);

		for (let i = 0; i < this.m_aBorders.length; ++i) this.m_aBorders[i].Recalculate(oTheme, oColorMap, dWidth, dHeight, oShape, true);
	};
	CShapeStructure.prototype.CheckContentStructs = function (aContentStructs) {
		for (let i = 0; i < this.m_aContent.length; ++i) this.m_aContent[i].CheckContentStructs(aContentStructs);
	};

	CShapeStructure.prototype.getAllBackgroundsBorders = function (aParaBackgrounds, aBackgrounds, aBorders, aComments) {
		for (let i = 0; i < this.m_aBorders.length; ++i) aBorders.push(this.m_aBorders[i]);

		for (let i = 0; i < this.m_aContent.length; ++i) this.m_aContent[i].getAllBackgroundsBorders(aParaBackgrounds, aBackgrounds, aBorders, aComments);
	};

	CShapeStructure.prototype.draw = function (graphics, transform, oTheme, oColorMap) {
		for (let i = 0; i < this.m_aBorders.length; ++i) this.m_aBorders[i].draw(graphics, undefined, transform, oTheme, oColorMap);

		for (let i = 0; i < this.m_aContent.length; ++i) this.m_aContent[i].draw(graphics, transform, oTheme, oColorMap);
	}

	CShapeStructure.prototype.Write_ToBinary = function (writer) {
		this.memory.Write_ToBinary(writer);
		writer.bWriteCompiledColors = true;
		writer.WriteLong(this.m_aBorders.length);
		for (let i = 0; i < this.m_aBorders.length; ++i) this.m_aBorders[i].Write_ToBinary(writer);

		writer.WriteLong(this.m_aContent.length);
		for (let i = 0; i < this.m_aContent.length; ++i) this.m_aContent[i].Write_ToBinary(writer);

		writer.bWriteCompiledColors = undefined;
	};
	CShapeStructure.prototype.Read_FromBinary = function (reader) {
		this.memory = new AscFormat.CPathMemory();
		this.memory.Read_FromBinary(reader);
		reader.bReadCompiledColors = true;
		reader.pathMemory = this.memory;
		let count = reader.GetLong();
		for (let i = 0; i < count; ++i) {
			let border = new ObjectToDraw();
			border.Read_FromBinary(reader);
			this.m_aBorders.push(border);
		}

		count = reader.GetLong();
		for (let i = 0; i < count; ++i) {
			let content = new CDocContentStructure();
			content.Read_FromBinary(reader);
			this.m_aContent.push(content);
		}
		reader.pathMemory = undefined;
		reader.bReadCompiledColors = undefined;
	};

	function CTableStructure() {
		this.m_nType = DRAW_COMMAND_TABLE;
		this.m_aContent = [];
		this.m_aBorders = [];
	}

	CTableStructure.prototype.Recalculate = function (oTheme, oColorMap, dWidth, dHeight, oShape) {
		var i;
		for (i = 0; i < this.m_aContent.length; ++i) {
			this.m_aContent[i].Recalculate(oTheme, oColorMap, dWidth, dHeight, oShape, true);
		}

		for (i = 0; i < this.m_aBorders.length; ++i) {
			this.m_aBorders[i].Recalculate(oTheme, oColorMap, dWidth, dHeight, oShape, true);
		}
	};
	CTableStructure.prototype.CheckContentStructs = function (aContentStructs) {
		var i;
		for (i = 0; i < this.m_aContent.length; ++i) {
			this.m_aContent[i].CheckContentStructs(aContentStructs);
		}
	};
	CTableStructure.prototype.draw = function (graphics, transform, oTheme, oColorMap) {
		var i;
		for (i = 0; i < this.m_aContent.length; ++i) {
			this.m_aContent[i].draw(graphics, transform, oTheme, oColorMap);
		}
	};

	CTableStructure.prototype.getAllBackgroundsBorders = function (aParaBackgrounds, aBackgrounds, aBorders, aComments) {
		var i;
		for (i = 0; i < this.m_aBorders.length; ++i) {
			aBorders.push(this.m_aBorders[i]);
		}
		for (i = 0; i < this.m_aContent.length; ++i) {
			this.m_aContent[i].getAllBackgroundsBorders(aParaBackgrounds, aBackgrounds, aBorders, aComments);
		}
	};

	const CLineStructure_DrawType_Content = 0;
	const CLineStructure_DrawType_Borders = 1;
	const CLineStructure_DrawType_Backgrounds = 2;
	const CLineStructure_DrawType_Foregrounds = 3;
	const CLineStructure_DrawType_ParagraphBackgrounds = 4;
	const CLineStructure_DrawType_Highlights = 5;
	const CLineStructure_DrawType_Underlines = 6;
	const CLineStructure_DrawType_DUnderlines = 7;
	const CLineStructure_DrawType_Strikeouts = 8;
	const CLineStructure_DrawType_DStrikeouts = 9;

	function CLineStructure(oLine) {
		this.m_nType = DRAW_COMMAND_LINE;
		this.m_oLine = oLine;
		this.m_aContent = [];//ObjectToDraw
		this.m_aBorders = [];
		this.m_aBackgrounds = [];
		this.m_aHighlights = [];
		this.m_aUnderlines = [];
		this.m_aStrikeouts = [];
		this.m_aDStrikeouts = [];
		this.m_aDUnderlines = [];
		this.m_aForegrounds = [];
		this.m_aParagraphBackgrounds = [];
		this.m_nDrawType = CLineStructure_DrawType_Content;
	}

	CLineStructure.prototype.addGroundElement = function (oTextElement) {
		const arrContent = this.getContentByDrawingType();
		const oLastElement = arrContent[arrContent.length - 1];
		if (oLastElement) {
			oLastElement.TextElement = oTextElement;
		}
	}
	CLineStructure.prototype.isBackgroundDrawType = function () {
		switch (this.m_nDrawType) {
			case CLineStructure_DrawType_Backgrounds:
			case CLineStructure_DrawType_ParagraphBackgrounds:
			case CLineStructure_DrawType_Highlights:
				return true;
			default:
				return false;
		}
	};
	CLineStructure.prototype.isForegroundDrawType = function () {
		switch (this.m_nDrawType) {
			case CLineStructure_DrawType_Underlines:
			case CLineStructure_DrawType_DUnderlines:
			case CLineStructure_DrawType_Strikeouts:
			case CLineStructure_DrawType_DStrikeouts:
			case CLineStructure_DrawType_Foregrounds:
				return true;
			default:
				return false;
		}
	};
	CLineStructure.prototype.isContentDrawType = function () {
		switch (this.m_nDrawType) {
			case CLineStructure_DrawType_Content:
				return true;
			default:
				return false;
		}
	};

	CLineStructure.prototype.getContentByDrawingType = function (nDrawingType) {
		nDrawingType = typeof nDrawingType === 'number' ? nDrawingType : this.m_nDrawType;
		switch (nDrawingType) {
			case CLineStructure_DrawType_Content:
				return this.m_aContent;
			case CLineStructure_DrawType_Borders:
				return this.m_aBorders;
			case CLineStructure_DrawType_Backgrounds:
				return this.m_aBackgrounds;
			case CLineStructure_DrawType_Foregrounds:
				return this.m_aForegrounds;
			case CLineStructure_DrawType_ParagraphBackgrounds:
				return this.m_aParagraphBackgrounds;
			case CLineStructure_DrawType_Highlights:
				return this.m_aHighlights;
			case CLineStructure_DrawType_Underlines:
				return this.m_aUnderlines;
			case CLineStructure_DrawType_DUnderlines:
				return this.m_aDUnderlines;
			case CLineStructure_DrawType_Strikeouts:
				return this.m_aStrikeouts;
			case CLineStructure_DrawType_DStrikeouts:
				return this.m_aDStrikeouts;

			default:
				return [];
		}
	};

	CLineStructure.prototype.Recalculate = function (oTheme, oColorMap, dWidth, dHeight, oShape) {
		var i;
		for (i = 0; i < this.m_aContent.length; ++i) {
			this.m_aContent[i].Recalculate(oTheme, oColorMap, dWidth, dHeight, oShape, true);
		}
		for (i = 0; i < this.m_aBorders.length; ++i) {
			this.m_aBorders[i].Recalculate(oTheme, oColorMap, dWidth, dHeight, oShape, true);
		}
		for (i = 0; i < this.m_aBackgrounds.length; ++i) {
			this.m_aBackgrounds[i].Recalculate(oTheme, oColorMap, dWidth, dHeight, oShape, true);
		}
		for (i = 0; i < this.m_aHighlights.length; ++i) {
			this.m_aHighlights[i].Recalculate(oTheme, oColorMap, dWidth, dHeight, oShape, true);
		}
		for (i = 0; i < this.m_aForegrounds.length; ++i) {
			this.m_aForegrounds[i].Recalculate(oTheme, oColorMap, dWidth, dHeight, oShape, true);
		}
		for (i = 0; i < this.m_aUnderlines.length; ++i) {
			this.m_aUnderlines[i].Recalculate(oTheme, oColorMap, dWidth, dHeight, oShape, true);
		}
		for (i = 0; i < this.m_aDUnderlines.length; ++i) {
			this.m_aDUnderlines[i].Recalculate(oTheme, oColorMap, dWidth, dHeight, oShape, true);
		}
		for (i = 0; i < this.m_aStrikeouts.length; ++i) {
			this.m_aStrikeouts[i].Recalculate(oTheme, oColorMap, dWidth, dHeight, oShape, true);
		}
		for (i = 0; i < this.m_aDStrikeouts.length; ++i) {
			this.m_aDStrikeouts[i].Recalculate(oTheme, oColorMap, dWidth, dHeight, oShape, true);
		}
		for (i = 0; i < this.m_aParagraphBackgrounds.length; ++i) {
			this.m_aParagraphBackgrounds[i].Recalculate(oTheme, oColorMap, dWidth, dHeight, oShape, true);
		}
	};

	CLineStructure.prototype.CheckContentStructs = function (aContentStructs) {
		const isN = AscFormat.isRealNumber;
		for (let nElement = 0; nElement < this.m_aContent.length; ++nElement) {
			let oElement = this.m_aContent[nElement];
			if (oElement.TextElement) {
				aContentStructs.push(oElement);
			}
		}
	};

	CLineStructure.prototype.getTextStructures = function (aTextStructures) {
		return this.CheckContentStructs(aTextStructures);
	};
	CLineStructure.prototype.GetAllWarped = function (aWarpedObjects) {
		var i;
		for (i = 0; i < this.m_aContent.length; ++i) {
			aWarpedObjects.push(this.m_aContent[i]);
		}
		for (i = 0; i < this.m_aBackgrounds.length; ++i) {
			aWarpedObjects.push(this.m_aBackgrounds[i]);
		}
		for (i = 0; i < this.m_aHighlights.length; ++i) {
			aWarpedObjects.push(this.m_aHighlights[i]);
		}
		for (i = 0; i < this.m_aForegrounds.length; ++i) {
			aWarpedObjects.push(this.m_aForegrounds[i]);
		}
		for (i = 0; i < this.m_aUnderlines.length; ++i) {
			aWarpedObjects.push(this.m_aUnderlines[i]);
		}
		for (i = 0; i < this.m_aDUnderlines.length; ++i) {
			aWarpedObjects.push(this.m_aDUnderlines[i]);
		}
		for (i = 0; i < this.m_aStrikeouts.length; ++i) {
			aWarpedObjects.push(this.m_aStrikeouts[i]);
		}
		for (i = 0; i < this.m_aDStrikeouts.length; ++i) {
			aWarpedObjects.push(this.m_aDStrikeouts[i]);
		}
	};
	CLineStructure.prototype.CheckBoundsWarped = function (graphics) {
		var i;
		for (i = 0; i < this.m_aContent.length; ++i) {
			this.m_aContent[i].draw(graphics, true);
		}
		for (i = 0; i < this.m_aBackgrounds.length; ++i) {
			this.m_aBackgrounds[i].draw(graphics, true);
		}
		for (i = 0; i < this.m_aHighlights.length; ++i) {
			this.m_aHighlights[i].draw(graphics, true);
		}
		for (i = 0; i < this.m_aForegrounds.length; ++i) {
			this.m_aForegrounds[i].draw(graphics, true);
		}
		for (i = 0; i < this.m_aUnderlines.length; ++i) {
			this.m_aUnderlines[i].draw(graphics, true);
		}
		for (i = 0; i < this.m_aDUnderlines.length; ++i) {
			this.m_aDUnderlines[i].draw(graphics, true);
		}
		for (i = 0; i < this.m_aStrikeouts.length; ++i) {
			this.m_aStrikeouts[i].draw(graphics, true);
		}
		for (i = 0; i < this.m_aDStrikeouts.length; ++i) {
			this.m_aDStrikeouts[i].draw(graphics, true);
		}
	};
	CLineStructure.prototype.draw = function (graphics, transform, oTheme, oColorMap) {
		var i;
		// for(i = 0; i < this.m_aParagraphBackgrounds.length; ++i)
		// {
		//     this.m_aParagraphBackgrounds[i].draw(graphics, undefined, transform, oTheme, oColorMap);
		// }
		//  for(i = 0; i < this.m_aBorders.length; ++i)
		//  {
		//      this.m_aBorders[i].draw(graphics);
		//  }
		//   for(i = 0; i < this.m_aBackgrounds.length; ++i)
		//   {
		//       this.m_aBackgrounds[i].draw(graphics, undefined, transform, oTheme, oColorMap);
		//   }
		for (i = 0; i < this.m_aContent.length; ++i) {
			this.m_aContent[i].draw(graphics, undefined, transform, oTheme, oColorMap);
		}
		for (i = 0; i < this.m_aForegrounds.length; ++i) {
			this.m_aForegrounds[i].draw(graphics, undefined, transform, oTheme, oColorMap);
		}
		for (i = 0; i < this.m_aUnderlines.length; ++i) {
			this.m_aUnderlines[i].draw(graphics, undefined, transform, oTheme, oColorMap);
		}
		for (i = 0; i < this.m_aDUnderlines.length; ++i) {
			this.m_aDUnderlines[i].draw(graphics, undefined, transform, oTheme, oColorMap);
		}
		for (i = 0; i < this.m_aStrikeouts.length; ++i) {
			this.m_aStrikeouts[i].draw(graphics, undefined, transform, oTheme, oColorMap);
		}
		for (i = 0; i < this.m_aDStrikeouts.length; ++i) {
			this.m_aDStrikeouts[i].draw(graphics, undefined, transform, oTheme, oColorMap);
		}
	};

	CLineStructure.prototype.getAllBackgroundsBorders = function (aParaBackgrounds, aBackgrounds, aBorders, aComments) {
		var i;
		for (i = 0; i < this.m_aParagraphBackgrounds.length; ++i) {
			aParaBackgrounds.push(this.m_aParagraphBackgrounds[i]);
		}
		for (i = 0; i < this.m_aBackgrounds.length; ++i) {
			aBackgrounds.push(this.m_aBackgrounds[i]);
			if (this.m_aBackgrounds[i].Comment) {
				aComments.push(this.m_aBackgrounds[i])
			}
		}
		for (i = 0; i < this.m_aHighlights.length; ++i) {
			aBackgrounds.push(this.m_aHighlights[i]);
		}
		for (i = 0; i < this.m_aBorders.length; ++i) {
			aBorders.push(this.m_aBorders[i]);
		}
	};
    CLineStructure.prototype.Write_ToBinary = function (writer) {


        writer.WriteLong(this.m_aContent.length);
        for(let idx = 0; idx < this.m_aContent.length; ++idx) {
            this.m_aContent[idx].Write_ToBinary(writer);
        }
        writer.WriteLong(this.m_aBorders.length);
        for(let idx = 0; idx < this.m_aBorders.length; ++idx) {
            this.m_aBorders[idx].Write_ToBinary(writer);
        }
        writer.WriteLong(this.m_aBackgrounds.length);
        for(let idx = 0; idx < this.m_aBackgrounds.length; ++idx) {
            this.m_aBackgrounds[idx].Write_ToBinary(writer);
        }
        writer.WriteLong(this.m_aHighlights.length);
        for(let idx = 0; idx < this.m_aHighlights.length; ++idx) {
            this.m_aHighlights[idx].Write_ToBinary(writer);
        }
        writer.WriteLong(this.m_aUnderlines.length);
        for(let idx = 0; idx < this.m_aUnderlines.length; ++idx) {
            this.m_aUnderlines[idx].Write_ToBinary(writer);
        }
        writer.WriteLong(this.m_aStrikeouts.length);
        for(let idx = 0; idx < this.m_aStrikeouts.length; ++idx) {
            this.m_aStrikeouts[idx].Write_ToBinary(writer);
        }
        writer.WriteLong(this.m_aDStrikeouts.length);
        for(let idx = 0; idx < this.m_aDStrikeouts.length; ++idx) {
            this.m_aDStrikeouts[idx].Write_ToBinary(writer);
        }
        writer.WriteLong(this.m_aDUnderlines.length);
        for(let idx = 0; idx < this.m_aDUnderlines.length; ++idx) {
            this.m_aDUnderlines[idx].Write_ToBinary(writer);
        }
        writer.WriteLong(this.m_aForegrounds.length);
        for(let idx = 0; idx < this.m_aForegrounds.length; ++idx) {
            this.m_aForegrounds[idx].Write_ToBinary(writer);
        }
        writer.WriteLong(this.m_aParagraphBackgrounds.length);
        for(let idx = 0; idx < this.m_aParagraphBackgrounds.length; ++idx) {
            this.m_aParagraphBackgrounds[idx].Write_ToBinary(writer);
        }
    };
    CLineStructure.prototype.Read_FromBinary = function (reader) {
        let count = reader.GetLong();
        for(let idx = 0; idx < count; ++idx) {
            let paragraph = new ObjectToDraw();
            paragraph.Read_FromBinary(reader);
            this.m_aContent.push(paragraph);
        }
        count = reader.GetLong();
        for(let idx = 0; idx < count; ++idx) {
            let paragraph = new ObjectToDraw();
            paragraph.Read_FromBinary(reader);
            this.m_aBorders.push(paragraph);
        }
        count = reader.GetLong();
        for(let idx = 0; idx < count; ++idx) {
            let paragraph = new ObjectToDraw();
            paragraph.Read_FromBinary(reader);
            this.m_aBackgrounds.push(paragraph);
        }
        count = reader.GetLong();
        for(let idx = 0; idx < count; ++idx) {
            let paragraph = new ObjectToDraw();
            paragraph.Read_FromBinary(reader);
            this.m_aHighlights.push(paragraph);
        }
        count = reader.GetLong();
        for(let idx = 0; idx < count; ++idx) {
            let paragraph = new ObjectToDraw();
            paragraph.Read_FromBinary(reader);
            this.m_aUnderlines.push(paragraph);
        }
        count = reader.GetLong();
        for(let idx = 0; idx < count; ++idx) {
            let paragraph = new ObjectToDraw();
            paragraph.Read_FromBinary(reader);
            this.m_aStrikeouts.push(paragraph);
        }
        count = reader.GetLong();
        for(let idx = 0; idx < count; ++idx) {
            let paragraph = new ObjectToDraw();
            paragraph.Read_FromBinary(reader);
            this.m_aDStrikeouts.push(paragraph);
        }
        count = reader.GetLong();
        for(let idx = 0; idx < count; ++idx) {
            let paragraph = new ObjectToDraw();
            paragraph.Read_FromBinary(reader);
            this.m_aDUnderlines.push(paragraph);
        }
        count = reader.GetLong();
        for(let idx = 0; idx < count; ++idx) {
            let paragraph = new ObjectToDraw();
            paragraph.Read_FromBinary(reader);
            this.m_aForegrounds.push(paragraph);
        }
        count = reader.GetLong();
        for(let idx = 0; idx < count; ++idx) {
            let paragraph = new ObjectToDraw();
            paragraph.Read_FromBinary(reader);
            this.m_aParagraphBackgrounds.push(paragraph);
        }
    };

	var DRAW_COMMAND_TABLE = 0x01;
	var DRAW_COMMAND_CONTENT = 0x02;
	var DRAW_COMMAND_PARAGRAPH = 0x03;
	var DRAW_COMMAND_LINE = 0x04;
	var DRAW_COMMAND_DRAWING = 0x05;
	var DRAW_COMMAND_HIDDEN_ELEM = 0x06;
	var DRAW_COMMAND_NO_CREATE_GEOM = 0x07;
	var DRAW_COMMAND_TABLE_ROW = 0x08;
	var DRAW_COMMAND_SHAPE = 0x09;

	function GetConstDescription(nConst) {
		switch (nConst) {
			case DRAW_COMMAND_TABLE: {
				return "DRAW_COMMAND_TABLE";
			}
			case DRAW_COMMAND_CONTENT: {
				return "DRAW_COMMAND_CONTENT";
			}
			case DRAW_COMMAND_PARAGRAPH: {
				return "DRAW_COMMAND_PARAGRAPH";
			}
			case DRAW_COMMAND_LINE: {
				return "DRAW_COMMAND_LINE";
			}
			case DRAW_COMMAND_DRAWING: {
				return "DRAW_COMMAND_DRAWING";
			}
			case DRAW_COMMAND_HIDDEN_ELEM: {
				return "DRAW_COMMAND_HIDDEN_ELEM";
			}
			case DRAW_COMMAND_NO_CREATE_GEOM: {
				return "DRAW_COMMAND_NO_CREATE_GEOM";
			}
			case DRAW_COMMAND_TABLE_ROW: {
				return "DRAW_COMMAND_TABLE_ROW";
			}
			default : {
				return "Unknown";
			}
		}
	}

	function CreatePenFromParams(oUnifill, nStyle, nLineCap, nLineJoin, dLineWidth, dSize) {
		var oLine = new AscFormat.CLn();
		oLine.setW(dSize * 36000 >> 0);
		oLine.setFill(oUnifill);


		//TODO
		// this.Color      = { R : 255, G : 255, B : 255, A : 255 };
		// this.Style      = 0;
		// this.LineCap    = 0;
		// this.LineJoin   = 0;
//
		// this.LineWidth  = 1;

		return oLine;
	}

	function CIdGenerator() {
		this.id = 0;
	}

	CIdGenerator.prototype.getId = function () {
		return this.id++;
	};

	function CTextDrawer(dWidth, dHeight, bDivByLInes, oTheme, bDivGlyphs, oSplitOptions, oIdGenerator) {
		AscCommon.CGraphicsBase.call(this, AscCommon.RendererType.TextDrawer, true);

		this.m_oFont = {
			Name: "", FontSize: -1, Style: -1
		};

		this.m_oTheme = oTheme;
		this.Width = dWidth;
		this.Height = dHeight;
		this.m_oTransform = new AscCommon.CMatrix();
		this.pathW = 1000000000;
		this.pathH = 1000000000;

		this.xKoeff = this.pathW;
		this.yKoeff = this.pathH;

		this.m_aStack = [];
		this.m_oDocContentStructure = null;
		this.m_aCommands = [];
		this.m_aDrawings = [];
		// RFonts
		this.m_oTextPr = null;
		this.m_oGrFonts = new AscCommon.CGrRFonts();
		this.m_oCurComment = null;

		this.m_oPen = new AscCommon.CPen();
		this.m_oBrush = new AscCommon.CBrush();

		this.m_oLine = null;
		this.m_oFill = null;


		// чтобы выставилось в первый раз
		this.m_oPen.Color.R = -1;
		this.m_oBrush.Color1.R = -1;
		this.m_oBrush.Color2.R = -1;

		// просто чтобы не создавать каждый раз
		this.m_oFontSlotFont = new AscCommon.CFontSetup();
		this.LastFontOriginInfo = {Name: "", Replace: null};

		this.m_bDivByLines = bDivByLInes === true;
		this.m_bDivGlyphs = bDivGlyphs === true;
		this.m_aByLines = null;
		this.m_nCurLineIndex = -1;
		this.m_aStackLineIndex = [];
		this.m_aStackCurRowMaxIndex = null;
		this.m_aByParagraphs = null;

		this.m_oObjectToDraw = null;

		this.m_bTurnOff = false;

		this.bCheckLines = false;
		this.lastX = null;
		this.lastY = null;
		if (this.m_bDivByLines) {
			this.m_aByLines = [];
			this.m_aStackCurRowMaxIndex = [];
		}
		else {
			this.m_aByParagraphs = [];
		}

		this.m_bIsTextDrawer = true;
		this.m_oSplitOptions = oSplitOptions || {};
		this.m_nCurrentSplitOptions = null;
		this.oIdGenerator = oIdGenerator;
		this.m_oTextElementsWithId = {};
		this.pathMemory = new AscFormat.CPathMemory();
	}

	CTextDrawer.prototype = Object.create(AscCommon.CGraphicsBase.prototype);
	CTextDrawer.prototype.constructor = CTextDrawer;
	CTextDrawer.prototype.addTextElementId = function (oElement) {
		if (!oElement) {
			return;
		}
		if (this.oIdGenerator) {
			if (!this.m_oTextElementsWithId[oElement.textDrawerId]) {
				oElement.textDrawerId = this.oIdGenerator.getId();
				this.m_oTextElementsWithId[oElement.textDrawerId] = oElement;
			}
		}
	};
	CTextDrawer.prototype.clearTextElements = function () {
		for (let id in this.m_oTextElementsWithId) {
			delete this.m_oTextElementsWithId[id].textDrawerId;
		}
		this.m_oTextElementsWithId = {};
	};
	CTextDrawer.prototype.addGroundElement = function (oElement) {
		const oTextElement = oElement && oElement.Additional && oElement.Additional.TextDrawer && oElement.Additional.TextDrawer.TextElement;
		this.addTextElementId(oTextElement);
		for (let nPos = this.m_aStack.length - 1; nPos >= 0; nPos -= 1) {
			if (this.m_aStack[nPos] instanceof CLineStructure) {
				this.m_aStack[nPos].addGroundElement(oTextElement);
			}
		}
	}

	CTextDrawer.prototype.SetObjectToDraw = function (oObjectToDraw) {
		this.m_oObjectToDraw = oObjectToDraw;
	};
// pen methods
	CTextDrawer.prototype.p_color = function (r, g, b, a) {
		if (this.m_oPen.Color.R != r || this.m_oPen.Color.G != g || this.m_oPen.Color.B != b) {
			this.m_oPen.Color.R = r;
			this.m_oPen.Color.G = g;
			this.m_oPen.Color.B = b;
		}
		if (this.m_oPen.Color.A != a) {
			this.m_oPen.Color.A = a;
		}
		var oLastCommand = this.m_aStack[this.m_aStack.length - 1];
		if (oLastCommand) {
			if (oLastCommand.m_nType === DRAW_COMMAND_LINE && oLastCommand.isForegroundDrawType()) {
				if (this.m_oTextPr && this.m_oTheme) {
					var oTextPr = this.m_oTextPr.Copy();
					if (!oTextPr.TextOutline) {
						oTextPr.TextOutline = new AscFormat.CLn();
					}
					oTextPr.TextOutline.Fill = this.CreateUnfilFromRGB(r, g, b);
					this.SetTextPr(oTextPr, this.m_oTheme, this.m_nCurrentSplitOptions !== null);
					return;
				}
			}
		}
		this.Get_PathToDraw(false, true);
	};
	CTextDrawer.prototype.AddSmartRect = function () {
	};
	CTextDrawer.prototype.p_width = function (w) {
		var val = w / 1000;
		if (this.m_oPen.Size != val) {
			this.m_oPen.Size = val;
		}
		this.Get_PathToDraw(false, true);
	};

// brush methods
	CTextDrawer.prototype.b_color1 = function (r, g, b, a) {
		if (this.m_oBrush.Color1.R !== r || this.m_oBrush.Color1.G !== g || this.m_oBrush.Color1.B !== b) {
			this.m_oBrush.Color1.R = r;
			this.m_oBrush.Color1.G = g;
			this.m_oBrush.Color1.B = b;
		}
		if (this.m_oBrush.Color1.A !== a) {
			this.m_oBrush.Color1.A = a;
		}
		var oLastCommand = this.m_aStack[this.m_aStack.length - 1];
		if (oLastCommand) {
			if (oLastCommand.m_nType === DRAW_COMMAND_LINE && oLastCommand.isContentDrawType()) {
				if (this.m_oTextPr && this.m_oTheme) {
					var oTextPr = this.m_oTextPr.Copy();
					oTextPr.Unifill = undefined;
					oTextPr.TextFill = undefined;
					oTextPr.FontRef = undefined;
					oTextPr.Color = new CDocumentColor(r, g, b, false);
					this.SetTextPr(oTextPr, this.m_oTheme, this.m_nCurrentSplitOptions !== null);
					return;
				}
			}
		}
		this.Get_PathToDraw(false, true);
	};

	CTextDrawer.prototype.set_fillColor = function (R, G, B) {
		this.m_oFill = AscFormat.CreateUniFillByUniColor(AscFormat.CreateUniColorRGB(R, G, B));
		this.Get_PathToDraw(false, true);
	};

	CTextDrawer.prototype.b_color2 = function (r, g, b, a) {
		if (this.m_oBrush.Color2.R != r || this.m_oBrush.Color2.G != g || this.m_oBrush.Color2.B != b) {
			this.m_oBrush.Color2.R = r;
			this.m_oBrush.Color2.G = g;
			this.m_oBrush.Color2.B = b;
		}
		if (this.m_oBrush.Color2.A != a) {
			this.m_oBrush.Color2.A = a;
		}
		this.Get_PathToDraw(false, true);
	};

	CTextDrawer.prototype.put_brushTexture = function (src, mode) {
		this.m_oBrush.Color1.R = -1;
		this.m_oBrush.Color1.G = -1;
		this.m_oBrush.Color1.B = -1;
		this.m_oBrush.Color1.A = -1;
		this.Get_PathToDraw(false, true);
	};

	CTextDrawer.prototype.SetShd = function (oShd) {
		if (oShd) {
			if (oShd.Value !== Asc.c_oAscShdNil) {
				if (oShd.Unifill) {
					this.m_oFill = oShd.Unifill;
				}
				else {
					if (oShd.Color) {
						this.m_oFill = this.CreateUnfilFromRGB(oShd.Color.r, oShd.Color.g, oShd.Color.b);
					}
					else {
						this.m_oFill = null;
					}
				}
			}
			else {
				this.m_oFill = null;
			}
		}
		else {
			this.m_oFill = null;
		}
		this.Get_PathToDraw(false, true);
	};

	CTextDrawer.prototype.SetBorder = function (oBorder) {
		if (oBorder && oBorder.Value !== border_None) {
			this.m_oLine = CreatePenFromParams(oBorder.Unifill ? oBorder.Unifill : this.CreateUnfilFromRGB(oBorder.Color.r, oBorder.Color.g, oBorder.Color.b), this.m_oPen.Style, this.m_oPen.LineCap, this.m_oPen.LineJoin, this.m_oPen.LineWidth, this.m_oPen.Size);
		}
		else {
			this.m_oLine = null;
		}
		//TODO
	};

	CTextDrawer.prototype.SetAdditionalProps = function (oProps) {
		oProps && this.SetTextPr(oProps, this.m_oTheme);
	};

	CTextDrawer.prototype.StartCheckTableDraw = function () {
		this.Start_Command(DRAW_COMMAND_TABLE);
	};

	CTextDrawer.prototype.EndCheckTableDraw = function () {
		this.End_Command();
	};
	CTextDrawer.prototype.setCurrentSplitOptions = function (nIndex) {
		this.m_nCurrentSplitOptions = typeof this.m_oSplitOptions[nIndex] === 'number' ? this.m_oSplitOptions[nIndex] : null;
	};
	CTextDrawer.prototype.Start_Command = function (commandId, param, index, nType) {
		this.m_aCommands.push(commandId);
		let oNewStructure = null;
		let bOld = false;
		switch (commandId) {
			case DRAW_COMMAND_NO_CREATE_GEOM: {
				break;
			}
			case DRAW_COMMAND_CONTENT: {
				oNewStructure = new CDocContentStructure();
				this.m_aStackLineIndex.push(this.m_nCurLineIndex);
				break;
			}
			case DRAW_COMMAND_PARAGRAPH: {

				if (param) {
					let oLast = this.m_aStack[this.m_aStack.length - 1];
					if (oLast) {
						let aLastContent = oLast.m_aContent;
						if (Array.isArray(aLastContent)) {
							for (let nIdx = 0; nIdx < aLastContent.length; ++nIdx) {
								let oElement = aLastContent[nIdx];
								if (oElement.m_oParagraph === param) {
									oNewStructure = oElement;
									bOld = true;
									this.setCurrentSplitOptions(nIdx);
									break;
								}
							}
						}
					}
				}
				if (!oNewStructure) {
					oNewStructure = new CParagraphStructure(param);
					const oDocStructure = this.m_aStack[this.m_aStack.length - 1];
					if (oDocStructure) {
						const oLastParagraph = oDocStructure.m_aContent[oDocStructure.m_aContent.length - 1];
						if (oLastParagraph && oLastParagraph.m_oLastWordStart) {
							oNewStructure.m_oLastWordStart.charIndex = oLastParagraph.m_oLastWordStart.charIndex;
						}
						this.setCurrentSplitOptions(oDocStructure.m_aContent.length);

					}
					if (!this.m_bDivByLines) {
						this.m_aByParagraphs[this.m_aByParagraphs.length] = [];
					}
				}
				break;
			}
			case DRAW_COMMAND_LINE: {
				var oPrevStruct = this.m_aStack[this.m_aStack.length - 1];
				if (oPrevStruct.m_nType === DRAW_COMMAND_PARAGRAPH && oPrevStruct.m_aContent[index]) {
					oPrevStruct.m_aContent[index].m_nDrawType = nType;
					this.m_aStack.push(oPrevStruct.m_aContent[index]);
				}
				else {
					oNewStructure = new CLineStructure(param);
					oNewStructure.m_nDrawType = nType;
					if (this.m_bDivByLines) {
						++this.m_nCurLineIndex;
						if (!Array.isArray(this.m_aByLines[this.m_nCurLineIndex])) {
							this.m_aByLines[this.m_nCurLineIndex] = [];
						}
						this.m_aByLines[this.m_nCurLineIndex].push(oNewStructure);
						if (this.m_aStackCurRowMaxIndex[this.m_aStackCurRowMaxIndex.length - 1] < this.m_nCurLineIndex) {
							this.m_aStackCurRowMaxIndex[this.m_aStackCurRowMaxIndex.length - 1] = this.m_nCurLineIndex;
						}
					}
					else {
						this.m_aByParagraphs[this.m_aByParagraphs.length - 1].push(oNewStructure);
					}
				}
				break;
			}
			case DRAW_COMMAND_TABLE: {
				oNewStructure = new CTableStructure();
				break;
			}
			case DRAW_COMMAND_DRAWING: {
				break;
			}
			case DRAW_COMMAND_HIDDEN_ELEM: {
				break;
			}
			case DRAW_COMMAND_TABLE_ROW: {
				this.m_aStackCurRowMaxIndex[this.m_aStackCurRowMaxIndex.length] = -1;
				break;
			}
			case DRAW_COMMAND_SHAPE: {
				oNewStructure = new CShapeStructure(this.pathMemory);
				break;
			}
		}
		if (oNewStructure) {
			if (this.m_aStack[this.m_aStack.length - 1] && !bOld) {
				this.m_aStack[this.m_aStack.length - 1].m_aContent.push(oNewStructure);
			}
			this.m_aStack.push(oNewStructure);
		}
	};

	CTextDrawer.prototype.End_Command = function () {
		var nCommandId = this.m_aCommands.pop();
		switch (nCommandId) {
			case DRAW_COMMAND_NO_CREATE_GEOM: {
				break;
			}
			case DRAW_COMMAND_CONTENT: {
				var oDocContentStructure = this.m_aStack.pop();
				if (this.m_aStack.length === 0) {
					this.m_oDocContentStructure = oDocContentStructure;
					this.m_oDocContentStructure.m_aByLines = this.m_aByLines;
					this.m_oDocContentStructure.m_aByParagraphs = this.m_aByParagraphs;
					this.m_oDocContentStructure.m_aDrawingsStruct = this.m_aDrawings;
					this.m_aByLines = [];
					this.m_aByParagraphs = [];
					this.m_aDrawings = [];
				}
				this.m_nCurLineIndex = this.m_aStackLineIndex.pop();
				break;
			}
			case DRAW_COMMAND_PARAGRAPH: {
				let oParaStruct = this.m_aStack.pop();
				if (oParaStruct && oParaStruct.checkWord) {
					oParaStruct.checkWord();
				}
				break;
			}
			case DRAW_COMMAND_LINE: {
				this.m_aStack.pop();
				break;
			}
			case DRAW_COMMAND_TABLE: {
				this.m_aStack.pop();
				break;
			}
			case DRAW_COMMAND_DRAWING: {
				break;
			}
			case DRAW_COMMAND_HIDDEN_ELEM: {
				break;
			}
			case DRAW_COMMAND_TABLE_ROW: {
				this.m_nCurLineIndex = this.m_aStackCurRowMaxIndex.pop();
				break;
			}
			case DRAW_COMMAND_SHAPE: {
				this.m_aStack.pop();
				break;
			}
		}
	};

	CTextDrawer.prototype.Get_PathToDraw = function (bStart, bStart2, x, y, oTextElement, bIsBulletSymbol) {
		var oPath = null;
		var oLastCommand = this.m_aStack[this.m_aStack.length - 1];
		var oLastObjectToDraw;
		if (oLastCommand) {
			switch (oLastCommand.m_nType) {
				case DRAW_COMMAND_LINE: {
					const arrContent = oLastCommand.getContentByDrawingType();
					switch (oLastCommand.m_nDrawType) {
						case CLineStructure_DrawType_Content: {
							if (arrContent.length === 0) {
								arrContent.push(new ObjectToDraw(this.GetFillFromTextPr(this.m_oTextPr), this.GetPenFromTextPr(this.m_oTextPr), this.Width, this.Height, new Geometry(), this.m_oTransform, x, y, this.m_oCurComment, oTextElement, oLastCommand, this.oIdGenerator && this.oIdGenerator.getId(), bIsBulletSymbol));
							}
							oLastObjectToDraw = arrContent[arrContent.length - 1];

							if (bStart2) {
								if (oLastObjectToDraw.geometry.isEmpty()) {
									oLastObjectToDraw.resetBrushPen(this.GetFillFromTextPr(this.m_oTextPr), this.GetPenFromTextPr(this.m_oTextPr), x, y, oTextElement, bIsBulletSymbol)
								}
								else {
									arrContent.push(new ObjectToDraw(this.GetFillFromTextPr(this.m_oTextPr), this.GetPenFromTextPr(this.m_oTextPr), this.Width, this.Height, new Geometry(), this.m_oTransform, x, y, null, oTextElement, oLastCommand, this.oIdGenerator && this.oIdGenerator.getId(), bIsBulletSymbol));
									oLastObjectToDraw = arrContent[arrContent.length - 1];
								}
							}
							break;
						}
						case CLineStructure_DrawType_Borders:
						case CLineStructure_DrawType_Backgrounds:
						case CLineStructure_DrawType_ParagraphBackgrounds:
						case CLineStructure_DrawType_Highlights: {
							if (arrContent.length === 0) {
								arrContent.push(new ObjectToDraw(this.m_oFill, this.m_oLine, this.Width, this.Height, new Geometry(), this.m_oTransform, x, y, null, oTextElement, oLastCommand, this.oIdGenerator && this.oIdGenerator.getId()))
							}
							oLastObjectToDraw = arrContent[arrContent.length - 1];

							if (bStart2) {
								if (oLastObjectToDraw.geometry.isEmpty()) {
									oLastObjectToDraw.resetBrushPen(this.m_oFill, this.m_oLine, x, y);
								}
								else {
									arrContent.push(new ObjectToDraw(this.m_oFill, this.m_oLine, this.Width, this.Height, new Geometry(), this.m_oTransform, x, y, null, oTextElement, oLastCommand, this.oIdGenerator && this.oIdGenerator.getId()));
									oLastObjectToDraw = arrContent[arrContent.length - 1];
								}
							}
							break;
						}
						case CLineStructure_DrawType_Foregrounds:
						case CLineStructure_DrawType_Underlines:
						case CLineStructure_DrawType_DUnderlines:
						case CLineStructure_DrawType_DStrikeouts:
						case CLineStructure_DrawType_Strikeouts: {
							if (arrContent.length === 0) {
								arrContent.push(new ObjectToDraw(this.GetFillFromTextPr(this.m_oTextPr), this.GetPenFromTextPr(this.m_oTextPr), this.Width, this.Height, new Geometry(), this.m_oTransform, x, y, null, oTextElement, oLastCommand, this.oIdGenerator && this.oIdGenerator.getId()))
							}
							oLastObjectToDraw = arrContent[arrContent.length - 1];

							if (bStart2) {
								if (oLastObjectToDraw.geometry.isEmpty()) {
									oLastObjectToDraw.resetBrushPen(this.GetFillFromTextPr(this.m_oTextPr), this.GetPenFromTextPr(this.m_oTextPr), x, y);
								}
								else {
									arrContent.push(new ObjectToDraw(this.GetFillFromTextPr(this.m_oTextPr), this.GetPenFromTextPr(this.m_oTextPr), this.Width, this.Height, new Geometry(), this.m_oTransform, x, y, null, oTextElement, oLastCommand, this.oIdGenerator && this.oIdGenerator.getId()));
									oLastObjectToDraw = arrContent[arrContent.length - 1];
								}
							}
							break;
						}
					}
					break;
				}
				case DRAW_COMMAND_TABLE: {
					if (oLastCommand.m_aBorders.length === 0 || bStart2) {
						oLastCommand.m_aBorders.push(new ObjectToDraw(this.m_oFill, this.m_oLine, this.Width, this.Height, new Geometry(), this.m_oTransform, x, y));
					}
					oLastObjectToDraw = oLastCommand.m_aBorders[oLastCommand.m_aBorders.length - 1];

					if (bStart2) {
						if (oLastObjectToDraw.geometry.isEmpty()) {
							oLastObjectToDraw.resetBrushPen(this.m_oFill, this.m_oLine, x, y);
						}
						else {
							oLastCommand.m_aBorders.push(new ObjectToDraw(this.m_oFill, this.m_oLine, this.Width, this.Height, new Geometry(), this.m_oTransform, x, y));
							oLastObjectToDraw = oLastCommand.m_aBorders[oLastCommand.m_aBorders.length - 1];
						}
					}
					oLastObjectToDraw.geometry.bDrawSmart = true;
					break;
				}
				case DRAW_COMMAND_SHAPE: {
					if (oLastCommand.m_aBorders.length === 0 || bStart2) {
						oLastCommand.m_aBorders.push(new ObjectToDraw(this.m_oFill, this.m_oLine, this.Width, this.Height, new Geometry(), this.m_oTransform, x, y));
					}
					oLastObjectToDraw = oLastCommand.m_aBorders[oLastCommand.m_aBorders.length - 1];

					if (bStart2) {
						if (oLastObjectToDraw.geometry.isEmpty()) {
							oLastObjectToDraw.resetBrushPen(this.m_oFill, this.m_oLine, x, y);
						}
						else {
							oLastCommand.m_aBorders.push(new ObjectToDraw(this.m_oFill, this.m_oLine, this.Width, this.Height, new Geometry(), this.m_oTransform, x, y));
							oLastObjectToDraw = oLastCommand.m_aBorders[oLastCommand.m_aBorders.length - 1];
						}
					}
					break;
				}
				case DRAW_COMMAND_PARAGRAPH: {
					break;
				}
			}
		}
		else {
			if (this.m_oObjectToDraw) {
				oLastObjectToDraw = this.m_oObjectToDraw;
			}
		}

		if (oLastObjectToDraw && oLastObjectToDraw.geometry) {
			oLastObjectToDraw.Comment = this.m_oCurComment;
			if (oLastObjectToDraw.geometry.pathLst.length === 0 || bStart) {
				oPath = this.pathMemory.AllocPath2();
				oPath.setPathW(this.pathW);
				oPath.setPathH(this.pathH);
				oPath.setExtrusionOk(false);
				oPath.setFill("none");
				oPath.setStroke(true);
				oLastObjectToDraw.geometry.AddPath(oPath)
			}
			else {
				oPath = oLastObjectToDraw.geometry.pathLst[oLastObjectToDraw.geometry.pathLst.length - 1];
			}
		}
		return oPath;
	};

	CTextDrawer.prototype.transform = function (sx, shy, shx, sy, tx, ty) {
		if (this.m_oTransform.sx != sx || this.m_oTransform.shx != shx || this.m_oTransform.shy != shy || this.m_oTransform.sy != sy || this.m_oTransform.tx != tx || this.m_oTransform.ty != ty) {
			this.m_oTransform.sx = sx;
			this.m_oTransform.shx = shx;
			this.m_oTransform.shy = shy;
			this.m_oTransform.sy = sy;
			this.m_oTransform.tx = tx;
			this.m_oTransform.ty = ty;
		}
	};
// path commands
	CTextDrawer.prototype._s = function () {
		if (this.m_bTurnOff) return;
		this.Get_PathToDraw(true);
	};
	CTextDrawer.prototype._e = function () {
	};
	CTextDrawer.prototype._z = function () {
		var oPathToDraw = this.Get_PathToDraw();
		if (oPathToDraw) {
			oPathToDraw.close();
		}
	}
	CTextDrawer.prototype._m = function (x, y) {
		if (this.m_bTurnOff) return;
		var oPathToDraw = this.Get_PathToDraw();

		let tr = this.GetTransform();
		let bUseTr = this.isStampAnnot;

		let _x = bUseTr ? tr.TransformPointX(x, y) : x;
		let _y = bUseTr ? tr.TransformPointY(x, y) : y;

		if (oPathToDraw) {
			oPathToDraw.moveTo(this.xKoeff * _x, this.yKoeff * _y);
		}
		this.lastX = x;
		this.lastY = y;
	};
	CTextDrawer.prototype._l = function (x, y) {
		if (this.m_bTurnOff) return;

		let tr = this.GetTransform();
		let bUseTr = this.isStampAnnot;

		let _x = bUseTr ? tr.TransformPointX(x, y) : x;
		let _y = bUseTr ? tr.TransformPointY(x, y) : y;

		if (this.bCheckLines) {
			if (Math.abs(x - this.lastX) < EPSILON_TEXT_AUTOFIT && Math.abs(x - this.lastX) < Math.abs(y - this.lastY)) {
				this.checkCurveBezier(this.lastX, this.lastY, this.lastX, this.lastY + (y - this.lastY) / 3, this.lastX, this.lastY + 2 * (y - this.lastY) / 3, x, y, PATH_DIV_EPSILON);
				this.lastX = x;
				this.lastY = y;
				return;
			}
			else if (Math.abs(y - this.lastY) < EPSILON_TEXT_AUTOFIT && Math.abs(y - this.lastY) < Math.abs(x - this.lastX)) {
				this.checkCurveBezier(this.lastX, this.lastY, this.lastX + (x - this.lastX) / 3, this.lastY, this.lastX + 2 * (x - this.lastX) / 3, this.lastY, x, y, PATH_DIV_EPSILON);
				this.lastX = x;
				this.lastY = y;
				return;
			}
		}
		var oPathToDraw = this.Get_PathToDraw();
		if (oPathToDraw) {
			oPathToDraw.lnTo(this.xKoeff * _x, this.yKoeff * _y);
		}
		this.lastX = x;
		this.lastY = y;
	};
	CTextDrawer.prototype._c = function (x1, y1, x2, y2, x3, y3) {
		let tr = this.GetTransform();
		let bUseTr = this.isStampAnnot;

		let _x1 = bUseTr ? tr.TransformPointX(x1, y1) : x1;
		let _y1 = bUseTr ? tr.TransformPointY(x1, y1) : y1;
		let _x2 = bUseTr ? tr.TransformPointX(x2, y2) : x2;
		let _y2 = bUseTr ? tr.TransformPointY(x2, y2) : y2;
		let _x3 = bUseTr ? tr.TransformPointX(x3, y3) : x3;
		let _y3 = bUseTr ? tr.TransformPointY(x3, y3) : y3;

		var oPathToDraw = this.Get_PathToDraw();
		if (oPathToDraw) {
			oPathToDraw.cubicBezTo(this.xKoeff * _x1, this.yKoeff * _y1, this.xKoeff * _x2, this.yKoeff * _y2, this.xKoeff * _x3, this.yKoeff * _y3);
		}
		this.lastX = x3;
		this.lastY = y3;
	};
	CTextDrawer.prototype._c2 = function (x1, y1, x2, y2) {
		let tr = this.GetTransform();
		let bUseTr = this.isStampAnnot;

		let _x1 = bUseTr ? tr.TransformPointX(x1, y1) : x1;
		let _y1 = bUseTr ? tr.TransformPointY(x1, y1) : y1;
		let _x2 = bUseTr ? tr.TransformPointX(x2, y2) : x2;
		let _y2 = bUseTr ? tr.TransformPointY(x2, y2) : y2;

		var oPathToDraw = this.Get_PathToDraw();
		if (oPathToDraw) {
			oPathToDraw.quadBezTo(this.xKoeff * _x1, this.yKoeff * _y1, this.xKoeff * _x2, this.yKoeff * _y2);
		}
		this.lastX = x2;
		this.lastY = y2;
	};
	CTextDrawer.prototype.ds = function () {
	};
	CTextDrawer.prototype.df = function () {
		var oPathToDraw = this.Get_PathToDraw();
		if (oPathToDraw) {
			oPathToDraw.setFill("norm");
		}
	};

	CTextDrawer.prototype.drawpath = function (type) {
	};

	CTextDrawer.prototype.SetFont = function (font) {
		if (null == font) return;

		var style = 0;
		if (font.Italic == true) style += 2;
		if (font.Bold == true) style += 1;

		var fontinfo = g_fontApplication.GetFontInfo(font.FontFamily.Name, style, this.LastFontOriginInfo);
		style = fontinfo.GetBaseStyle(style);

		if (this.m_oFont.Name != fontinfo.Name) {
			this.m_oFont.Name = fontinfo.Name;
		}
		if (this.m_oFont.FontSize != font.FontSize) {
			this.m_oFont.FontSize = font.FontSize;
		}
		if (this.m_oFont.Style != style) {
			this.m_oFont.Style = style;
		}
	};
	CTextDrawer.prototype.FillText = function (x, y, text) {
		this.FillTextCode(x, y, text.charCodeAt(0));
	};

	CTextDrawer.prototype.CheckAddNewPath = function (x, y, oTextElement, isBulletSymbol) {

		if (this.m_bDivGlyphs === true) {
			this.addTextElementId(oTextElement);
			this.Get_PathToDraw(false, true, x, y, oTextElement, isBulletSymbol);
		}
	};

	CTextDrawer.prototype.CheckSpaceDraw = function (oSpaceElement, bCheckGeometry) {
		this.addTextElementId(oSpaceElement);
		for (let nPos = 0; nPos < this.m_aStack.length; ++nPos) {
			if (this.m_aStack[nPos] instanceof CParagraphStructure) {
				this.m_aStack[nPos].checkWord(oSpaceElement, bCheckGeometry);
			}
		}
	};

	CTextDrawer.prototype.FillTextCode = function (x, y, code) {
		this.CheckAddNewPath(x, y);
		AscCommon.g_oTextMeasurer.SetFontInternal(this.m_oFont.Name, this.m_oFont.FontSize, Math.max(this.m_oFont.Style, 0));

		if (null != this.LastFontOriginInfo.Replace) {
			code = g_fontApplication.GetReplaceGlyph(code, this.LastFontOriginInfo.Replace);
		}

		if (AscCommon.g_oTextMeasurer.m_oManager) AscCommon.g_oTextMeasurer.m_oManager.LoadStringPathCode(code, false, x, y, this);
	};
	CTextDrawer.prototype.tg = function (gid, x, y) {
		g_oTextMeasurer.SetFontInternal(this.m_oFont.Name, this.m_oFont.FontSize, Math.max(this.m_oFont.Style, 0));
		g_oTextMeasurer.m_oManager.LoadStringPathCode(gid, true, x, y, this);
	};

	CTextDrawer.prototype.SetFontInternal = function (name, size, style) {
		var fontinfo = g_fontApplication.GetFontInfo(name, style);
		if (this.m_oFont.Name !== fontinfo.Name) {
			this.m_oFont.Name = fontinfo.Name;
		}
		if (this.m_oFont.FontSize !== size) {
			this.m_oFont.FontSize = size;
		}
		if (this.m_oFont.Style != style) {
			this.m_oFont.Style = style;
		}
	};

	CTextDrawer.prototype.SetFontSlot = function (slot, fontSizeKoef) {
		var _rfonts = this.m_oGrFonts;
		var _lastFont = this.m_oFontSlotFont;

		switch (slot) {
			case fontslot_ASCII: {
				_lastFont.Name = _rfonts.Ascii.Name;
				_lastFont.Size = this.m_oTextPr.FontSize;
				_lastFont.Bold = this.m_oTextPr.Bold;
				_lastFont.Italic = this.m_oTextPr.Italic;

				break;
			}
			case fontslot_CS: {
				_lastFont.Name = _rfonts.CS.Name;
				_lastFont.Size = this.m_oTextPr.FontSizeCS;
				_lastFont.Bold = this.m_oTextPr.BoldCS;
				_lastFont.Italic = this.m_oTextPr.ItalicCS;

				break;
			}
			case fontslot_EastAsia: {
				_lastFont.Name = _rfonts.EastAsia.Name;
				_lastFont.Size = this.m_oTextPr.FontSize;
				_lastFont.Bold = this.m_oTextPr.Bold;
				_lastFont.Italic = this.m_oTextPr.Italic;

				break;
			}
			case fontslot_HAnsi:
			default: {
				_lastFont.Name = _rfonts.HAnsi.Name;
				_lastFont.Size = this.m_oTextPr.FontSize;
				_lastFont.Bold = this.m_oTextPr.Bold;
				_lastFont.Italic = this.m_oTextPr.Italic;

				break;
			}
		}
		if (undefined !== fontSizeKoef) _lastFont.Size *= fontSizeKoef;

		var style = 0;
		if (_lastFont.Italic == true) style += 2;
		if (_lastFont.Bold == true) style += 1;

		var fontinfo = g_fontApplication.GetFontInfo(_lastFont.Name, style, this.LastFontOriginInfo);

		if (this.m_oFont.Name != fontinfo.Name) {
			this.m_oFont.Name = fontinfo.Name;
		}
		if (this.m_oFont.FontSize != _lastFont.Size) {
			this.m_oFont.FontSize = _lastFont.Size;
		}
		if (this.m_oFont.Style != style) {
			this.m_oFont.Style = style;
		}
	};

	CTextDrawer.prototype.transform3 = function (m) {
		this.transform(m.sx, m.shy, m.shx, m.sy, m.tx, m.ty);
	};
	CTextDrawer.prototype.reset = function () {
		this.transform(1, 0, 0, 1, 0, 0);
	};

	CTextDrawer.prototype.FillText2 = function (x, y, text) {
		this.FillText(x, y, text);
	};

	CTextDrawer.prototype.GetFont = function () {
		return this.m_oFont;
	};
	CTextDrawer.prototype.put_GlobalAlpha = function (enable, alpha) {
	};

	CTextDrawer.prototype.checkCurveBezier = function (x0, y0, x1, y1, x2, y2, x3, y3, dEpsilon) {
		var _epsilon = dEpsilon ? dEpsilon : UNDERLINE_DIV_EPSILON;
		var arr_point = AscFormat.partition_bezier4(x0, y0, x1, y1, x2, y2, x3, y3, _epsilon), i,
			count = arr_point.length >> 2;
		for (i = 0; i < count; ++i) {
			var k = 4 * i;
			this._c(arr_point[k + 1].x, arr_point[k + 1].y, arr_point[k + 2].x, arr_point[k + 2].y, arr_point[k + 3].x, arr_point[k + 3].y);
		}
	};
// smart methods for horizontal / vertical lines
	CTextDrawer.prototype.drawHorLine = function (align, y, x, r, penW, bMath) {
		if (bMath) {
			var oTextPr = this.GetTextPr();
			var oCopyTextPr;
			if (oTextPr) {
				oCopyTextPr = oTextPr.Copy();
				oCopyTextPr.TextOutline = new AscFormat.CLn();
				oCopyTextPr.TextOutline.w = 36000.0 * penW >> 0;
				var oUnifill = oCopyTextPr.TextFill ? oCopyTextPr.TextFill : oCopyTextPr.Unifill;
				if ((!oUnifill || !oUnifill.fill || !oUnifill.fill.type === Asc.c_oAscFill._SOLID || !oUnifill.fill.color) && oCopyTextPr.Color) {
					oUnifill = AscFormat.CreateUniFillByUniColor(AscFormat.CreateUniColorRGB(oCopyTextPr.Color.r, oCopyTextPr.Color.g, oCopyTextPr.Color.b))
				}
				if (oUnifill) {
					oCopyTextPr.TextOutline.Fill = oUnifill;
				}
				this.SetTextPr(oCopyTextPr, this.m_oTheme);
			}
			this._s();
			this._m(x, y);
			this._l(r, y);
			this.ds();
			if (oCopyTextPr) {
				this.SetTextPr(oTextPr, this.m_oTheme);
			}
			return;
		}

		this._s();
		this._m(x, y);

		this.checkCurveBezier(x, y, x + ((r - x) / 3), y, x + (2 / 3) * (r - x), y, r, y);// this._l(r, y);
		this._l(r, y + penW);
		this.checkCurveBezier(r, y + penW, x + (2 / 3) * (r - x), y + penW, x + ((r - x) / 3), y + penW, x, y + penW);//this._l(x, y + penW);
		this._z();
		this.ds();
		this.df();
	};

	CTextDrawer.prototype.drawHorLine2 = function (align, y, x, r, penW) {
		var _y = y;
		switch (align) {
			case 0: {
				_y = y + penW / 2;
				break;
			}
			case 1: {
				break;
			}
			case 2: {
				_y = y - penW / 2;
				break;
			}
		}
		// if(!AdditionalData)
		// {
		//     this.p_width(1000 * penW);
//
//
		//     this._s();
		//     this._m(x, (_y - penW));
		//     this._l(r, (_y - penW));
		//     this.ds();
//
		//     this._s();
		//     this._m(x, (_y + penW));
		//     this._l(r, (_y + penW));
		//     this.ds();
//
		//     this._e();
		// }
		// else
		{

			this._s();
			this._m(x, (_y - penW));

			this.checkCurveBezier(x, _y - penW, x + ((r - x) / 3), _y - penW, x + (2 / 3) * (r - x), _y - penW, r, _y - penW);//this._l(r, (_y - penW ));
			this._l(r, _y);
			this.checkCurveBezier(r, _y, x + (2 / 3) * (r - x), _y, x + ((r - x) / 3), _y, x, _y);//this._l(x, (_y - penW + penW));
			this._z();
			this.ds();
			this.df();

			this._s();
			this._m(x, (_y + penW));
			this.checkCurveBezier(x, _y + penW, x + ((r - x) / 3), _y + penW, x + (2 / 3) * (r - x), _y + penW, r, _y + penW);//this._l(r, (_y + penW ));
			this._l(r, (_y + penW + penW));
			this.checkCurveBezier(r, (_y + penW + penW), x + (2 / 3) * (r - x), _y + penW + penW, x + ((r - x) / 3), _y + penW + penW, x, _y + penW + penW);//this._l(x, (_y + penW + penW));
			this._z();
			this.ds();

			this._e();
		}
	};

	CTextDrawer.prototype.drawVerLine = function (align, x, y, b, penW, bMath) {
		if (bMath) {
			var oTextPr = this.GetTextPr();
			var oCopyTextPr;
			if (oTextPr) {
				oCopyTextPr = oTextPr.Copy();
				oCopyTextPr.TextOutline = new AscFormat.CLn();
				oCopyTextPr.TextOutline.w = 36000.0 * penW >> 0;
				var oUnifill = oCopyTextPr.TextFill ? oCopyTextPr.TextFill : oCopyTextPr.Unifill;
				if ((!oUnifill || !oUnifill.fill || !oUnifill.fill.type === Asc.c_oAscFill.FILL_TYPE_SOLID || !oUnifill.fill.color) && oCopyTextPr.Color) {
					oUnifill = AscFormat.CreateUniFillByUniColor(AscFormat.CreateUniColorRGB(oCopyTextPr.Color.r, oCopyTextPr.Color.g, oCopyTextPr.Color.b))
				}
				if (oUnifill) {
					oCopyTextPr.TextOutline.Fill = oUnifill;
				}
				this.SetTextPr(oCopyTextPr, this.m_oTheme);
			}
			this._s();

			var _x = x;
			switch (align) {
				case 0: {
					_x = x + penW / 2;
					break;
				}
				case 1: {
					break;
				}
				case 2: {
					_x = x - penW / 2;
				}
			}
			this._m(_x, y);
			this._l(_x, b);

			this.ds();
			if (oCopyTextPr) {
				this.SetTextPr(oTextPr, this.m_oTheme);
			}
			return;
		}

		var nLastCommand = this.m_aCommands[this.m_aCommands.length - 1], bOldVal;
		if (nLastCommand === DRAW_COMMAND_TABLE) {
			bOldVal = this.bCheckLines;
			this.bCheckLines = false;
		}
		this.p_width(1000 * penW);
		this._s();

		var _x = x;
		switch (align) {
			case 0: {
				_x = x + penW / 2;
				break;
			}
			case 1: {
				break;
			}
			case 2: {
				_x = x - penW / 2;
			}
		}
		this._m(_x, y);
		this._l(_x, b);

		this.ds();
		if (nLastCommand === DRAW_COMMAND_TABLE) {
			this.bCheckLines = bOldVal;
		}
	};

	CTextDrawer.prototype.drawHorLineExt = function (align, y, x, r, penW, leftMW, rightMW) {
		var nLastCommand = this.m_aCommands[this.m_aCommands.length - 1];
		if (nLastCommand === DRAW_COMMAND_TABLE) {
			var bOldVal = this.bCheckLines;
			this.bCheckLines = false;
			this.p_width(penW * 1000);
			this._s();
			this._m(x, y);
			this._l(r, y);
			this.ds();
			this.bCheckLines = bOldVal;
		}
		else {
			this.drawHorLine(align, y, x + leftMW, r + rightMW, penW);
		}
	};

	CTextDrawer.prototype.DrawTextArtComment = function (Element) {
		this.m_oCurComment = Element;
		this.rect(Element.x0, Element.y0, Element.x1 - Element.x0, Element.y1 - Element.y0);
		this.df();
		this.m_oCurComment = null;
	};

	CTextDrawer.prototype.rect = function (x, y, w, h) {
		if (this.m_bTurnOff) return;
		var oLastCommand = this.m_aStack[this.m_aStack.length - 1];
		if (oLastCommand && oLastCommand.m_nType === DRAW_COMMAND_LINE && oLastCommand.isBackgroundDrawType()) {
			this.Get_PathToDraw(true, true);
			this._m(x, y);

			this.checkCurveBezier(x, y, x + (w / 3), y, x + (2 / 3) * w, y, x + w, y);// this._l(r, y);
			this._l(x + w, y + h);
			this.checkCurveBezier(x + w, y + h, x + (2 / 3) * (w), y + h, x + (w / 3), y + h, x, y + h);//this._l(x, y + penW);
			this._z();
			this.ds();
		}
		else {
			var _x = x;
			var _y = y;
			var _r = (x + w);
			var _b = (y + h);

			this.Get_PathToDraw(true, true);
			this._m(_x, _y);
			this._l(_r, _y);
			this._l(_r, _b);
			this._l(_x, _b);
			this._l(_x, _y);
		}
	};

	CTextDrawer.prototype.TableRect = function (x, y, w, h) {
		this.rect(x, y, w, h);
		this.df();
	};

	CTextDrawer.prototype.GetTransform = function () {
		return this.m_oTransform;
	};
	CTextDrawer.prototype.GetLineWidth = function () {
		return 0;
	};
	CTextDrawer.prototype.GetPen = function () {
		return 0;
	};
	CTextDrawer.prototype.GetBrush = function () {
		return 0;
	};

	CTextDrawer.prototype.StartClipPath = function () {
		this.m_bTurnOff = true;
	};

	CTextDrawer.prototype.EndClipPath = function () {
		this.m_bTurnOff = false;
	};

	CTextDrawer.prototype.SetTextPr = function (textPr, theme, bForceGetPath) {
		if (theme && textPr && textPr.ReplaceThemeFonts) textPr.ReplaceThemeFonts(theme.themeElements.fontScheme);

		var bNeedGetPath = !!bForceGetPath;
		if (!this.CheckCompareFillBrush(textPr, this.m_oTextPr)) {
			bNeedGetPath = true;
		}
		this.m_oTextPr = textPr;
		if (bNeedGetPath) {
			this.Get_PathToDraw(false, true);
		}
		if (theme) this.m_oGrFonts.checkFromTheme(theme.themeElements.fontScheme, this.m_oTextPr.RFonts); else this.m_oGrFonts = this.m_oTextPr.RFonts;
	};

	CTextDrawer.prototype.CheckCompareFillBrush = function (oTextPr1, oTextPr2) {
		if (!oTextPr1 && oTextPr2 || oTextPr1 && !oTextPr2) return false;
		if (oTextPr1 && oTextPr2) {
			var oFill1 = this.GetFillFromTextPr(oTextPr1);
			var oFill2 = this.GetFillFromTextPr(oTextPr2);
			if (!CompareBrushes(oFill1, oFill2)) return false;
			var oPen1 = this.GetPenFromTextPr(oTextPr1);
			var oPen2 = this.GetPenFromTextPr(oTextPr2);
			if (!CompareBrushes(oPen1, oPen2)) return false;
		}
		return true;
	};

	CTextDrawer.prototype.GetFillFromTextPr = function (oTextPr) {
		if (oTextPr) {
			if (oTextPr.TextFill) {
				return oTextPr.TextFill;
			}
			if (oTextPr.Unifill) {
				return oTextPr.Unifill;
			}
			if (oTextPr.Color) {
				var oColor = oTextPr.Color;
				if (oColor.Auto && oTextPr.FontRef && oTextPr.FontRef.Color && this.m_oTheme) {
					var oColorMap = AscFormat.GetDefaultColorMap();
					var oApi = window && window.editor;
					if (oApi) {
						var oDoc = oApi.WordControl && oApi.WordControl.m_oLogicDocument;
						if (oDoc && oDoc.Get_ColorMap) {
							var _cm = oDoc.Get_ColorMap();
							if (_cm) {
								oColorMap = _cm;
							}
						}
					}
					oTextPr.FontRef.Color.check(this.m_oTheme, oColorMap);
					var RGBA = oTextPr.FontRef.Color.RGBA;
					oColor = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B, RGBA.A);
				}
				return this.CreateUnfilFromRGB(oColor.r, oColor.g, oColor.b);
			}
			return null;
		}
		else {
			if (this.m_oBrush.Color1.R !== -1) {
				var Color = this.m_oBrush.Color1;
				return this.CreateUnfilFromRGB(Color.R, Color.G, Color.B);
			}
			else {
				return this.CreateUnfilFromRGB(0, 0, 0);
			}
		}
	};

	CTextDrawer.prototype.CreateUnfilFromRGB = function (r, g, b) {
		let oFill = AscFormat.CreateUnfilFromRGB(r, g, b);
		oFill.check(this.m_oTheme, AscFormat.GetDefaultColorMap());
		return oFill;
	};

	CTextDrawer.prototype.GetPenFromTextPr = function (oTextPr) {
		if (oTextPr) {
			return oTextPr.TextOutline;
		}
		return null;
	};

	CTextDrawer.prototype.GetTextPr = function () {
		return this.m_oTextPr;
	};
	CTextDrawer.prototype.isSupportTextOutline = function () {
		return true;
	};

	function PolygonWrapper(oPolygon) {
		this.oPolygon = oPolygon;
		var dCurLen = 0;
		this.aLength = [];
		this.aLength[0] = 0;
		var oPrevPoint = oPolygon[0], oCurPoint, dDX, dDY;
		for (var i = 1; i < oPolygon.length; ++i) {
			oCurPoint = oPolygon[i];
			dDX = oCurPoint.x - oPrevPoint.x;
			dDY = oCurPoint.y - oPrevPoint.y;
			dCurLen += Math.sqrt(dDX * dDX + dDY * dDY);
			this.aLength[i] = dCurLen;
			oPrevPoint = oCurPoint;
		}
		this.dLen = dCurLen;
		this.nPointsCount = this.aLength.length;
	}

	PolygonWrapper.prototype.getPointOnPolygon = function (dCT, bNeedPoints) {
		var dFindLen = this.dLen * dCT;
		var nIndex = this.nPointsCount >> 1;
		var nStartIndex = 0, nDelta = nIndex - nStartIndex, dNextBool, nTempIndex;
		nTempIndex = nIndex + 1;
		dNextBool = (nTempIndex < this.nPointsCount) && (this.aLength[nTempIndex] <= dFindLen);
		while (nDelta > 0 && (this.aLength[nIndex] > dFindLen || dNextBool)) {
			if (dNextBool) {
				nStartIndex = nIndex;
			}
			nIndex = nStartIndex + (nDelta >> 1);
			nTempIndex = nIndex + 1;
			dNextBool = (nTempIndex < this.nPointsCount) && (this.aLength[nTempIndex] <= dFindLen);
			nDelta = nIndex - nStartIndex;
		}
		if (nTempIndex === this.nPointsCount) {
			--nTempIndex;
			--nIndex;
		}
		var t;
		var dDiv = this.aLength[nTempIndex] - this.aLength[nIndex];
		if (dDiv !== 0) {
			t = (dFindLen - this.aLength[nIndex]) / dDiv;
		}
		else {
			t = 0;
		}

		var oPoint1 = this.oPolygon[nIndex], oPoint2 = this.oPolygon[nTempIndex];
		var oRetPoint1 = oPoint1, oRetPoint2 = oPoint2;
		if (bNeedPoints) {
			var nRightIndex = nTempIndex, nLeftIndex = nIndex;
			var oLeftPoint = oPoint1, oRightPoint = oPoint2;
			var dx = oPoint1.x - oPoint2.x, dy = oPoint1.y - oPoint2.y;
			while (nRightIndex + 1 < this.oPolygon.length && Math.abs(dx) < EPSILON_TEXT_AUTOFIT && Math.abs(dy) < EPSILON_TEXT_AUTOFIT) {

				dx = oRightPoint.x - oLeftPoint.x;
				dy = oRightPoint.y - oLeftPoint.y;
				oRightPoint = this.oPolygon[++nRightIndex]
			}
			while (nLeftIndex > 0 && Math.abs(dx) < EPSILON_TEXT_AUTOFIT && Math.abs(dy) < EPSILON_TEXT_AUTOFIT) {
				dx = oRightPoint.x - oLeftPoint.x;
				dy = oRightPoint.y - oLeftPoint.y;
				oLeftPoint = this.oPolygon[--nLeftIndex];
			}
			if (Math.abs(dx) < EPSILON_TEXT_AUTOFIT && Math.abs(dy) < EPSILON_TEXT_AUTOFIT) {
				oRetPoint1 = {x: oRetPoint1.x + EPSILON_TEXT_AUTOFIT, y: oRetPoint1.y};
			}
			else {
				oRetPoint1 = oLeftPoint;
				oRetPoint2 = oRightPoint;
			}
		}
		return {
			x: oPoint1.x + t * (oPoint2.x - oPoint1.x),
			y: oPoint1.y + t * (oPoint2.y - oPoint1.y),
			oP1: oRetPoint1,
			oP2: oRetPoint2
		};
	};


	function ObjectsToDrawBetweenTwoPolygons(aObjectsToDraw, oBoundsController, oPolygonWrapper1, oPolygonWrapper2) {
		var i;
		for (i = 0; i < aObjectsToDraw.length; ++i) {
			aObjectsToDraw[i].geometry.checkBetweenPolygons(oBoundsController, oPolygonWrapper1, oPolygonWrapper2);
		}
	}

	function CompareBrushes(oFill1, oFill2) {
		if (oFill1 && !oFill2 || !oFill1 && oFill2 || (oFill1 && oFill2 && !oFill1.IsIdentical(oFill2))) return false;
		return true;
	}

	function ComparePens(oPen1, oPen2) {
		if (oPen1 && !oPen2 || !oPen1 && oPen2 || (oPen1 && oPen2 && !oPen1.IsIdentical(oPen2))) return false;
		return true;
	}


	function GetRectContentWidth(oContent, dMaxWidth) {
		var _maxWidth = AscFormat.isRealNumber(dMaxWidth) ? dMaxWidth : 100000;

		oContent.Reset(0, 0, _maxWidth, 100000);
		oContent.Recalculate_Page(0, true);
		var max_width = 0;
		for (var i = 0; i < oContent.Content.length; ++i) {
			var par = oContent.Content[i];

			if (par instanceof Paragraph) {
				for (var j = 0; j < par.Lines.length; ++j) {
					if (par.Lines[j].Ranges[0].W > max_width) {
						max_width = par.Lines[j].Ranges[0].W;
					}
				}
			}
		}
		return max_width + 2;
	}

	//--------------------------------------------------------export----------------------------------------------------
	window['AscFormat'] = window['AscFormat'] || {};
	window['AscFormat'].PATH_DIV_EPSILON = PATH_DIV_EPSILON;
	window['AscFormat'].ParaDrawingStruct = ParaDrawingStruct;
	window['AscFormat'].DRAW_COMMAND_TABLE = DRAW_COMMAND_TABLE;
	window['AscFormat'].DRAW_COMMAND_CONTENT = DRAW_COMMAND_CONTENT;
	window['AscFormat'].DRAW_COMMAND_PARAGRAPH = DRAW_COMMAND_PARAGRAPH;
	window['AscFormat'].DRAW_COMMAND_LINE = DRAW_COMMAND_LINE;
	window['AscFormat'].DRAW_COMMAND_DRAWING = DRAW_COMMAND_DRAWING;
	window['AscFormat'].DRAW_COMMAND_HIDDEN_ELEM = DRAW_COMMAND_HIDDEN_ELEM;
	window['AscFormat'].DRAW_COMMAND_NO_CREATE_GEOM = DRAW_COMMAND_NO_CREATE_GEOM;
	window['AscFormat'].DRAW_COMMAND_TABLE_ROW = DRAW_COMMAND_TABLE_ROW;
	window['AscFormat'].DRAW_COMMAND_SHAPE = DRAW_COMMAND_SHAPE;
	window['AscFormat'].CreatePenFromParams = CreatePenFromParams;
	window['AscFormat'].CTextDrawer = CTextDrawer;
	window['AscFormat'].PolygonWrapper = PolygonWrapper;
	window['AscFormat'].GetRectContentWidth = GetRectContentWidth;
	window['AscFormat'].CIdGenerator = CIdGenerator;
	window['AscFormat'].CShapeStructure = CShapeStructure;

	window['AscFormat'].CLineStructure_DrawType_Content = CLineStructure_DrawType_Content;
	window['AscFormat'].CLineStructure_DrawType_Borders = CLineStructure_DrawType_Borders;
	window['AscFormat'].CLineStructure_DrawType_Backgrounds = CLineStructure_DrawType_Backgrounds;
	window['AscFormat'].CLineStructure_DrawType_Foregrounds = CLineStructure_DrawType_Foregrounds;
	window['AscFormat'].CLineStructure_DrawType_ParagraphBackgrounds = CLineStructure_DrawType_ParagraphBackgrounds;
	window['AscFormat'].CLineStructure_DrawType_Highlights = CLineStructure_DrawType_Highlights;
	window['AscFormat'].CLineStructure_DrawType_Underlines = CLineStructure_DrawType_Underlines;
	window['AscFormat'].CLineStructure_DrawType_DUnderlines = CLineStructure_DrawType_DUnderlines;
	window['AscFormat'].CLineStructure_DrawType_Strikeouts = CLineStructure_DrawType_Strikeouts;
	window['AscFormat'].CLineStructure_DrawType_DStrikeouts = CLineStructure_DrawType_DStrikeouts;
})(window);
