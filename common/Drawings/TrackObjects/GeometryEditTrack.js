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
(function(window, undefined) {

    var PathType = {
        POINT: 0,
        LINE: 1,
        ARC: 2,
        BEZIER_3: 3,
        BEZIER_4: 4,
        END: 5
    };

    function EditShapeGeometryTrack(originalObject, drawingObjects, extX, extY) {
        AscFormat.ExecuteNoHistory(function() {
            this.originalObject = originalObject;
            this.drawingObjects = drawingObjects;
            this.originalGeometry = this.getOriginalObjectGeometry();
            this.geometry = this.originalGeometry.createDuplicate();
            this.geometry.parent = originalObject.spPr.geometry.parent;
            this.originalShape = originalObject;
            this.shapeWidth = extX ? extX : originalObject.extX;
            this.shapeHeight = extY ? extY : originalObject.extY;
            var oPen = originalObject.pen;
            var oBrush = originalObject.brush;
            this.transform = originalObject.transform;
            this.invertTransform = originalObject.invertTransform;
            this.overlayObject = new AscFormat.OverlayObject(this.geometry, this.shapeWidth, this.shapeHeight, oBrush, oPen, this.transform);
            this.xMin = 0;
            this.yMin = 0;
            this.xMax = this.shapeWidth;
            this.yMax = this.shapeHeight;
            this.geometry.Recalculate(this.shapeWidth, this.shapeHeight);
            this.arrPathCommandsType = [];
            this.convertToBezier();
            this.createGeometryEditList();
            this.originalX = originalObject.x;
            this.originalY = originalObject.y;
            this.originalExtX = originalObject.extX;
            this.originalExtY = originalObject.extY;
            this.originalRot = originalObject.rot;

            var oPen1 = new AscFormat.CLn();
            oPen1.w = 15000;
            oPen1.Fill = AscFormat.CreateSolidFillRGBA(255, 255, 255, 255);
            var oPen2 = new AscFormat.CLn();
            oPen2.w = 15000;
            oPen2.Fill = AscFormat.CreateSolidFillRGBA(0x61, 0x9e, 0xde, 255);
            oPen2.prstDash = 0;

            this.pen1 = oPen1;
            this.pen2 = oPen2;

            this.overlayGeometry = this.geometry.createDuplicate();
            this.overlayObjectTrack = new AscFormat.OverlayObject(this.overlayGeometry, this.shapeWidth, this.shapeHeight, null, oPen1, this.transform);

            this.addedPointIdx = null;

        }, this, []);
    }
    EditShapeGeometryTrack.prototype.getOriginalObjectGeometry = function() {
        if(Asc.editor.isPdfEditor() && this.originalObject instanceof AscPDF.CAnnotationPolygon) {
            return this.originalObject.GetGeometryEdit();
        }
        return this.originalObject.spPr.geometry;
    };
    EditShapeGeometryTrack.prototype.draw = function(overlay)
    {
        var gmEditPoint = this.getGmEditPt();
        if(gmEditPoint)
        {
            var dOldAlpha = null;
            var oGraphics = overlay.Graphics ? overlay.Graphics : overlay;


            if(AscFormat.isRealNumber(this.originalShape.selectStartPage) && overlay.SetCurrentPage)
            {
                overlay.SetCurrentPage(this.originalShape.selectStartPage);
            }

            if(AscFormat.isRealNumber(oGraphics.globalAlpha) && oGraphics.put_GlobalAlpha)
            {
                dOldAlpha = oGraphics.globalAlpha;
                oGraphics.put_GlobalAlpha(false, 1);
            }
            if(overlay.DrawGeomEditPoint && !Asc.editor.isPdfEditor())
            {
                overlay.DrawGeomEditPoint(this.transform, gmEditPoint);
            }
            this.overlayObjectTrack.pen = this.pen1;
            this.overlayObjectTrack.draw(overlay, this.transform);
            this.overlayObjectTrack.pen = this.pen2;
            this.overlayObjectTrack.draw(overlay, this.transform);
            if(AscFormat.isRealNumber(dOldAlpha) && oGraphics.put_GlobalAlpha)
            {
                oGraphics.put_GlobalAlpha(true, dOldAlpha);
            }
        }
        //Check correct

        //this.overlayObject.TransformMatrix = this.originalShape.transform;
        //this.overlayObject.draw(overlay);
    };


    EditShapeGeometryTrack.prototype.drawSelect = function(oDrawingDocument, nGmEditPointIdx) {
        if(!this.isCorrect()) {
            return;
        }
        var geometry = this.geometry;
        var gmEditList = this.gmEditList;
        var gmEditPoint = this.getGmEditPt();
        var pathLst = geometry.pathLst;
        var matrix = this.transform;
        var oBounds = this.getBounds();
        oBounds.min_x -= 5;
        oBounds.min_y -= 5;
        oBounds.max_x += 5;
        oBounds.max_y += 5;
        if(Asc.editor.isPdfEditor()) {
            gmEditPoint = null;
        }
        oDrawingDocument.AutoShapesTrack.DrawGeometryEdit(matrix, pathLst, gmEditList, gmEditPoint, oBounds);
    };


    EditShapeGeometryTrack.prototype.getGmSelection = function() {
        return this.drawingObjects.selection.geometrySelection;
    };

    EditShapeGeometryTrack.prototype.getGmEditPtIdx = function() {
        var oGeomSelection = this.getGmSelection();
        if(!oGeomSelection) {
            return null;
        }
        return oGeomSelection.getGmEditPtIdx();
    };

    EditShapeGeometryTrack.prototype.getGmEditPt = function() {
        if(!this.isCorrect()) {
            return null;
        }
        if(this.addedPointIdx !== null) {
            return this.gmEditList[this.addedPointIdx];
        }
        var oPt = this.gmEditList[this.getGmEditPtIdx()];
        return oPt || null;
    };

    EditShapeGeometryTrack.prototype.getOriginalPt = function() {
        if(!this.isCorrect()) {
            return null;
        }
        if(this.originalAddedPoint) {
            return this.originalAddedPoint;
        }
        var oGeomSelection = this.drawingObjects.selection.geometrySelection;
        if(!oGeomSelection) {
            return null;
        }
        return oGeomSelection.getTrack().getGmEditPt();
    };

    EditShapeGeometryTrack.prototype.track = function(e, posX, posY) {
        AscFormat.ExecuteNoHistory(function(){
            var geometry = this.geometry;
            var gmEditPoint = this.getGmEditPt();
            if(!gmEditPoint) {
                return;
            }
            geometry.gdLstInfo = [];
            geometry.cnxLstInfo = [];

            for(var i = 0; i < geometry.pathLst.length; i++) {
                geometry.pathLst[i].ArrPathCommandInfo = [];
            }
            var invert_transform = this.originalShape.invertTransform;
            var _relative_x = invert_transform.TransformPointX(posX, posY);
            var _relative_y = invert_transform.TransformPointY(posX, posY);
            var originalPoint = this.getOriginalPt();
            var nextPoint = gmEditPoint.nextPoint;
            var prevPoint = gmEditPoint.prevPoint;
            var currentPath = geometry.pathLst[gmEditPoint.pathIndex];
            var arrPathCommand = currentPath.ArrPathCommand;
            this.gmEditPtIdx = this.drawingObjects.selection.geometrySelection.getGmEditPtIdx();

            var cur_command_type_array = this.arrPathCommandsType[gmEditPoint.pathIndex];
            var cur_command_type_1 = cur_command_type_array[gmEditPoint.pathC1];
            var cur_command_type_2 = cur_command_type_array[gmEditPoint.pathC2];

            if(gmEditPoint.isHitInFirstCPoint) {
                arrPathCommand[gmEditPoint.pathC1].X1 = _relative_x;
                arrPathCommand[gmEditPoint.pathC1].Y1 = _relative_y;

                if(cur_command_type_1 === PathType.LINE) {
                    cur_command_type_array[gmEditPoint.pathC1] = PathType.BEZIER_4;
                }

                var g1X = gmEditPoint.g1X;
                var g1Y = gmEditPoint.g1Y;

                gmEditPoint.g1X = _relative_x;
                gmEditPoint.g1Y = _relative_y;

                if(cur_command_type_1 === PathType.ARC && cur_command_type_2 === PathType.ARC) {
                    var isEllipseArc = Math.abs((gmEditPoint.X - g1X) * (gmEditPoint.g2Y - g1Y) - (gmEditPoint.g2X - g1X) * (gmEditPoint.Y - g1Y));
                    if(Math.floor(isEllipseArc) === 0) {
                        var g2X = gmEditPoint.g1X - gmEditPoint.X;
                        var g2Y = gmEditPoint.g1Y - gmEditPoint.Y;

                        arrPathCommand[gmEditPoint.pathC2].X0 = gmEditPoint.X - g2X;
                        arrPathCommand[gmEditPoint.pathC2].Y0 = gmEditPoint.Y - g2Y;
                        gmEditPoint.g2X = gmEditPoint.X - g2X;
                        gmEditPoint.g2Y = gmEditPoint.Y - g2Y;
                    } else {
                        var pathElemType =  this.arrPathCommandsType[gmEditPoint.pathIndex];
                        pathElemType[gmEditPoint.pathC1] = PathType.BEZIER_4;
                        pathElemType[gmEditPoint.pathC2] = PathType.BEZIER_4;
                    }
                }
            } else if(gmEditPoint.isHitInSecondCPoint) {
                arrPathCommand[gmEditPoint.pathC2].X0 = _relative_x;
                arrPathCommand[gmEditPoint.pathC2].Y0 = _relative_y;

                if(cur_command_type_2 === PathType.LINE) {
                    cur_command_type_array[gmEditPoint.pathC2] = PathType.BEZIER_4;
                }

                var g2X = gmEditPoint.g2X;
                var g2Y = gmEditPoint.g2Y;

                gmEditPoint.g2X = _relative_x;
                gmEditPoint.g2Y = _relative_y;

                if(cur_command_type_1 === PathType.ARC && cur_command_type_2 === PathType.ARC) {


                    var isEllipseArc = Math.abs((gmEditPoint.X - gmEditPoint.g1X) * (g2Y - gmEditPoint.g1Y) - (g2X - gmEditPoint.g1X) * (gmEditPoint.Y - gmEditPoint.g1Y));

                    if(Math.floor(isEllipseArc) === 0) {
                        var g1X = gmEditPoint.g2X - gmEditPoint.X;
                        var g1Y = gmEditPoint.g2Y - gmEditPoint.Y;

                        arrPathCommand[gmEditPoint.pathC1].X1 = gmEditPoint.X - g1X;
                        arrPathCommand[gmEditPoint.pathC1].Y1 = gmEditPoint.Y - g1Y;
                        gmEditPoint.g1X = gmEditPoint.X - g1X;
                        gmEditPoint.g1Y = gmEditPoint.Y - g1Y;
                    } else {
                        var pathElemType =  this.arrPathCommandsType[gmEditPoint.pathIndex];
                        pathElemType[gmEditPoint.pathC1] = PathType.BEZIER_4;
                        pathElemType[gmEditPoint.pathC2] = PathType.BEZIER_4;
                    }
                }
            } else {
                var X0, X1, Y0, Y1;
                //second curve relative to point
                var pathCommand = arrPathCommand[gmEditPoint.pathC1];
                var isPathCommand = (gmEditPoint.g1X !== undefined && gmEditPoint.g1Y !== undefined);

                X0 = cur_command_type_1 === PathType.LINE ? (prevPoint.X + _relative_x / 2) / (3 / 2) : prevPoint.g2X;
                Y0 = cur_command_type_1 === PathType.LINE ? (prevPoint.Y + _relative_y / 2) / (3 / 2) : prevPoint.g2Y;

                if (isPathCommand) {
                    X1 = cur_command_type_1 === PathType.LINE ? (prevPoint.X + _relative_x * 2) / 3 : _relative_x - originalPoint.X + originalPoint.g1X;
                    Y1 = cur_command_type_1 === PathType.LINE ? (prevPoint.Y + _relative_y * 2) / 3 : _relative_y - originalPoint.Y + originalPoint.g1Y;
                }

                var id = pathCommand.id === PathType.POINT ? PathType.POINT : PathType.BEZIER_4;
                var command = {
                    id: id,
                    X0: X0,
                    Y0: Y0,
                    X1: X1,
                    Y1: Y1,
                    X2: _relative_x,
                    Y2: _relative_y,
                    X: _relative_x,
                    Y: _relative_y
                };

                if(gmEditPoint.g1X !== undefined && gmEditPoint.g1Y !== undefined) {
                    gmEditPoint.g1X = command.X1;
                    gmEditPoint.g1Y = command.Y1;
                }

                gmEditPoint.X = _relative_x;
                gmEditPoint.Y = _relative_y;

                if(gmEditPoint.pathC1 && gmEditPoint.pathC2 && (gmEditPoint.pathC1 > gmEditPoint.pathC2)) {
                    arrPathCommand[gmEditPoint.pathC2 - 1].X = _relative_x;
                    arrPathCommand[gmEditPoint.pathC2 - 1].Y = _relative_y;
                }
                arrPathCommand[gmEditPoint.pathC1] = command;

                //second curve relative to point
                if(gmEditPoint.pathC2) {
                    var X0, X1, Y0, Y1;
                    var pathCommand = arrPathCommand[gmEditPoint.pathC2];
                    var isPathCommand = (gmEditPoint.g2X !== undefined && gmEditPoint.g2Y !== undefined);

                    if (isPathCommand) {
                        X0 = cur_command_type_2 === PathType.LINE ? (nextPoint.X + _relative_x * 2) / 3 : _relative_x - originalPoint.X + originalPoint.g2X;
                        Y0 = cur_command_type_2 === PathType.LINE ? (nextPoint.Y + _relative_y * 2) / 3 : _relative_y - originalPoint.Y + originalPoint.g2Y;
                    }

                    X1 = cur_command_type_2 === PathType.LINE ? (nextPoint.X + _relative_x / 2) / (3 / 2) : nextPoint.g1X;
                    Y1 = cur_command_type_2 === PathType.LINE ? (nextPoint.Y + _relative_y / 2) / (3 / 2) : nextPoint.g1Y;

                    var id = pathCommand.id === PathType.POINT ? PathType.POINT : PathType.BEZIER_4;
                    var command = {
                        id: id,
                        X0: X0,
                        Y0: Y0,
                        X1: X1,
                        Y1: Y1,
                        X2: nextPoint.X,
                        Y2: nextPoint.Y,
                        X: nextPoint.X,
                        Y: nextPoint.Y
                    };

                    if (isPathCommand) {
                        gmEditPoint.g2X = command.X0;
                        gmEditPoint.g2Y = command.Y0;
                    }

                    gmEditPoint.X = _relative_x;
                    gmEditPoint.Y = _relative_y;

                    arrPathCommand[gmEditPoint.pathC2] = command;
                }
            }
            this.overlayGeometry.pathLst.length = 1;
            var oDrawPath = this.overlayGeometry.pathLst[0];
            if(oDrawPath) {
                oDrawPath.ArrPathCommand.length = 0;
                oDrawPath.stroke = true;
                if(prevPoint) {
                    oDrawPath.ArrPathCommand.push({id:AscFormat.moveTo, X: prevPoint.X, Y: prevPoint.Y});
                    if(arrPathCommand[gmEditPoint.pathC1]) {
                        oDrawPath.ArrPathCommand.push(arrPathCommand[gmEditPoint.pathC1]);
                    }
                    if(arrPathCommand[gmEditPoint.pathC2]) {
                        oDrawPath.ArrPathCommand.push(arrPathCommand[gmEditPoint.pathC2]);
                    }
                }
                else {
                    oDrawPath.ArrPathCommand.push({id:AscFormat.moveTo, X: gmEditPoint.X, Y: gmEditPoint.Y});
                    if(arrPathCommand[gmEditPoint.pathC2]) {
                        oDrawPath.ArrPathCommand.push(arrPathCommand[gmEditPoint.pathC2]);
                    }
                }
            }
        }, this, []);
    };

    EditShapeGeometryTrack.prototype.getBounds = function() {
        var bounds_checker = new  AscFormat.CSlideBoundsChecker();
        bounds_checker.init(Page_Width, Page_Height, Page_Width, Page_Height);
        this.overlayObject.TransformMatrix = this.originalShape.transform;
        this.overlayObject.draw(bounds_checker);
        var tr = this.originalShape.transform;
        var arr_p_x = [];
        var arr_p_y = [];
        arr_p_x.push(tr.TransformPointX(0,0));
        arr_p_y.push(tr.TransformPointY(0,0));
        arr_p_x.push(tr.TransformPointX(this.originalShape.extX,0));
        arr_p_y.push(tr.TransformPointY(this.originalShape.extX,0));
        arr_p_x.push(tr.TransformPointX(this.originalShape.extX,this.originalShape.extY));
        arr_p_y.push(tr.TransformPointY(this.originalShape.extX,this.originalShape.extY));
        arr_p_x.push(tr.TransformPointX(0,this.originalShape.extY));
        arr_p_y.push(tr.TransformPointY(0,this.originalShape.extY));

        this.calculateMinMax();
        var oRectBounds = this.getRectBounds();
        arr_p_x.push(oRectBounds.XLT);
        arr_p_x.push(oRectBounds.XRB);
        arr_p_y.push(oRectBounds.YLT);
        arr_p_y.push(oRectBounds.YRB);


        bounds_checker.Bounds.min_x = Math.min.apply(Math, arr_p_x);
        bounds_checker.Bounds.max_x = Math.max.apply(Math, arr_p_x);
        bounds_checker.Bounds.min_y = Math.min.apply(Math, arr_p_y);
        bounds_checker.Bounds.max_y = Math.max.apply(Math, arr_p_y);

        var oOffset = this.getXfrmOffset();
        bounds_checker.Bounds.posX = oOffset.OffX;
        bounds_checker.Bounds.posY = oOffset.OffY;
        bounds_checker.Bounds.extX = this.xMax - this.xMin;
        bounds_checker.Bounds.extY = this.yMax - this.yMin;

        return bounds_checker.Bounds;
    };

    EditShapeGeometryTrack.prototype.addPathCommandInfo = function (command, arrPathElem, x1, y1, x2, y2, x3, y3) {
        AscFormat.ExecuteNoHistory(function () {
            switch (command) {
                case 0: {
                    var path = new AscFormat.Path();
                    path.setExtrusionOk(x1 || false);
                    path.setFill(y1 || "norm");
                    path.setStroke(x2 != undefined ? x2 : true);
                    path.setPathW(y2);
                    path.setPathH(x3);
                    this.AddPath(path);
                    break;
                }
                case 1: {
                    this.geometry.pathLst[arrPathElem].moveTo(x1, y1);
                    break;
                }
                case 2: {
                    this.geometry.pathLst[arrPathElem].lnTo(x1, y1);
                    break;
                }
                case 3: {
                    this.geometry.pathLst[arrPathElem].arcTo(x1/*wR*/, y1/*hR*/, x2/*stAng*/, y2/*swAng*/);
                    break;
                }
                case 4: {
                    this.geometry.pathLst[arrPathElem].quadBezTo(x1, y1, x2, y2);
                    break;
                }
                case 5: {
                    this.geometry.pathLst[arrPathElem].cubicBezTo(x1, y1, x2, y2, x3, y3);
                    break;
                }
                case 6: {
                    this.geometry.pathLst[arrPathElem].close();
                    break;
                }
            }
        }, this, []);
    };

    EditShapeGeometryTrack.prototype.getRectBounds = function() {
        var oTr = this.transform;
        var aTX = [];
        var aTY = [];
        aTX.push(oTr.TransformPointX(this.xMin, this.yMin));
        aTX.push(oTr.TransformPointX(this.xMax, this.yMin));
        aTX.push(oTr.TransformPointX(this.xMax, this.yMax));
        aTX.push(oTr.TransformPointX(this.xMin, this.yMax));
        aTY.push(oTr.TransformPointY(this.xMin, this.yMin));
        aTY.push(oTr.TransformPointY(this.xMax, this.yMin));
        aTY.push(oTr.TransformPointY(this.xMax, this.yMax));
        aTY.push(oTr.TransformPointY(this.xMin, this.yMax));
        var dXLT = Math.min.apply(Math, aTX);
        var dXRB = Math.max.apply(Math, aTX);
        var dYLT = Math.min.apply(Math, aTY);
        var dYRB = Math.max.apply(Math, aTY);
        return {XLT: dXLT, XRB: dXRB, YLT: dYLT, YRB: dYRB }
    };

    EditShapeGeometryTrack.prototype.getXfrmOffset = function() {
        var oRectBounds = this.getRectBounds();
        var dExtX = this.xMax - this.xMin;
        var dExtY = this.yMax - this.yMin;
        var dXLT = oRectBounds.XLT;
        var dXRB = oRectBounds.XRB;
        var dYLT = oRectBounds.YLT;
        var dYRB = oRectBounds.YRB;
        var dXC = (dXLT + dXRB) / 2.0;
        var dYC = (dYLT + dYRB) / 2.0;
        var dOffX = dXC - dExtX / 2.0;
        var dOffY = dYC - dExtY / 2.0;
        var oGroup = this.originalObject.group;
        if(oGroup) {
            dOffX -= oGroup.transform.tx;
            dOffY -= oGroup.transform.ty;
        }
        return {OffX: dOffX, OffY: dOffY};

    };
		EditShapeGeometryTrack.prototype.checkDrawingPartWithHistory = function () {
			if (this.originalObject.checkDrawingPartWithHistory) {
				const newObject = this.originalObject.checkDrawingPartWithHistory();
				if (newObject) {
					this.originalObject = newObject;
					this.originalShape = newObject;
				}
			}
		};
    EditShapeGeometryTrack.prototype.trackEnd = function(bWord) {
        this.addCommandsInPathInfo();
        //set new extents
        var dExtX = this.xMax - this.xMin;
        var dExtY = this.yMax - this.yMin;
        var oSpPr = this.originalObject.spPr;
        var oXfrm = oSpPr.xfrm;
        var oOffset;
        if(this.originalObject.animMotionTrack) {
            oOffset = this.getXfrmOffset();
            this.originalObject.updateAnimation(oOffset.OffX, oOffset.OffY, dExtX, dExtY, 0, this.geometry, true);
        }
        else {
            oXfrm.setExtX(dExtX);
            oXfrm.setExtY(dExtY);
            oXfrm.setRot(0);
            //set new position
            if(bWord && !this.originalObject.group) {
                oXfrm.setOffX(0);
                oXfrm.setOffY(0);
            }
            else {
                oOffset = this.getXfrmOffset();
                oXfrm.setOffX(oOffset.OffX);
                oXfrm.setOffY(oOffset.OffY);
            }
            oSpPr.setGeometry(this.geometry.createDuplicate());
            this.originalObject.checkDrawingBaseCoords();
        }

        if(this.addedPointIdx !== null) {
            var oGmSelection = this.getGmSelection();
            if(oGmSelection) {
                oGmSelection.setGmEditPointIdx(this.addedPointIdx);
            }
        }
        if(this.drawingObjects) {
            this.drawingObjects.resetConnectors([this.originalObject]);
        }
    };

    EditShapeGeometryTrack.prototype.convertToBezier = function() {
        var geometry = this.geometry;
        AscFormat.ExecuteNoHistory(
            function () {
                var countArc = 0;
                for(var j = 0; j < geometry.pathLst.length; j++) {
                    geometry.pathLst[j].ArrPathCommandInfo = [];
                    var pathPoints = geometry.pathLst[j].ArrPathCommand;
                    var arrCommandsType = [];
                    for (var i = 0; i < pathPoints.length; i++) {
                        var elem = pathPoints[i];
                        var elemX = null, elemY = null;
                        switch (elem.id) {
                            case PathType.POINT:
                                elemX = elem.X;
                                elemY = elem.Y;
                                arrCommandsType.push(PathType.POINT);
                                break;
                            case PathType.LINE:
                                elemX = elem.X;
                                elemY = elem.Y;
                                arrCommandsType.push(PathType.LINE);
                                break;
                            case PathType.ARC:
                                var oDrawer = new AscFormat.PathAccumulator();
                                AscFormat.ArcToCurvers(oDrawer, elem.stX, elem.stY, elem.wR, elem.hR, elem.stAng, elem.swAng);
                                var aPathCommands = oDrawer.pathCommand;
                                pathPoints.splice(i, 1);
                                for(var nIdx = 0; nIdx < aPathCommands.length; ++nIdx) {
                                    var oCommand = aPathCommands[nIdx];
                                    switch (oCommand.id) {
                                        case AscFormat.moveTo: {
                                            if(nIdx === 0 && pathPoints[i - 1] && pathPoints[i - 1].id === PathType.POINT) {
                                                pathPoints[i - 1].X = oCommand.X;
                                                pathPoints[i - 1].Y = oCommand.Y;
                                            }
                                            break;
                                        }
                                        case AscFormat.bezier4: {
                                            if((oCommand.X0 !== oCommand.X1 || oCommand.X1 !== oCommand.X2) && (oCommand.Y0 !== oCommand.Y1 || oCommand.Y1 !== oCommand.Y2)) {
                                                var elemArc = {
                                                    id: PathType.ARC,
                                                    X0: oCommand.X0,
                                                    Y0: oCommand.Y0,
                                                    X1: oCommand.X1,
                                                    Y1: oCommand.Y1,
                                                    X2: oCommand.X2,
                                                    Y2: oCommand.Y2
                                                };
                                                arrCommandsType.push(PathType.ARC);
                                                pathPoints.splice(i, 0, elemArc);
                                                i++;
                                            }
                                            break;
                                        }
                                    }
                                }
                                i = i - 1;
                                countArc++;
                                break;
                            case PathType.BEZIER_3:
                                elemX = elem.X1;
                                elemY = elem.Y1;
                                arrCommandsType.push(PathType.BEZIER_3);
                                break;
                            case PathType.BEZIER_4:
                                elemX = elem.X2;
                                elemY = elem.Y2;
                                arrCommandsType.push(PathType.BEZIER_4);
                                break;
                            case PathType.END:
                                arrCommandsType.push(PathType.END);
                                break;
                        }

                        if (elemX !== undefined && elemY !== undefined && elem.id !== PathType.ARC && elem.id !== PathType.END) {
                            pathPoints[i] = elem;
                        }
                    }
                    var start_index = 0;
                    // insert pathCommand when end point is not equal to the start point, then draw a line between them
                    for (var cur_index = 1; cur_index < pathPoints.length; cur_index++) {

                        if(pathPoints[cur_index].id === PathType.POINT) {
                            start_index = cur_index;
                        }

                        if (pathPoints[cur_index].id === PathType.END && (!pathPoints[cur_index + 1] || pathPoints[cur_index + 1].id === PathType.POINT)) {
                            var prevCommand = pathPoints[cur_index - 1];
                            var pointCommand = pathPoints[start_index];
                            var prevCommandX = this.getCommandLastPointX(prevCommand);
                            var prevCommandY = this.getCommandLastPointY(prevCommand);

                            var firstPointX = parseFloat(pathPoints[start_index].X.toFixed(2));
                            var firstPointY = parseFloat(pathPoints[start_index].Y.toFixed(2));
                            var lastPointX = parseFloat(prevCommandX.toFixed(2));
                            var lastPointY = parseFloat(prevCommandY.toFixed(2));
                            if (firstPointX !== lastPointX || firstPointY !== lastPointY) {

                                pathPoints.splice(cur_index, 0,
                                    {
                                        id: PathType.LINE,
                                        X: pointCommand.X,
                                        Y: pointCommand.Y
                                    });
                                arrCommandsType.splice(cur_index, 0, PathType.LINE);
                                ++cur_index;
                            }
                        }
                    }

                    var oThis = this;
                    pathPoints.forEach(function (elem, cur_index) {

                        var prevCommand = pathPoints[cur_index - 1];
                        if(prevCommand) {
                            var prevCommandX = oThis.getCommandLastPointX(prevCommand);
                            var prevCommandY = oThis.getCommandLastPointY(prevCommand);
                        }

                        switch (elem.id) {
                            case 1:
                                pathPoints[cur_index] = {
                                    id: PathType.LINE,
                                    X0: (prevCommandX + elem.X / 2) / (3 / 2),
                                    Y0: (prevCommandY + elem.Y / 2) / (3 / 2),
                                    X1: (prevCommandX + elem.X * 2) / 3,
                                    Y1: (prevCommandY + elem.Y * 2) / 3,
                                    X2: elem.X,
                                    Y2: elem.Y
                                };
                                break;
                            case 3:
                                pathPoints[cur_index] = {
                                    id: PathType.BEZIER_3,
                                    X0: (elem.X0 + prevCommandX) / 2,
                                    Y0: (elem.Y0 + prevCommandY) / 2,
                                    X1: (elem.X1 + elem.X0) / 2,
                                    Y1: (elem.Y1 + elem.Y0) / 2,
                                    X2: elem.X1,
                                    Y2: elem.Y1
                                };
                                break;
                        }
                    });
                    if(this.arrPathCommandsType.length < geometry.pathLst.length) {
                        this.arrPathCommandsType.push(arrCommandsType);
                    }
                }
                geometry.setPreset(null);
                geometry.rectS = null;
                geometry.ahXYLst.length = 0;
                geometry.ahXYLstInfo.length = 0;
                geometry.ahPolarLst.length = 0;
                geometry.ahPolarLstInfo.length = 0;
                geometry.cnxLstInfo.length = 0;
                geometry.cnxLst.length = 0;
            }, this, []);
    };

    EditShapeGeometryTrack.prototype.createGeometryEditList = function() {
        AscFormat.ExecuteNoHistory(function(){
            var geometry = this.geometry;
            this.gmEditList = [];
            for(var j = 0; j < geometry.pathLst.length; j++) {

                var start_index = 0, isFirstAndLastPointsEqual = false;
                var pathPoints = geometry.pathLst[j].ArrPathCommand;

                for (var index = 0; index < pathPoints.length; index++) {
                    var curCommand = pathPoints[index];
                    var nextPath = pathPoints[index + 1];

                    if (curCommand.id !== PathType.END) {
                        var nextIndex = 0;
                        var isAddingStartPoint = false;

                        if (!nextPath || nextPath.id === PathType.POINT || nextPath.id === PathType.END) {

                            nextIndex = isFirstAndLastPointsEqual ? (index === start_index ? null : start_index + 1) : null;

                            if (nextPath) {
                                start_index = nextPath.id === PathType.POINT ? index + 1 : index + 2;
                                isFirstAndLastPointsEqual = false;
                            }
                        } else {
                            nextIndex = index + 1;
                        }

                        if (curCommand.id === PathType.POINT) {
                            //finding last point in figure element
                            var i = 1;
                            while((index + i <= pathPoints.length - 1) && pathPoints[index + i].id !== PathType.POINT) {
                                ++i;
                            }
                            if(pathPoints[index + i - 1].id === PathType.END) {
                                --i;
                            }
                            var firstPoint = pathPoints[start_index];
                            var firstPointX = this.getCommandLastPointX(firstPoint);
                            firstPointX = parseFloat(firstPointX.toFixed(2));
                            var firstPointY = this.getCommandLastPointY(firstPoint);
                            firstPointY = parseFloat(firstPointY.toFixed(2));

                            var lastPoint = pathPoints[index + i - 1];
                            var lastPointX = this.getCommandLastPointX(lastPoint);
                            lastPointX = parseFloat(lastPointX.toFixed(2));
                            var lastPointY =  this.getCommandLastPointY(lastPoint);
                            lastPointY = parseFloat(lastPointY.toFixed(2));

                            (firstPointX !== lastPointX || firstPointY !== lastPointY) ? isAddingStartPoint = true : isFirstAndLastPointsEqual = true;
                        }

                        if(pathPoints[index].id !== PathType.POINT || isAddingStartPoint) {
                            var nextCommand = pathPoints[nextIndex];

                            var g1X = curCommand.X1;
                            var g1Y = curCommand.Y1;
                            var g2X = nextCommand ? nextCommand.X0 : undefined;
                            var g2Y = nextCommand ? nextCommand.Y0 : undefined;


                            var curPoint = {
                                g1X: g1X,
                                g1Y: g1Y,
                                g2X: g2X,
                                g2Y: g2Y,
                                X: (this.getCommandLastPointX(curCommand)),
                                Y: (this.getCommandLastPointY(curCommand)),
                                pathC1: index,
                                pathC2: nextIndex,
                                pathIndex: j
                            };
                            this.gmEditList.push(curPoint);
                        }
                    }
                    curCommand.id = (curCommand.id !== PathType.POINT && curCommand.id !== PathType.END) ? PathType.BEZIER_4 : curCommand.id;
                }
            }
            var startIndex = 0;
            for (var cur_index = 0; cur_index < this.gmEditList.length; cur_index++) {
                if(this.gmEditList[cur_index].pathC2 > this.gmEditList[cur_index].pathC1) {
                    this.gmEditList[cur_index].nextPoint = this.gmEditList[cur_index + 1];
                    this.gmEditList[cur_index + 1].prevPoint = this.gmEditList[cur_index];
                } else {
                    this.gmEditList[cur_index].nextPoint =  this.gmEditList[startIndex];
                    this.gmEditList[startIndex].prevPoint = this.gmEditList[cur_index];
                    startIndex = cur_index + 1;
                }
            }

            //update gmEditPoint coords
            var gmEditPoint = this.getGmEditPt();
            if(gmEditPoint) {
                var pointC1 = gmEditPoint.pathC1;
                var pointC2 = gmEditPoint.pathC2;
                var isHitInFirstCPoint = gmEditPoint.isHitInFirstCPoint;
                var isHitInSecondCPoint = gmEditPoint.isHitInSecondCPoint;
                var oThis = this;
                this.gmEditList.forEach(function(elem) {
                    if(elem.pathIndex === gmEditPoint.pathIndex && elem.pathC1 === pointC1 && elem.pathC2 === pointC2) {
                        gmEditPoint = elem;
                        gmEditPoint.isHitInFirstCPoint = isHitInFirstCPoint;
                        gmEditPoint.isHitInSecondCPoint = isHitInSecondCPoint;
                    }
                })
            }
        }, this, []);
    };

    EditShapeGeometryTrack.prototype.calculateMinMax = function() {
        var geometry = this.geometry;
        var last_x = this.gmEditList[0].X,
            last_y = this.gmEditList[0].Y,
            xMin = last_x, yMin = last_y, xMax = last_x, yMax = last_y;
        for (var i = 0; i < geometry.pathLst.length; i++) {
            var arrPathCommand = geometry.pathLst[i].ArrPathCommand;
            geometry.pathLst[i].ArrPathCommandInfo = [];
            for (var j = 0; j < arrPathCommand.length; ++j) {

                var path_command = arrPathCommand[j];

                if (path_command.id === PathType.BEZIER_4) {

                    var bezier_polygon = AscFormat.partition_bezier4(last_x, last_y, path_command.X0, path_command.Y0, path_command.X1, path_command.Y1, path_command.X2, path_command.Y2, AscFormat.APPROXIMATE_EPSILON);
                    for (var point_index = 1; point_index < bezier_polygon.length; ++point_index) {
                        var cur_point = bezier_polygon[point_index];
                        if (xMin > cur_point.x)
                            xMin = cur_point.x;
                        if (xMax < cur_point.x)
                            xMax = cur_point.x;
                        if (yMin > cur_point.y)
                            yMin = cur_point.y;
                        if (yMax < cur_point.y)
                            yMax = cur_point.y;

                        last_x = path_command.X2;
                        last_y = path_command.Y2;
                    }
                }
            }
        }
        this.xMin = xMin;
        this.xMax = xMax;
        this.yMin = yMin;
        this.yMax = yMax;

    };

    EditShapeGeometryTrack.prototype.addCommandsInPathInfo = function() {
        AscFormat.ExecuteNoHistory(
            function(){
                var geometry = this.geometry;
                this.geometry.setPreset(null);
                this.calculateMinMax();
                var w = this.xMax - this.xMin, h = this.yMax - this.yMin;
                var kw, kh, pathW, pathH;
                if (w > 0) {
                    pathW = 43200;
                    kw = 43200 / w;
                }
                else {
                    pathW = 0;
                    kw = 0;
                }
                if (h > 0) {
                    pathH = 43200;
                    kh = 43200 / h;
                }
                else {
                    pathH = 0;
                    kh = 0;
                }

                for (var i = 0; i < geometry.pathLst.length; i++) {
                    var oPath = geometry.pathLst[i];
                    oPath.ArrPathCommandInfo.length = 0;
                    var arrPathCommand = oPath.ArrPathCommand;
                    var lastX = null, lastY = null;
                    for (var j = 0; j < arrPathCommand.length; ++j) {

                        switch (arrPathCommand[j].id) {
                            case PathType.POINT: {
                                lastX = arrPathCommand[j].X;
                                lastY = arrPathCommand[j].Y;
                                this.addPathCommandInfo(1, i, (((arrPathCommand[j].X - this.xMin) * kw) >> 0) + "", (((arrPathCommand[j].Y - this.yMin) * kh) >> 0) + "");
                                break;
                            }
                            case PathType.BEZIER_4: {
                                //check if it is possible to add line
                                var bLine = false;
                                if(AscFormat.isRealNumber(lastX) && AscFormat.isRealNumber(lastY)) {
                                    var dX = arrPathCommand[j].X0 - lastX;
                                    var dY = arrPathCommand[j].Y0 - lastY;
                                    var dEps = 0.01;
                                    if(AscFormat.fApproxEqual(dY, 0, dEps)) {
                                        if(AscFormat.fApproxEqual(arrPathCommand[j].Y1 - lastY, 0, dEps) && AscFormat.fApproxEqual(arrPathCommand[j].Y2 - lastY, 0, dEps)) {
                                            bLine = true;
                                        }
                                    }
                                    else {
                                        var dK = dX / dY;
                                        dX = arrPathCommand[j].X1 - lastX;
                                        dY = arrPathCommand[j].Y1 - lastY;
                                        if(!AscFormat.fApproxEqual(dY, 0, dEps)) {
                                            var dK1 = dX / dY;
                                            if(AscFormat.fApproxEqual(dK1, dK, dEps)) {
                                                dX = arrPathCommand[j].X2 - lastX;
                                                dY = arrPathCommand[j].Y2 - lastY;
                                                if(!AscFormat.fApproxEqual(dY, 0, dEps)) {
                                                    dK1 = dX / dY;
                                                    if(AscFormat.fApproxEqual(dK1, dK, dEps)) {
                                                        bLine = true;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                if(bLine) {
                                    this.addPathCommandInfo(2, i, (((arrPathCommand[j].X2 - this.xMin) * kw) >> 0) + "", (((arrPathCommand[j].Y2 - this.yMin) * kh) >> 0) + "");
                                }
                                else {
                                    this.addPathCommandInfo(5, i, (((arrPathCommand[j].X0 - this.xMin) * kw) >> 0) + "", (((arrPathCommand[j].Y0 - this.yMin) * kh) >> 0) + "", (((arrPathCommand[j].X1 - this.xMin) * kw) >> 0) + "", (((arrPathCommand[j].Y1 - this.yMin) * kh) >> 0) + "", (((arrPathCommand[j].X2 - this.xMin) * kw) >> 0) + "", (((arrPathCommand[j].Y2 - this.yMin) * kh) >> 0) + "");
                                }
                                lastX = arrPathCommand[j].X2;
                                lastY = arrPathCommand[j].Y2;
                                break;
                            }
                            case PathType.END: {
                                this.addPathCommandInfo(6, i);
                            }
                        }
                    }
                    geometry.pathLst[i].pathW = pathW;
                    geometry.pathLst[i].pathH = pathH;
                }
            }, this, []
        );
    };


    EditShapeGeometryTrack.prototype.findBezier4Param = function(XT, YT, X0, Y0, X1, Y1, X2, Y2, X3, Y3) {
       var nSteps = 1000;
       var dStride = 1/nSteps;
       var dT = 0;
       var dTResult = dT;
       var CX = this.bezier4Pos(dT, X0, X1, X2, X3);
       var CY = this.bezier4Pos(dT, Y0, Y1, Y2, Y3);
       var dDist = Math.abs(XT - CX) + Math.abs(YT - CY);
       var dMinDist = dDist;
       for(var nStep = 0; nStep < nSteps; ++nStep) {
           dT+= dStride;
           CX = this.bezier4Pos(dT, X0, X1, X2, X3);
           CY = this.bezier4Pos(dT, Y0, Y1, Y2, Y3);
           dDist = Math.abs(XT - CX) + Math.abs(YT - CY);
           if(dDist < dMinDist) {
               dTResult = dT;
               dMinDist = dDist;
           }
       }
       return dTResult;
    };
    EditShapeGeometryTrack.prototype.findBezier3Param = function(XT, YT, X0, Y0, X1, Y1, X2, Y2) {
       var nSteps = 1000;
       var dStride = 1/nSteps;
       var dT = 0;
       var dTResult = dT;
       var CX = this.bezier3Pos(dT, X0, X1, X2);
       var CY = this.bezier3Pos(dT, Y0, Y1, Y2);
       var dDist = Math.abs(XT - CX) + Math.abs(YT - CY);
       var dMinDist = dDist;
       for(var nStep = 0; nStep < nSteps; ++nStep) {
           dT+= dStride;
           CX = this.bezier3Pos(dT, X0, X1, X2);
           CY = this.bezier3Pos(dT, Y0, Y1, Y2);
           dDist = Math.abs(XT - CX) + Math.abs(YT - CY);
           if(dDist < dMinDist) {
               dTResult = dT;
               dMinDist = dDist;
           }
       }
       return dTResult;
    };

    EditShapeGeometryTrack.prototype.bezier4Pos = function(t, C0, C1, C2, C3) {
       var dDT = 1 - t;
       var dDT2 = dDT*dDT;
       var dDT3 = dDT2*dDT;
       var dT = t;
       var dT2 = dT*dT;
       var dT3 = dT2*dT;
       return C0*dDT3 + 3*C1*t*dDT2 + 3*C2*dDT*t*t + C3*t*t*t;
    };
    EditShapeGeometryTrack.prototype.bezier3Pos = function(t, C0, C1, C2) {
       var dDT = 1 - t;
       var dDT2 = dDT*dDT;
       var dT = t;
       var dT2 = dT*dT;
       return C0*dDT2 + 2*C1*t*dDT + C2*dT2;
    };

    EditShapeGeometryTrack.prototype.addPoint = function(oAddingPoint, X, Y) {
        return AscFormat.ExecuteNoHistory(function() {
            var geometry = this.geometry;
            var commandIndex = oAddingPoint.commandIndex;
            var pathIndex = oAddingPoint.pathIndex;
            var tx = this.invertTransform.TransformPointX(X, Y);
            var ty = this.invertTransform.TransformPointY(X, Y);

            var pathElem = geometry.pathLst[pathIndex].ArrPathCommand;
            var curCommand = pathElem[commandIndex];

            var gmEditListElem = this.gmEditList.filter(function(elem) {
                return elem.pathC1 === commandIndex;
            })[0];
            var prevCommand_1 = gmEditListElem.prevPoint;
            var prevCommand_2 = prevCommand_1.prevPoint;

            var curCommandX = this.getCommandLastPointX(curCommand);
            var curCommandY = this.getCommandLastPointY(curCommand);

            var X0 = prevCommand_1.X + (tx - prevCommand_2.X) / 4;
            var Y0 = prevCommand_1.Y + (ty - prevCommand_2.Y) / 4;
            var X1 = tx - (curCommandX - prevCommand_1.X) / 4;
            var Y1 = ty - (curCommandY - prevCommand_1.Y) / 4;
            var newPathElem = {id: PathType.BEZIER_4, X0: X0, Y0: Y0, X1: X1, Y1: Y1, X2: tx, Y2: ty};
            ///curCommand.X0 = tx + (curCommandX - prevCommand_1.X) / 4;
            ///curCommand.Y0 = ty + (curCommandY - prevCommand_1.Y) / 4;

            //pathElem.splice(commandIndex, 0, newPathElem);
//            this.arrPathCommandsType[pathIndex].splice(commandIndex, 0, PathType.BEZIER_4);
            var oPrevCommand = pathElem[commandIndex - 1];
            if(!oPrevCommand) {
                return;
            }
            var dSplitT;
            var dPrevX = this.getCommandLastPointX(oPrevCommand);
            var dPrevY = this.getCommandLastPointY(oPrevCommand);
            var aCommands;
            var oFirstCommand, oSecondCommand;
            var nIdx;
            if(curCommand.id === AscFormat.bezier4) {
                dSplitT = this.findBezier4Param(tx, ty, dPrevX, dPrevY, curCommand.X0, curCommand.Y0,
                    curCommand.X1, curCommand.Y1,
                    curCommand.X2, curCommand.Y2);
                aCommands = splitCurveAt(dSplitT, dPrevX, dPrevY, curCommand.X0, curCommand.Y0,
                    curCommand.X1, curCommand.Y1,
                    curCommand.X2, curCommand.Y2);
                nIdx = 2;
                oFirstCommand = {
                    id:AscFormat.bezier4,
                    X0: aCommands[nIdx++],
                    Y0: aCommands[nIdx++],
                    X1: aCommands[nIdx++],
                    Y1: aCommands[nIdx++],
                    X2: aCommands[nIdx++],
                    Y2: aCommands[nIdx++]
                };
                oSecondCommand = {
                    id:AscFormat.bezier4,
                    X0: aCommands[nIdx++],
                    Y0: aCommands[nIdx++],
                    X1: aCommands[nIdx++],
                    Y1: aCommands[nIdx++],
                    X2: aCommands[nIdx++],
                    Y2: aCommands[nIdx++]
                };

                pathElem.splice(commandIndex, 1);
                this.arrPathCommandsType[pathIndex].splice(commandIndex, 1);

                pathElem.splice(commandIndex, 0, oFirstCommand);
                this.arrPathCommandsType[pathIndex].splice(commandIndex, 0, PathType.BEZIER_4);
                pathElem.splice(commandIndex + 1, 0, oSecondCommand);
                this.arrPathCommandsType[pathIndex].splice(commandIndex + 1, 0, PathType.BEZIER_4);
            }
            else if(curCommand.id === AscFormat.bezier3) {
                dSplitT = this.findBezier3Param(tx, ty, dPrevX, dPrevY, curCommand.X0, curCommand.Y0,
                    curCommand.X1, curCommand.Y1);
                aCommands = splitCurveAt(dSplitT, dPrevX, dPrevY, curCommand.X0, curCommand.Y0,
                    curCommand.X1, curCommand.Y1);
                nIdx = 2;
                oFirstCommand = {
                    id:AscFormat.bezier4,
                    X0: aCommands[nIdx++],
                    Y0: aCommands[nIdx++],
                    X1: aCommands[nIdx++],
                    Y1: aCommands[nIdx++],
                    X2: aCommands[nIdx++],
                    Y2: aCommands[nIdx++]
                };
                oSecondCommand = {
                    id:AscFormat.bezier4,
                    X0: aCommands[nIdx++],
                    Y0: aCommands[nIdx++],
                    X1: aCommands[nIdx++],
                    Y1: aCommands[nIdx++],
                    X2: aCommands[nIdx++],
                    Y2: aCommands[nIdx++]
                };
                pathElem.splice(commandIndex, 1);
                this.arrPathCommandsType[pathIndex].splice(commandIndex, 1);

                pathElem.splice(commandIndex, 0, oFirstCommand);
                this.arrPathCommandsType[pathIndex].splice(commandIndex, 0, PathType.BEZIER_4);
                pathElem.splice(commandIndex + 1, 0, oSecondCommand);
                this.arrPathCommandsType[pathIndex].splice(commandIndex + 1, 0, PathType.BEZIER_4);
            }
            geometry.pathLst[pathIndex].ArrPathCommandInfo = [];
            this.addCommandsInPathInfo();
            this.createGeometryEditList();
            var oHitData = this.hitToGmEditLst(X, Y, true);
            if(oHitData) {
                this.addedPointIdx = oHitData.gmEditPointIdx;
                var oPt = this.gmEditList[this.addedPointIdx];
                if(oPt) {
                    var oOriginalAddedPoint = {
                        g1X: oPt.g1X,
                        g1Y: oPt.g1Y,
                        g2X: oPt.g2X,
                        g2Y: oPt.g2Y,
                        X: oPt.X,
                        Y: oPt.Y,
                        pathC1: oPt.pathC1,
                        pathC2: oPt.pathC2,
                        pathIndex: oPt.pathC2
                    };
                    this.originalAddedPoint = oOriginalAddedPoint;
                }
                var oGeomSelection = this.getGmSelection();
                if(oGeomSelection) {
                    oGeomSelection.resetGmEditPointIdx();
                }
            }
            // AscCommon.History.Create_NewPoint(0);
            // this.trackEnd();
            // editor.WordControl.m_oLogicDocument.Recalculate();
        }, this, []);

    };
    EditShapeGeometryTrack.prototype.getCommandLastPointX = function(oCommand) {
        return this.getCommandLastPointCoord(oCommand.X, oCommand.X1, oCommand.X2);
    };
    EditShapeGeometryTrack.prototype.getCommandLastPointY = function(oCommand) {
        return this.getCommandLastPointCoord(oCommand.Y, oCommand.Y1, oCommand.Y2);
    };
    EditShapeGeometryTrack.prototype.getCommandLastPointCoord = function(C, C1, C2) {
        if(C !== undefined) {
            return C;
        }
        if(C2 !== undefined) {
            return C2;
        }
        if(C1 !== undefined) {
            return C1;
        }
        return null;
    };

    EditShapeGeometryTrack.prototype.deletePoint = function() {
        AscFormat.ExecuteNoHistory(function() {
            var geometry = this.geometry;
            var gmEditPoint = this.getGmEditPt();
            if(!gmEditPoint) {
                return;
            }
                var pathIndex = gmEditPoint.pathIndex,
                pathElem = geometry.pathLst[pathIndex],
                arrayCommands = geometry.pathLst[pathIndex].ArrPathCommand;

            // if(pathElem && pathElem.stroke === true && pathElem.fill === "none") {
            //     return;
            // }

            var pathC1 = gmEditPoint.pathC1,
                pathC2 = gmEditPoint.pathC2,
                pointCount = 0;

            var increment_index = pathC1;
            var decrement_index = pathC1;

            while(arrayCommands[decrement_index] && arrayCommands[decrement_index].id !== PathType.POINT) {
                --decrement_index;
                ++pointCount;
            }
            while(arrayCommands[increment_index + 1] && arrayCommands[increment_index + 1].id !== PathType.END) {
                ++increment_index;
                ++pointCount;
            }

            if(pointCount > 2) {

                if (pathC1 > pathC2) {
                    if(arrayCommands[pathC1 - 1]) {
                        var prevCommandX = this.getCommandLastPointX(arrayCommands[pathC1 - 1]);
                        var prevCommandY = this.getCommandLastPointY(arrayCommands[pathC1 - 1]);
                        arrayCommands[decrement_index] = {id: PathType.POINT, X: prevCommandX, Y: prevCommandY};
                    }
                    // var t = pathC1;
                    // pathC1 = pathC2;
                    // pathC2 = t;
                }
                var curArrCommandsType = this.arrPathCommandsType[pathIndex];

                //if next command is line, then recalculate to make it
                var nextPath = gmEditPoint.nextPoint.pathC1;
                if (curArrCommandsType[nextPath] && (curArrCommandsType[nextPath] === PathType.LINE)) {
                    var prevX = gmEditPoint.prevPoint.X, prevY = gmEditPoint.prevPoint.Y,
                        nextX = gmEditPoint.nextPoint.X, nextY = gmEditPoint.nextPoint.Y;
                    arrayCommands[nextPath].X0 = (nextX + prevX * 2) / 3;
                    arrayCommands[nextPath].Y0 = (nextY + prevY * 2) / 3;
                    arrayCommands[nextPath].X1 = (nextX + prevX / 2) / (3 / 2);
                    arrayCommands[nextPath].Y1 = (nextY + prevY / 2) / (3 / 2);
                }
                var oNextCommand = arrayCommands[pathC2];
                var oFirstCommand = arrayCommands.splice(pathC1, 1)[0];
                curArrCommandsType.splice(pathC1, 1);
                if(oFirstCommand.id === AscFormat.bezier3 && oNextCommand.id === AscFormat.bezier3) {
                    oNextCommand.X0 = (oNextCommand.X0 + oFirstCommand.X0)/2;
                    oNextCommand.Y0 = (oNextCommand.Y0 + oFirstCommand.Y0)/2;
                }
                if(oFirstCommand.id === AscFormat.bezier4 && oNextCommand.id === AscFormat.bezier4) {
                    oNextCommand.X0 = oFirstCommand.X0;
                    oNextCommand.Y0 = oFirstCommand.Y0;
                }

                this.createGeometryEditList();
                this.addCommandsInPathInfo();
            }
            var oGeomSelection = this.getGmSelection();
            if(oGeomSelection) {
                oGeomSelection.resetGmEditPointIdx();
            }
        }, this, []);
    };

    //EditShapeGeometryTrack.prototype.AddGeomPoint = function(id, X, Y, g1X, g1Y, g2X, g2Y, pathC1, pathC2, prevPoint, nextPoint, isHitInFirstCPoint, isHitInSecondCPoint, isStartPoint, pathIndex)
    //{
    //    this.gmEditPoint = {};
    //    this.gmEditPoint.id = id;
    //    this.gmEditPoint.X = X;
    //    this.gmEditPoint.Y = Y;
    //    this.gmEditPoint.g1X = g1X;
    //    this.gmEditPoint.g1Y = g1Y;
    //    this.gmEditPoint.g2X = g2X;
    //    this.gmEditPoint.g2Y = g2Y;
    //    this.gmEditPoint.pathC1 = pathC1;
    //    this.gmEditPoint.pathC2 = pathC2;
    //    this.gmEditPoint.nextPoint = {
    //        id: nextPoint.id, X: nextPoint.X, Y: nextPoint.Y, g1X: nextPoint.g1X, g1Y: nextPoint.g1Y,
    //        g2X: nextPoint.g2X, g2Y: nextPoint.g2Y, pathC1 : nextPoint.pathC1, pathC2 : nextPoint.pathC2
    //    };
    //    this.gmEditPoint.prevPoint = {
    //        id: prevPoint.id, X: prevPoint.X, Y: prevPoint.Y, g1X: prevPoint.g1X, g1Y: prevPoint.g1Y,
    //        g2X: prevPoint.g2X, g2Y: prevPoint.g2Y, pathC1 : prevPoint.pathC1, pathC2 : prevPoint.pathC2
    //    };
    //        this.gmEditPoint.isHitInFirstCPoint = isHitInFirstCPoint;
    //    this.gmEditPoint.isHitInSecondCPoint = isHitInSecondCPoint;
    //    this.gmEditPoint.isStartPoint = isStartPoint;
    //    this.gmEditPoint.pathIndex = pathIndex;
    //
    //    this.originalEditPoint = {X: X, Y: Y, g1X: g1X, g1Y: g1Y, g2X: g2X, g2Y: g2Y};
    //};

    EditShapeGeometryTrack.prototype.hitToGmEditLst = function(x, y, findNearest) {
        var dx, dy;
        var distance =  this.originalObject.convertPixToMM(AscCommon.global_mouseEvent.KoefPixToMM * AscCommon.TRACK_CIRCLE_RADIUS);
        var tx = this.invertTransform.TransformPointX(x, y);
        var ty = this.invertTransform.TransformPointY(x, y);
        var minDist = 10000;
        var oCandidate = null;
        for (var i = this.gmEditList.length - 1; i >= 0; i--) {
            var gmArr = this.gmEditList[i];
            dx = tx - gmArr.X;
            dy = ty - gmArr.Y;
            var dist = Math.sqrt(dx * dx + dy * dy);
            if (dist < minDist) {
                minDist = dist;
                oCandidate = new CGeomHitData(i, null, null, false);
            }
        }
        if(findNearest) {
            return oCandidate;
        }
        if (minDist < distance) {
            return oCandidate;
        }
        return null;
    };

    EditShapeGeometryTrack.prototype.hitToGeomEdit = function(oCanvas, x, y) {
        if(!this.isCorrect()) {
            return null;
        }
        var dxC1, dyC1, dxC2, dyC2;
        var distance =  this.originalObject.convertPixToMM(AscCommon.global_mouseEvent.KoefPixToMM * AscCommon.TRACK_CIRCLE_RADIUS);
        var geometry = this.geometry;


        var gmEditPoint = this.getGmEditPt();
        var tx = this.invertTransform.TransformPointX(x, y);
        var ty = this.invertTransform.TransformPointY(x, y);
        if(gmEditPoint) {
            // не разрешаем ломать линии в pdf
            if (Asc.editor.isPdfEditor() == false) {
                dxC1 = tx - gmEditPoint.g1X;
                dyC1 = ty - gmEditPoint.g1Y;
                dxC2 = tx - gmEditPoint.g2X;
                dyC2 = ty - gmEditPoint.g2Y;
                if (Math.sqrt(dxC1 * dxC1 + dyC1 * dyC1) < distance) {
                    return new CGeomHitData(this.getGmEditPtIdx(), true, false, false);
                } else if (Math.sqrt(dxC2 * dxC2 + dyC2 * dyC2) < distance) {
                    return new CGeomHitData(this.getGmEditPtIdx(), false, true, false);
                }
            }
        }
        var oGeomData = this.hitToGmEditLst(x, y, false);
        if(oGeomData) {
            return oGeomData;
        }

        // не разрешаем ломать линии в pdf
        if (Asc.editor.isPdfEditor())
            return null;

        var oAddingPoint = {pathIndex: null, commandIndex: null};
        var isHitInPath = geometry.hitInPath(oCanvas, tx, ty, oAddingPoint);
        if(isHitInPath) {
            return new CGeomHitData(null, null, null, oAddingPoint);
        }
        return null;
    };

    EditShapeGeometryTrack.prototype.isCorrect = function() {
        if(this.originalGeometry !== this.getOriginalObjectGeometry() ||
            !AscFormat.fApproxEqual(this.originalX, this.originalObject.x) ||
            !AscFormat.fApproxEqual(this.originalY, this.originalObject.y) ||
            !AscFormat.fApproxEqual(this.originalExtX, this.originalObject.extX) ||
            !AscFormat.fApproxEqual(this.originalExtY, this.originalObject.extY) ||
            !AscFormat.fApproxEqual(this.originalRot, this.originalObject.rot)) {
            return false;
        }
        return true;
    };


    function CGeomHitData(gmEditPointIdx, isHitInFirstCPoint, isHitInSecondCPoint, addingNewPoint) {
        this.gmEditPointIdx = gmEditPointIdx;
        this.isHitInFirstCPoint = isHitInFirstCPoint;
        this.isHitInSecondCPoint = isHitInSecondCPoint;
        this.addingNewPoint = addingNewPoint;
    }
    CGeomHitData.prototype.getPtIdx = function() {
        return this.gmEditPointIdx;
    };

    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].EditShapeGeometryTrack = EditShapeGeometryTrack;



    function splitCurveAt(t, x1, y1, x2, y2, x3, y3, x4, y4) {
        var x1_, y1_, x2_, y2_, x3_, y3_, x4_, y4_;
        if(x4 === undefined || y4 === undefined) {
            x1_ = x1;
            y1_ = y1;
            x2_ = x1 + (2/3)*(x2 - x1);
            y2_ = y1 + (2/3)*(y2 - y1);
            x3_ = x3 + (2/3)*(x2 - x3);
            y3_ = y3 + (2/3)*(y2 - y3);
            x4_ = x3;
            y4_ = y3;
        }
        else {
            x1_ = x1;
            y1_ = y1;
            x2_ = x2;
            y2_ = y2;
            x3_ = x3;
            y3_ = y3;
            x4_ = x4;
            y4_ = y4;
        }
        var x12 = (x2_-x1_)*t+x1_;
        var y12 = (y2_-y1_)*t+y1_;

        var x23 = (x3_-x2_)*t+x2_;
        var y23 = (y3_-y2_)*t+y2_;

        var x34 = (x4_-x3_)*t+x3_;
        var y34 = (y4_-y3_)*t+y3_;

        var x123 = (x23-x12)*t+x12;
        var y123 = (y23-y12)*t+y12;

        var x234 = (x34-x23)*t+x23;
        var y234 = (y34-y23)*t+y23;

        var x1234 = (x234-x123)*t+x123;
        var y1234 = (y234-y123)*t+y123;
        return  [x1_, y1_, x12, y12, x123, y123, x1234, y1234, x234, y234, x34, y34, x4_, y4_];
    }
})(window);

