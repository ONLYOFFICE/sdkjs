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
    function gcd(n, m) {
        return m === 0 ? n : gcd(m, n % m);
    }
    function lcm(n, m) {
        return n * m / gcd(n, m);
    }

    function CDrawingNode(element, par, idx) {
        this.element = element;
        this.partner = null;
        this.par = par;
        this.idx = idx || 0;
        if(Array.isArray(element)) {
            this.children = [];
            for(let nIdx = 0; nIdx < element.length; ++nIdx) {
                let oElement = element[nIdx];
                if(oElement instanceof CDrawingNode) {
                    if(!oElement.partner) {
                        oElement.par = this;
                        this.children.push(oElement);
                    }
                }
                else {
                    this.children.push(new CDrawingNode(oElement, this, nIdx));
                }
            }

        }
    }
    CDrawingNode.prototype.children = [];
    CDrawingNode.prototype.equals = function(oNode) {
        if(!Array.isArray(this.element)) {
            return this.element.compareForMorph(oNode.element, null) === oNode.element;
        }
        return this.element === oNode.element;
    };
    CDrawingNode.prototype.forEachDescendant = function(callback, T) {
        this.children.forEach(function(node) {
            node.forEach(callback, T);
        });
    };
    CDrawingNode.prototype.forEach = function(callback, T) {
        callback.call(T, this);
        this.children.forEach(function(node) {
            node.forEach(callback, T);
        });
    };

    function CDiffMatching() {
        this.clear = true;
    }
    CDiffMatching.prototype.get = function(oNode) {
        return oNode.partner;
    };
    CDiffMatching.prototype.put = function(oNode1, oNode2) {
        oNode1.partner = oNode2;
        oNode2.partner = oNode1;
        this.clear = false;
    };
    function CDiffChange(oOperation) {
        this.pos = -1;
        this.deleteCount = 0;
        this.insert = [];

        var oAnchor = oOperation.anchor;
        this.pos = oAnchor.index;
        if(Array.isArray(oOperation.remove)) {
            this.deleteCount = oOperation.remove.length;
        }
        var nIndex, oNode;
        if(Array.isArray(oOperation.insert)) {
            for(nIndex = 0; nIndex < oOperation.insert.length; ++nIndex) {
                oNode = oOperation.insert[nIndex];
                this.insert.push(oNode.element);
            }
        }
    }
    CDiffChange.prototype.getPos = function() {
        return this.pos;
    };
    function compareDrawings(aDrawings1, aDrawings2) {
        let oBaseNode = new CDrawingNode(aDrawings1, null);
        let oReplaceNode = new CDrawingNode(aDrawings2, null);
        let oMatching = new CDiffMatching();
        oMatching.put(oBaseNode, oReplaceNode);
        let oDiff  = new AscCommon.Diff(oBaseNode, oReplaceNode);
        oDiff.equals = function(a, b)
        {
            return a.equals(b);
        };
        oDiff.clear = true;
        oDiff.matchTrees(oMatching);
        if(!oDiff.clear) {
            compareDrawings(oBaseNode.children, oReplaceNode.children);
        }

        return [oBaseNode, oReplaceNode];
    }

    function CMorphObjectBase(oTexturesCache, nRelH1, nRelH2) {
        this.cache = oTexturesCache;
        const isN = AscFormat.isRealNumber;
        this.relHeight1 = isN(nRelH1) ? nRelH1 : null;
        this.relHeight2 = isN(nRelH2) ? nRelH2 : null;

        this.relHeight = null;

        this.relTime = 0.0;
    }
    CMorphObjectBase.prototype.morph = function (dRelTime) {
        if(this.relHeight1 !== null && this.relHeight2 !== null) {
            this.relHeight = this.relHeight1 + dRelTime * (this.relHeight2 - this.relHeight1);
        }
        else if(this.relHeight1 !== null) {
            this.relHeight = this.relHeight1;
        }
        else {
            this.relHeight = this.relHeight2;
        }
        this.relTime = dRelTime;
    };
    CMorphObjectBase.prototype.morphObjects = function (dRelTime) {

    };
    CMorphObjectBase.prototype.draw = function (oGraphics) {

    };
    CMorphObjectBase.prototype.getValBetween = function(dVal1, dVal2) {
        return dVal1 + (dVal2 - dVal1) * this.relTime;
    };
    CMorphObjectBase.prototype.getMorph = function () {
        return this.cache.parent;
    };
    CMorphObjectBase.prototype.getPlayer1 = function() {
        return this.getMorph().player1;
    };
    CMorphObjectBase.prototype.getPlayer2 = function() {
        return this.getMorph().player2;
    };
    function GetDrawingFormatId(oDrawing) {
        if(!oDrawing) {
            return null;
        }
        if(AscFormat.IsTrueDrawing(oDrawing)) {
            return oDrawing.GetId();
        }
        if(oDrawing instanceof CWrapperBase) {
            if(oDrawing.formatDrawing) {
                return GetDrawingFormatId(oDrawing.formatDrawing);
            }
        }
        return null;
    }

    function CMorphedPath(oTexturesCache, oPath1, nRelH1, oBrush1, oPen1, oTransform1,
                          oPath2, nRelH2, oBrush2, oPen2, oTransform2) {
        CMorphObjectBase.call(this, oTexturesCache, nRelH1, nRelH2);


        if(oPath1.fill !== oPath2.fill || oPath1.stroke !== oPath2.stroke) {
            return;
        }
        this.path1 = oPath1;
        this.brush1 = oBrush1;
        this.pen1 = oPen1;
        this.transform1 = oTransform1;

        this.path2 = oPath2;
        this.brush2 = oBrush2;
        this.pen2 = oPen2;
        this.transform2 = oTransform2;
        this.path1T = null;
        this.path2T = null;
        AscFormat.ExecuteNoHistory(function() {

            this.path1T = new AscFormat.Path();
            this.path1T.setParent(this.path1.parent);
            this.path1.convertToBezierCurves(this.path1T, this.transform1);

            this.path2T = new AscFormat.Path();
            this.path2T.setParent(this.path2.parent);
            this.path2.convertToBezierCurves(this.path2T, this.transform2);
        }, this, []);


        this.path = null;
        this.pen = null;
        this.brush = null;

        this.contours1 = [];
        this.contours2 = [];

        let aContours1 = [];
        let aContours2 = [];
        const aCommands1 = this.path1T.ArrPathCommand;
        const aCommands2 = this.path2T.ArrPathCommand;


        for(let nCmd = 0; nCmd < aCommands1.length; ++nCmd) {
            let oCmd = aCommands1[nCmd];
            if(oCmd.id === AscFormat.moveTo) {
                aContours1.push([]);
            }
            if(aContours1.length === 0) {
                aContours1.push([]);
            }
            let aContour = aContours1[aContours1.length - 1];
            aContour.push(oCmd);
        }
        for(let nCmd = 0; nCmd < aCommands2.length; ++nCmd) {
            let oCmd = aCommands2[nCmd];
            if(oCmd.id === AscFormat.moveTo) {
                aContours2.push([]);
            }
            if(aContours2.length === 0) {
                aContours2.push([]);
            }
            let aContour = aContours2[aContours2.length - 1];
            aContour.push(oCmd);
        }
        if(aContours1.length === aContours2.length) {
            const nContoursCount = aContours1.length;
            for(let nCnt = 0; nCnt < nContoursCount; ++nCnt) {
                let aContour1 = aContours1[nCnt];
                let aContour2 = aContours2[nCnt];
                let oFirstCmd1 = aContour1[0];
                let oFirstCmd2 = aContour2[0];
                let oLastCmd1 = aContour1[aContour1.length - 1];
                let oLastCmd2 = aContour2[aContour2.length - 1];
                if(oLastCmd1.id !== oLastCmd2.id) {
                    return;
                }

                if(oFirstCmd1.id !== AscFormat.moveTo) {
                    return;
                }
                if(oFirstCmd2.id !== AscFormat.moveTo) {
                    return;
                }
                let n = aContour1.length - 1;
                if(oLastCmd1.id === AscFormat.close) {
                    n--;
                }
                let m = aContour2.length - 1;
                if(oLastCmd2.id === AscFormat.close) {
                    m--;
                }
                const nLCM = lcm(n, m);
                const n1 = nLCM / n;
                const m1 = nLCM / m;
                function getBezierCommands(aCommands, nSplitCount) {
                    let aBezier = [];
                    let oLastCommand = aCommands[aCommands.length - 1];
                    let nLastIdx = aCommands.length - 1;
                    if(oLastCommand.id === AscFormat.close) {
                        --nLastIdx;
                    }
                    let dLastX, dLastY;
                    for(let nCmd = 0; nCmd <= nLastIdx; ++nCmd) {
                        let oCmd = aCommands[nCmd];
                        if(oCmd.id === AscFormat.moveTo) {
                            dLastX = oCmd.X;
                            dLastY = oCmd.Y;
                        }
                        else if(oCmd.id === AscFormat.bezier4) {
                            let dX0 = dLastX;
                            let dY0 = dLastY;
                            let dX1 = oCmd.X0;
                            let dY1 = oCmd.Y0;
                            let dX2 = oCmd.X1;
                            let dY2 = oCmd.Y1;
                            let dX3 = oCmd.X2;
                            let dY3 = oCmd.Y2;
                            let aSplitCommand = AscFormat.splitBezier4OnParts(dX0, dY0, dX1, dY1, dX2, dY2, dX3, dY3, nSplitCount);
                            if(!aSplitCommand) {
                                return;
                            }
                            aBezier = aBezier.concat(aSplitCommand);

                            dLastX = oCmd.X2;
                            dLastY = oCmd.Y2;
                        }
                        else {
                            return null;
                        }
                    }
                    return aBezier;
                }
                let aBezier1 = getBezierCommands(aCommands1, n1);
                let aBezier2 = getBezierCommands(aCommands2, m1);
                if(!aBezier1 || !aBezier2 || aBezier1.length !== aBezier2.length) {
                    return;
                }

                function fillContour(aContourT, oFirstCommand, oLastCommand, aBezier, oT) {
                    let oCmd = oFirstCommand;
                    let dX = oCmd.X;
                    let dY = oCmd.Y;
                    aContourT.push([dX, dY]);
                    for(let nCmd = 0; nCmd < aBezier1.length; ++nCmd) {
                        let aBezier4 =  aBezier[nCmd];
                        let dX0 = aBezier4[2];
                        let dY0 = aBezier4[3];
                        let dX1 = aBezier4[4];
                        let dY1 = aBezier4[5];
                        let dX2 = aBezier4[6];
                        let dY2 = aBezier4[7];
                        aContourT.push([dX0, dY0, dX1, dY1, dX2, dY2]);
                    }
                    if(oLastCommand.id === AscFormat.close) {
                        aContourT.push([]);
                    }

                }

                let aContourT1 = [];
                this.contours1.push(aContourT1);
                fillContour(aContourT1, oFirstCmd1, oLastCmd1, aBezier1, this.transform1);

                let aContourT2 = [];
                this.contours2.push(aContourT2);
                fillContour(aContourT2, oFirstCmd2, oLastCmd2, aBezier2, this.transform2);
            }
        }


        const oPath = AscFormat.ExecuteNoHistory(function() {

            let oPath = new AscFormat.Path();
            return oPath;

        }, this, []);
        oPath.fill = oPath1.fill;
        oPath.stroke = oPath1.stroke;
        let aPathCommands = oPath.ArrPathCommand;
        for(let nContour = 0; nContour < this.contours1.length; ++nContour) {
            let aContour1 = this.contours1[nContour];
            let aContour2 = this.contours2[nContour];
            if(aContour1.length !== aContour2.length) {
                return;
            }
            for(let nCmd = 0; nCmd < aContour1.length; ++nCmd) {
                let aCmd1 = aContour1[nCmd];
                let aCmd2 = aContour2[nCmd];
                if(aCmd1.length !== aCmd2.length) {
                    return;
                }
                if(aCmd1.length === 2) {
                    aPathCommands.push({
                        id: AscFormat.moveTo,
                        X: aCmd1[0],
                        Y: aCmd1[1],

                        cmd1: aCmd1,
                        cmd2: aCmd2
                    });
                }
                else if (aCmd1.length === 6) {
                    aPathCommands.push({
                        id: AscFormat.bezier4,
                        X0: aCmd1[0],
                        Y0: aCmd1[1],
                        X1: aCmd1[2],
                        Y1: aCmd1[3],
                        X2: aCmd1[4],
                        Y2: aCmd1[5],

                        cmd1: aCmd1,
                        cmd2: aCmd2
                    });
                }
                else if (aCmd1.length === 0) {
                    aPathCommands.push({
                        id: AscFormat.close,
                        cmd1: aCmd1,
                        cmd2: aCmd2
                    });
                }
            }
        }
        this.path = oPath;
        this.morph(1);
    }
    AscFormat.InitClassWithoutType(CMorphedPath, CMorphObjectBase);
    CMorphedPath.prototype.morph = function (dTime, oInvTransform) {
        if(!this.isValid()) {
            return;
        }
        if(!this.path) {
            return;
        }

        CMorphObjectBase.prototype.morph.call(this, dTime);
        let aCommands = this.path.ArrPathCommand;
        let XT, YT;
        for(let nCmd = 0; nCmd < aCommands.length; ++nCmd) {
            let oCmd = aCommands[nCmd];
            let aCmd1 = oCmd.cmd1;
            let aCmd2 = oCmd.cmd2;
            if(!aCmd1 || !aCmd2) {
                return;
            }
            if(oCmd.id === AscFormat.moveTo) {
                oCmd.X = this.getValBetween(aCmd1[0], aCmd2[0]);
                oCmd.Y = this.getValBetween(aCmd1[1], aCmd2[1]);
                if(oInvTransform) {
                    XT = oInvTransform.TransformPointX(oCmd.X, oCmd.Y);
                    YT = oInvTransform.TransformPointY(oCmd.X, oCmd.Y);
                    oCmd.X = XT;
                    oCmd.Y = YT;
                }
            }
            else if(oCmd.id === AscFormat.bezier4) {
                oCmd.X0 = this.getValBetween(aCmd1[0], aCmd2[0]);
                oCmd.Y0 = this.getValBetween(aCmd1[1], aCmd2[1]);
                oCmd.X1 = this.getValBetween(aCmd1[2], aCmd2[2]);
                oCmd.Y1 = this.getValBetween(aCmd1[3], aCmd2[3]);
                oCmd.X2 = this.getValBetween(aCmd1[4], aCmd2[4]);
                oCmd.Y2 = this.getValBetween(aCmd1[5], aCmd2[5]);
                if(oInvTransform) {
                    XT = oInvTransform.TransformPointX(oCmd.X0, oCmd.Y0);
                    YT = oInvTransform.TransformPointY(oCmd.X0, oCmd.Y0);
                    oCmd.X0 = XT;
                    oCmd.Y0 = YT;
                    XT = oInvTransform.TransformPointX(oCmd.X1, oCmd.Y1);
                    YT = oInvTransform.TransformPointY(oCmd.X1, oCmd.Y1);
                    oCmd.X1 = XT;
                    oCmd.Y1 = YT;
                    XT = oInvTransform.TransformPointX(oCmd.X2, oCmd.Y2);
                    YT = oInvTransform.TransformPointY(oCmd.X2, oCmd.Y2);
                    oCmd.X2 = XT;
                    oCmd.Y2 = YT;
                }
            }
            else if(oCmd.id === AscFormat.close) {

            }
            else {
                return;
            }
        }
    };
    CMorphedPath.prototype.draw = function(oGraphics) {
    };
    CMorphedPath.prototype.isValid = function() {
        return !!this.path;
    };
    CMorphedPath.prototype.getPath = function() {
        return this.path;
    };
    function CComplexMorphObject(oTexturesCache, nRelH1, nRelH2) {
        CMorphObjectBase.call(this, oTexturesCache, nRelH1, nRelH2);
        this.morphedObjects = [];
    }
    AscFormat.InitClassWithoutType(CComplexMorphObject, CMorphObjectBase);
    CComplexMorphObject.prototype.morph = function (dTime) {
        CMorphObjectBase.prototype.morph.call(this, dTime);
        for(let nIdx = 0; nIdx < this.morphedObjects.length; ++ nIdx) {
            this.morphedObjects[nIdx].morph(dTime);
        }
    };
    CComplexMorphObject.prototype.draw = function (oGraphics) {
        for(let nIdx = 0; nIdx < this.morphedObjects.length; ++ nIdx) {
            this.morphedObjects[nIdx].draw(oGraphics);
        }
    };
    CComplexMorphObject.prototype.addMorphObject = function (oMorphObject) {
        this.morphedObjects.push(oMorphObject);
    };


    function CShapeComplexMorph(oTexturesCache, nRelH1, nRelH2, oShape1, oShape2, bNoText) {
        CComplexMorphObject.call(this, oTexturesCache, nRelH1, nRelH2);
        this.shape1 = oShape1;
        this.shape2 = oShape2;
        const oGeometry1 = this.shape1.getMorphGeometry();
        const oGeometry2 = this.shape2.getMorphGeometry();

        const oPlayer1 = this.getPlayer1();
        const oPlayer2 = this.getPlayer2();
        const oDrawData1 = this.getDrawingData(this.shape1, oPlayer1);
        const oDrawData2 = this.getDrawingData(this.shape2, oPlayer2);
        if(!oDrawData1.isVisible && !oDrawData2.isVisible) {
            return;
        }
        if(!oDrawData1.isVisible) {
            this.addMorphObject(new CMorphedAppearObject(oTexturesCache, this.shape2, nRelH2))
            return;
        }
        if(!oDrawData2.isVisible) {
            this.addMorphObject(new CMorphedDisappearObject(oTexturesCache, this.shape1, nRelH1))
            return;
        }

        const oGeometryMorph = new CGeometryMorphObject(this.cache, this.relHeight1, this.relHeight2,
            oGeometry1, oDrawData1.brush, oDrawData1.pen, oDrawData1.transform,
            oGeometry2, oDrawData2.brush, oDrawData2.pen, oDrawData2.transform, !!(oDrawData1.animParams || oDrawData2.animParams || oShape1.group || oShape2.group));
        if(oGeometryMorph.isValid()) {
            this.addMorphObject(oGeometryMorph);
        }
        else {
            this.addMorphObject(new CStretchTextureTransform(oTexturesCache, nRelH1, nRelH2, this.shape1, this.shape2, true));
        }
        if(this.shape1.getObjectType() === AscDFH.historyitem_type_Shape) {
            const oContent1 = this.shape1.getDocContent();
            const oContent2 = this.shape2.getDocContent();
            let bNoText_ = bNoText;
            if(bNoText_) {
                if(oContent1 && oContent1.HaveMaths()) {
                    bNoText_ = false;
                }
                else if(oContent2 && oContent2.HaveMaths()) {
                    bNoText_ = false;
                }
            }
            if(!bNoText_) {
                const oTransform1 = oDrawData1.transformText;
                const oTransform2 = oDrawData2.transformText;
                if(oContent1 || oContent2) {
                    this.addMorphObject(new CContentMorphObject(oTexturesCache, nRelH1, nRelH2,
                        oContent1, oTransform1,
                        oContent2, oTransform2));
                }
            }
        }
    }
    AscFormat.InitClassWithoutType(CShapeComplexMorph, CComplexMorphObject);
    CShapeComplexMorph.prototype.getDrawingData = function(oDrawing, oAnimPlayer) {


        let sFormatId = oDrawing.GetId();
        let oAnimParams = oAnimPlayer.getDrawingParams(sFormatId, true);

        function createParams(oTransform, oTransformText, oBrush, oPen, dOpacity) {
            let bVisible = true;
            if(oAnimParams) {
                bVisible = oAnimParams.isVisible !== false;
            }
            return {
                transform: oTransform,
                transformText: oTransformText,
                brush: oBrush,
                pen: oPen,
                opacity: dOpacity || 0,
                isVisible: bVisible,
                animParams: oAnimParams
            };
        }

        let oTransform, oTransformText, oBrush, oPen, dOpacity;
        if(!oAnimParams) {
            oTransform = oDrawing.transform;
            oTransformText = oDrawing.transformText;
            if(oDrawing.blipFill) {
                oBrush = new AscFormat.CUniFill();
                oBrush.fill = oDrawing.blipFill;
            }
            else {
                oBrush = oDrawing.brush;
            }
            oPen = oDrawing.pen;
            dOpacity = 0;
        }
        else {
            dOpacity = oAnimParams.opacity;
            oBrush = oAnimParams.brush;
            oPen = oAnimParams.pen;
            if(oAnimParams.transform) {
                let oBounds = oDrawing.getBoundsByDrawing(true);
                let oShiftMatrix = new AscCommon.CMatrix();
                AscCommon.global_MatrixTransformer.TranslateAppend(oShiftMatrix, -oBounds.x, -oBounds.y);

                function transformMtx(oMtx) {
                    if(!oMtx) {
                        return null;
                    }
                    let oResult = oMtx.CreateDublicate();
                    AscCommon.global_MatrixTransformer.MultiplyAppend(oResult, oShiftMatrix);
                    AscCommon.global_MatrixTransformer.MultiplyAppend(oResult, oAnimParams.transform);
                    return oResult;
                }
                oTransform = transformMtx(oDrawing.transform);
                oTransformText = transformMtx(oDrawing.transformText);
            }
            else {
                oTransform = oDrawing.transform;
                oTransformText = oDrawing.transformText;
            }
        }

        return createParams(oTransform, oTransformText, oBrush, oPen, dOpacity);
    };

    function CGeometryMorphObject(oTexturesCache, nRelH1, nRelH2,
                                  oGeometry1, oBrush1, oPen1, oTransform1,
                                  oGeometry2, oBrush2, oPen2, oTransform2, bForceNoTransform) {
        CMorphObjectBase.call(this, oTexturesCache, nRelH1, nRelH2);
        this.geometry1 = oGeometry1;
        this.brush1 = oBrush1;
        this.pen1 = oPen1;
        this.transform1 = oTransform1;
        this.geometry2 = oGeometry2;
        this.brush2 = oBrush2;
        this.pen2 = oPen2;
        this.transform2 = oTransform2;
        this.geometry = null;
        this.morphedPaths = [];
        this.textureShape1 = null;
        this.textureShape2 = null;
        this.transformMorph = false;
        this.forceNoTransform = bForceNoTransform;

        this.init();
    }
    AscFormat.InitClassWithoutType(CGeometryMorphObject, CMorphObjectBase);
    CGeometryMorphObject.prototype.init = function() {
        if(!this.geometry1 || !this.geometry2) {
            return;
        }
        if(this.pen1 && this.pen2) {
            const nDash1 = this.pen1.prstDash !== 0 ? this.pen1.prstDash : 6;
            const nDash2 = this.pen2.prstDash !== 0 ? this.pen2.prstDash : 6;
            if(nDash1 !== nDash2) {
                return;
            }
        }
        if(!this.forceNoTransform && this.geometry1.isEqualForMorph(this.geometry2)) {

            let oParent1 =  this.geometry1.parent && this.geometry1.parent.parent;
            let oParent2 =  this.geometry2.parent && this.geometry2.parent.parent;
            if(oParent1 && oParent2 && oParent1.flipH === oParent2.flipH && oParent1.flipV === oParent2.flipV) {
                this.geometry = AscFormat.ExecuteNoHistory(function() {return this.geometry1.createDuplicate();}, this, []);
                this.drawObject = new AscFormat.ObjectToDraw(new AscFormat.CUniFill(), new AscFormat.CLn(), 100, 100, this.geometry, new AscCommon.CMatrix(), 0, 0, null, null);
                this.textureShape1 = CGeometryTextureMorph.prototype.createShape.call(this, AscFormat.ExecuteNoHistory(function () { return new AscFormat.CreateGeometry("rect");}, this, []),
                    this.brush1, AscFormat.CreateNoFillLine(), new AscCommon.CMatrix());
                this.textureShape2 = CGeometryTextureMorph.prototype.createShape.call(this, AscFormat.ExecuteNoHistory(function () { return new AscFormat.CreateGeometry("rect");}, this, []),
                    this.brush2, AscFormat.CreateNoFillLine(), new AscCommon.CMatrix());
                this.transformMorph = true;
                return;
            }
        }
        const aPathLst1 = this.geometry1.pathLst;
        const aPathLst2 = this.geometry2.pathLst;

        let bTextureMorph = true;
        if(aPathLst1.length === aPathLst2.length) {
            const aPaths = [];
            const nPathCount = aPathLst1.length;
            const aMorphs = [];
            let nPath;
            for(nPath = 0; nPath < nPathCount; ++nPath) {
                let oPath1 = aPathLst1[nPath];
                let oPath2 = aPathLst2[nPath];
                let oPathMorph = new CMorphedPath(this.cache, oPath1, this.relHeight1, this.brush1, this.pen1, this.transform1,
                    oPath2, this.relHeight2, this.brush2, this.pen2, this.transform2);
                if(!oPathMorph.isValid()) {
                    break;
                }
                aPaths.push(oPathMorph.getPath());
                aMorphs.push(oPathMorph);
            }
            bTextureMorph = (nPath < nPathCount);
            if(!bTextureMorph) {
                this.morphedPaths = aMorphs;
                this.geometry = AscFormat.ExecuteNoHistory(function () { return new AscFormat.Geometry();}, this, []);
                this.geometry.pathLst = aPaths;
                this.drawObject = new AscFormat.ObjectToDraw(new AscFormat.CUniFill(), new AscFormat.CLn(), 100, 100, this.geometry, new AscCommon.CMatrix(), 0, 0, null, null);
                this.textureShape1 = CGeometryTextureMorph.prototype.createShape.call(this, AscFormat.ExecuteNoHistory(function () { return new AscFormat.CreateGeometry("rect");}, this, []),
                    this.brush1, AscFormat.CreateNoFillLine(), new AscCommon.CMatrix());
                this.textureShape2 = CGeometryTextureMorph.prototype.createShape.call(this, AscFormat.ExecuteNoHistory(function () { return new AscFormat.CreateGeometry("rect");}, this, []),
                    this.brush2, AscFormat.CreateNoFillLine(), new AscCommon.CMatrix());
            }
            return;
        }
    };
    CGeometryMorphObject.prototype.morph = function(dRelTime) {
        if(!this.isValid()) {
            return;
        }
        CMorphObjectBase.prototype.morph.call(this, dRelTime);
        if(this.transformMorph) {
            let oXfrm;
            AscFormat.ExecuteNoHistory(function() {
                oXfrm = new AscFormat.CXfrm();
                let oParent1 = this.geometry1.parent && this.geometry1.parent.parent;
                let oParent2 = this.geometry2.parent && this.geometry2.parent.parent;
                let iN = AscFormat.isRealNumber;
                if(iN(oParent1.x) && iN(oParent2.x)) {
                    oXfrm.offX = this.getValBetween(oParent1.x, oParent2.x);
                }
                if(iN(oParent1.y) && iN(oParent2.y)) {
                    oXfrm.offY = this.getValBetween(oParent1.y, oParent2.y);
                }
                if(iN(oParent1.extX) && iN(oParent2.extX)) {
                    oXfrm.extX = this.getValBetween(oParent1.extX, oParent2.extX);
                }
                if(iN(oParent1.extY) && iN(oParent2.extY)) {
                    oXfrm.extY = this.getValBetween(oParent1.extY, oParent2.extY);
                }
                oXfrm.flipH = oParent1.flipH;
                oXfrm.flipV = oParent1.flipV;
                let nRot1 = AscFormat.normalizeRotate(oParent1.rot || 0);
                let nRot2 = AscFormat.normalizeRotate(oParent2.rot || 0);
                let nAbsDiff1 = Math.abs(nRot2 - nRot1);
                if(nAbsDiff1 <= Math.PI) {
                    oXfrm.rot = this.getValBetween(nRot1, nRot2);
                }
                else {
                    if(nRot1 > nRot2) {
                        oXfrm.rot = AscFormat.normalizeRotate(this.getValBetween(nRot1 - 2*Math.PI, nRot2));
                    }
                    else {

                        oXfrm.rot = AscFormat.normalizeRotate(this.getValBetween(nRot1, nRot2 - 2*Math.PI));
                    }
                }
                const oAvLst = this.geometry1.avLst;
                const oGdLst1 = this.geometry1.gdLst;
                const oGdLst2 = this.geometry2.gdLst;
                for(let sKey in oAvLst) {
                    if(oAvLst.hasOwnProperty(sKey)) {
                        let nVal1 = parseInt(oGdLst1[sKey]);
                        let nVal2 = parseInt(oGdLst2[sKey]);
                        if(AscFormat.isRealNumber(nVal1) && AscFormat.isRealNumber(nVal2)) {
                            this.drawObject.geometry.setAdjValue(sKey, this.getValBetween(nVal1, nVal2) + 0.5 >> 0);
                        }
                    }
                }
            }, this, []);

            const oT = this.drawObject.transform;
            oT.Reset();
            var hc = oXfrm.extX * 0.5;
            var vc = oXfrm.extY * 0.5;
            AscCommon.global_MatrixTransformer.TranslateAppend(oT, -hc, -vc);
            if (oXfrm.flipH)
                AscCommon.global_MatrixTransformer.ScaleAppend(oT, -1, 1);
            if (oXfrm.flipV)
                AscCommon. global_MatrixTransformer.ScaleAppend(oT, 1, -1);
            AscCommon.global_MatrixTransformer.RotateRadAppend(oT, -oXfrm.rot);
            AscCommon.global_MatrixTransformer.TranslateAppend(oT, oXfrm.offX + hc, oXfrm.offY + vc);


            this.drawObject.geometry.Recalculate(oXfrm.extX, oXfrm.extY, false);
            this.drawObject.extX = oXfrm.extX;
            this.drawObject.extY = oXfrm.extY;
            return;
        }
        const oT = this.drawObject.transform;
        const oT1 = this.transform1;
        const oT2 = this.transform2;
        oT.tx = this.getValBetween(oT1.tx, oT2.tx);
        oT.ty = this.getValBetween(oT1.ty, oT2.ty);
        oT.sx = this.getValBetween(oT1.sx, oT2.sx);
        oT.sy = this.getValBetween(oT1.sy, oT2.sy);
        oT.shx = this.getValBetween(oT1.shx, oT2.shx);
        oT.shy = this.getValBetween(oT1.shy, oT2.shy);
        const oInvT = AscCommon.global_MatrixTransformer.Invert(oT);
        const nPathsCount = this.morphedPaths.length;
        for(let nIdx = 0; nIdx < nPathsCount; ++nIdx) {
            this.morphedPaths[nIdx].morph(dRelTime, oInvT);
        }
    };
    CGeometryMorphObject.prototype.morphBrush = function(oBrush1, oBrush2, dScale) {

        if(!oBrush1 && !oBrush2) {
            return null;
        }
        let oBrush = oBrush1;

        if(oBrush1 && oBrush1.isEqual(oBrush2)) {
            return oBrush1;
        }
        const isN = AscFormat.isRealNumber;
        if(oBrush1 && oBrush1.isSolidFill() &&  oBrush2 && oBrush2.isSolidFill()) {
            const oRGBA1 = oBrush1.getRGBAColor();
            const oRGBA2 = oBrush2.getRGBAColor();
            const R = this.getValBetween(oRGBA1.R, oRGBA2.R) + 0.5 >> 0;
            const G = this.getValBetween(oRGBA1.G, oRGBA2.G) + 0.5 >> 0;
            const B = this.getValBetween(oRGBA1.B, oRGBA2.B) + 0.5 >> 0;
            const A = this.getValBetween(oRGBA1.A, oRGBA2.A) + 0.5 >> 0;
            const dTransparent1 = isN(oBrush1.transparent) ? oBrush1.transparent : 255;
            const dTransparent2 = isN(oBrush2.transparent) ? oBrush2.transparent : 255;
            const dTransparent = this.getValBetween(dTransparent1, dTransparent2);
            oBrush = AscFormat.CreateSolidFillRGBA(R, G, B, A);
            oBrush.transparent = dTransparent;
            return oBrush;
        }
        else if(oBrush1 && oBrush1.isNoFill() ||  oBrush2 && oBrush2.isNoFill()) {
            if(this.relTime < 0.5) {
                return oBrush1;
            }
            else {
                return oBrush2;
            }
        }
        else if(oBrush1 && oBrush1.isBlipFill() && oBrush2 && oBrush2.isBlipFill()) {
            const sRasterImageId1 = oBrush1.fill.RasterImageId;
            const sRasterImageId2 = oBrush2.fill.RasterImageId;
            if(sRasterImageId1 === sRasterImageId2) {
                if(!oBrush1.fill.tile && !oBrush2.fill.tile) {

                    oBrush = oBrush1.createDuplicate();
                    if(oBrush1.fill.srcRect || oBrush2.fill.srcRect) {
                        let l1, t1, r1, b1;
                        let l2, t2, r2, b2;
                        if(oBrush1.fill.srcRect) {
                            let oR = oBrush1.fill.srcRect;
                            l1 = isN(oR.l) ? oR.l : 0;
                            t1 = isN(oR.t) ? oR.t : 0;
                            r1 = isN(oR.r) ? oR.r : 100;
                            b1 = isN(oR.b) ? oR.b : 100;
                        }
                        else {
                            l1 = 0;
                            t1 = 0;
                            r1 = 100;
                            b1 = 100;
                        }
                        if(oBrush2.fill.srcRect) {
                            let oR = oBrush2.fill.srcRect;
                            l2 = isN(oR.l) ? oR.l : 0;
                            t2 = isN(oR.t) ? oR.t : 0;
                            r2 = isN(oR.r) ? oR.r : 100;
                            b2 = isN(oR.b) ? oR.b : 100;
                        }
                        else {
                            l2 = 0;
                            t2 = 0;
                            r2 = 100;
                            b2 = 100;
                        }
                        oBrush.fill.srcRect = new AscFormat.CSrcRect();
                        oBrush.fill.srcRect.l = this.getValBetween(l1, l2);
                        oBrush.fill.srcRect.t = this.getValBetween(t1, t2);
                        oBrush.fill.srcRect.r = this.getValBetween(r1, r2);
                        oBrush.fill.srcRect.b = this.getValBetween(b1, b2);
                    }

                    const dTransparent1 = isN(oBrush1.transparent) ? oBrush1.transparent : 255;
                    const dTransparent2 = isN(oBrush2.transparent) ? oBrush2.transparent : 255;
                    const dTransparent = this.getValBetween(dTransparent1, dTransparent2);
                    oBrush.transparent = dTransparent;
                    return oBrush;
                }
            }
        }

        const oShapeDrawer = new AscCommon.CShapeDrawer();
        oShapeDrawer.bIsCheckBounds = true;
        oShapeDrawer.Graphics = new AscFormat.CSlideBoundsChecker();
        this.drawObject.check_bounds(oShapeDrawer);
        const dBoundsW = oShapeDrawer.max_x - oShapeDrawer.min_x;
        const dBoundsH = oShapeDrawer.max_y - oShapeDrawer.min_y;
        this.textureShape1.calcGeometry.Recalculate(dBoundsW, dBoundsH);
        this.textureShape1.brush = oBrush1;
        this.textureShape1.pen = AscFormat.CreateNoFillLine();
        this.textureShape1.bounds.reset(0, 0, dBoundsW, dBoundsH);
        this.textureShape1.extX = dBoundsW;
        this.textureShape1.extY = dBoundsH;
        const oTexture1 = this.cache.checkMorphTexture(this.textureShape1.Id, dScale, oBrush1 && oBrush1.isBlipFill());
        if(!oTexture1) {
            return null;
        }

        this.textureShape2.calcGeometry.Recalculate(dBoundsW, dBoundsH);
        this.textureShape2.brush = oBrush2;
        this.textureShape2.pen = AscFormat.CreateNoFillLine();
        this.textureShape2.bounds.reset(0, 0, dBoundsW, dBoundsH);
        this.textureShape2.extX = dBoundsW;
        this.textureShape2.extY = dBoundsH;
        const oTexture2 = this.cache.checkMorphTexture(this.textureShape2.Id, dScale, oBrush2 && oBrush2.isBlipFill());
        if(!oTexture2) {
            return null;
        }

        oBrush = new AscFormat.CreateBlipFillUniFillFromUrl("");
        oBrush.IsTransitionTextures = true;
        oBrush.alpha1 = 1 - this.relTime;
        oBrush.alpha2 = this.relTime;
        oBrush.canvas1 = oTexture1.canvas;
        oBrush.canvas2 = oTexture2.canvas;
        return oBrush;
    };
    CGeometryMorphObject.prototype.morphPen = function(oPen1, oPen2) {
        if(oPen1 && oPen1.isEqual(oPen2)) {
            return oPen1;
        }
        const  oResultPen1 = oPen1 ? oPen1 : AscFormat.CreateNoFillLine();
        const  oResultPen2 = oPen2 ? oPen2 : AscFormat.CreateNoFillLine();
        const oComparePen = oResultPen1.compare(oResultPen2);
        const oPen = new AscFormat.CLn();
        const isN = AscFormat.isRealNumber;
        const nW1 = isN(oResultPen1.w) ? oResultPen1.w : 12700;
        const nW2 = isN(oResultPen2.w) ? oResultPen2.w : 12700;
        const nW = (this.getValBetween(nW1, nW2) + 0.5) >> 0;
        oPen.w = nW;
        oPen.Fill = this.morphBrush(oResultPen1.Fill, oResultPen2.Fill, 1.0);
        oPen.prstDash = oComparePen.prstDash;
        oPen.Join = oComparePen.Join;
        oPen.headEnd = oComparePen.headEnd;
        oPen.tailEnd = oComparePen.tailEnd;
        oPen.algn = oComparePen.algn;
        oPen.cap = oComparePen.cap;
        oPen.cmpd = oComparePen.cmpd;
        return oPen;
    };
    CGeometryMorphObject.prototype.draw = function(oGraphics) {
        if(!this.isValid()) {
            return;
        }
        const dScale = oGraphics.m_oCoordTransform.sx;
        this.drawObject.brush = this.morphBrush(this.brush1, this.brush2, dScale);
        this.drawObject.pen = this.morphPen(this.pen1, this.pen2);
        this.drawObject.draw(oGraphics);
    };
    CGeometryMorphObject.prototype.isValid = function() {
        return !!this.geometry;
    };

    function CGeometryTextureMorph(oTexturesCache, nRelH1, nRelH2,
                                   oGeometry1, oBrush1, oPen1, oTransform1,
                                   oGeometry2, oBrush2, oPen2, oTransform2) {
        CMorphObjectBase.call(this, oTexturesCache, nRelH1, nRelH2);
        this.shape1 = this.createShape(oGeometry1, oBrush1, oPen1, oTransform1);
        this.shape2 = this.createShape(oGeometry2, oBrush2, oPen2, oTransform2);
    }

    AscFormat.InitClassWithoutType(CGeometryTextureMorph, CMorphObjectBase);
    CGeometryTextureMorph.prototype.createShape = function (oGeometry, oBrush, oPen, oTransform) {
        return AscFormat.ExecuteNoHistory(function() {
            AscCommon.g_oTableId.TurnOn();
            const oShape = new AscFormat.CShape();
            AscCommon.g_oTableId.TurnOff();
            oShape.checkEmptySpPrAndXfrm(null);
            oShape.calcGeometry = oGeometry;
            oShape.spPr.geometry = oGeometry;
            oShape.brush = oBrush;
            oShape.pen = oPen;
            oShape.localTransform = oTransform;
            oShape.transform = oTransform;
            oShape.recalculateBounds();
            return oShape;
        }, this, []);
    };
    CGeometryTextureMorph.prototype.draw = function (oGraphics) {
        const dScale = oGraphics.m_oCoordTransform.sx;
        const oTexture1 = this.cache.checkMorphTexture(this.shape1.GetId(), dScale);
        if(!oTexture1) {
            return;
        }
        const oTexture2 = this.cache.checkMorphTexture(this.shape2.GetId(), dScale);
        if(!oTexture2) {
            return;
        }
        const oBounds1 = this.shape1.bounds;
        const oBounds2 = this.shape2.bounds;
        const oCenter1 = oBounds1.getCenter();
        const oCenter2 = oBounds2.getCenter();
        const dW = this.getValBetween(oBounds1.w, oBounds2.w);
        const dH = this.getValBetween(oBounds1.h, oBounds2.h);
        const dXC = this.getValBetween(oCenter1.x, oCenter2.x);
        const dYC = this.getValBetween(oCenter1.y, oCenter2.y);
        const dX = dXC - dW / 2;
        const dY = dYC - dH / 2;
        const dAlpha1 = 1 - this.relTime;
        const dAlpha2 = this.relTime;
        const oT = oGraphics.m_oCoordTransform;
        const nX = (oT.tx + dX * dScale + 0.5) >> 0;
        const nY = (oT.ty + dY * dScale + 0.5) >> 0;
        const nW = dW * dScale + 0.5 >> 0;
        const nH = dH * dScale + 0.5 >> 0;
        oTexture1.drawInRect(oGraphics, dAlpha1, nX, nY, nW, nH);
        oTexture2.drawInRect(oGraphics, dAlpha2, nX, nY, nW, nH);
    };

    function CContentMorphObject(oTexturesCache, nRelH1, nRelH2,
                                 oContent1, oTransform1,
                                 oContent2, oTransform2) {
        CComplexMorphObject.call(this, oTexturesCache, nRelH1, nRelH2);
        this.transform1 = oTransform1;
        this.transform2 = oTransform2;

        const isN = AscFormat.isRealNumber;

        let oTextDrawer1, oTextDrawer2;
        let oDrawWrapper1, oDrawWrapper2;
        let oDocStruct1, oDocStruct2;
        if(oContent1) {
            oTextDrawer1 = new AscFormat.CTextDrawer(oContent1.XLimit, oContent1.YLimit, false, oContent1.Get_Theme(), true);
            oContent1.Draw(oContent1.StartPage, oTextDrawer1);
            oDocStruct1 = oTextDrawer1.m_oDocContentStructure;
            oDrawWrapper1 = new CTextDrawerStructureWrapper(this.getMorph(), oDocStruct1, oTransform1, oContent1.Get_Theme(), oContent1.Get_ColorMap(), oContent1.Is_DrawingShape(true));
        }
        if(oContent2) {
            oTextDrawer2 = new AscFormat.CTextDrawer(oContent2.XLimit, oContent2.YLimit, false, oContent2.Get_Theme(), true);
            oContent2.Draw(oContent2.StartPage, oTextDrawer2);
            oDocStruct2 = oTextDrawer2.m_oDocContentStructure;
            oDrawWrapper2 = new CTextDrawerStructureWrapper(this.getMorph(), oDocStruct2, oTransform2, oContent2.Get_Theme(), oContent2.Get_ColorMap(), oContent2.Is_DrawingShape(true));
        }
        if(oDrawWrapper1 && !oDrawWrapper2) {
            this.addMorphObject(new CMorphedDisappearObject(oTexturesCache, oDrawWrapper1, nRelH1));
        }
        else if(!oDrawWrapper1 && oDrawWrapper2) {
            this.addMorphObject(new CMorphedAppearObject(oTexturesCache, oDrawWrapper2, nRelH2));
        }
        else if(oDrawWrapper1 && oDrawWrapper2) {
            let bTexture = true;
            if(!oContent1.HaveMaths() && !oContent2.HaveMaths()) {
                const aParStructs1 = oDocStruct1.getParagraphStructures();
                const aParStructs2 = oDocStruct2.getParagraphStructures();
                if(aParStructs1.length === aParStructs2.length) {
                    let nPar;
                    let nStructuresCount = 0;
                    let nMaxCount = 200;
                    for(nPar = 0; nPar < aParStructs1.length; ++nPar) {
                        let oParStruct1 = aParStructs1[nPar];
                        let oParStruct2 = aParStructs2[nPar];
                        let aTextStructs1 = oParStruct1.getTextStructures();
                        let aTextStructs2 = oParStruct2.getTextStructures();
                        if(aTextStructs1.length !== aTextStructs2.length) {
                            break;
                        }
                        nStructuresCount += aTextStructs1.length;
                        if(nStructuresCount > nMaxCount) {
                            break;
                        }
                        let nText;
                        for(nText = 0; nText < aTextStructs1.length; ++nText) {
                            let oTextStruct1 = aTextStructs1[nText];
                            let oTextStruct2 = aTextStructs2[nText];
                            if(!isN(oTextStruct1.Code) || oTextStruct1.Code !== oTextStruct2.Code) {
                                break;
                            }
                        }
                        if(nText < aTextStructs1.length) {
                            break;
                        }
                    }
                    if(nPar === aParStructs1.length && nStructuresCount < nMaxCount) {
                        bTexture = false;


                        for(nPar = 0; nPar < aParStructs1.length; ++nPar) {
                            let oParStruct1 = aParStructs1[nPar];
                            let oParStruct2 = aParStructs2[nPar];
                            let aTextStructs1 = oParStruct1.getTextStructures();
                            let aTextStructs2 = oParStruct2.getTextStructures();
                            if(aTextStructs1.length !== aTextStructs2.length) {
                                break;
                            }


                            let aContent1 = oParStruct1.m_aContent;
                            for(let nIdx = 0; nIdx < aContent1.length; ++nIdx) {
                                let oLine = aContent1[nIdx];
                                let aBg = oLine.m_aBackgrounds;
                                let aBorders = oLine.m_aBorders;
                                let aPBg = oLine.m_aParagraphBackgrounds;
                                let aUS = oLine.m_aUnderlinesStrikeouts;
                                let aDisappear = [].concat(aBg).concat(aBorders).concat(aPBg).concat(aUS);
                                for(let nObj = 0; nObj < aDisappear.length; ++nObj) {
                                    let oWrapper = new CObjectForDrawWrapper(this.getMorph(), aDisappear[nObj], oTransform1, oContent1.Get_Theme(), oContent1.Get_ColorMap(), oContent1.Is_DrawingShape(true));
                                    this.addMorphObject(new CMorphedDisappearObject(oTexturesCache, oWrapper, nRelH1, false))
                                }
                            }
                            let aContent2 = oParStruct2.m_aContent;
                            for(let nIdx = 0; nIdx < aContent2.length; ++nIdx) {
                                let oLine = aContent2[nIdx];
                                let aBg = oLine.m_aBackgrounds;
                                let aBorders = oLine.m_aBorders;
                                let aPBg = oLine.m_aParagraphBackgrounds;
                                let aUS = oLine.m_aUnderlinesStrikeouts;
                                let aAppear = [].concat(aBg).concat(aBorders).concat(aPBg).concat(aUS);
                                for(let nObj = 0; nObj < aAppear.length; ++nObj) {
                                    let oWrapper = new CObjectForDrawWrapper(this.getMorph(), aAppear[nObj], oTransform2, oContent2.Get_Theme(), oContent2.Get_ColorMap(), oContent2.Is_DrawingShape(true));
                                    this.addMorphObject(new CMorphedAppearObject(oTexturesCache, oWrapper, nRelH2, false))
                                }
                            }
                            let nText;
                            for(nText = 0; nText < aTextStructs1.length; ++nText) {
                                let oTextStruct1 = aTextStructs1[nText];
                                let oTextStruct2 = aTextStructs2[nText];
                                if(oTextStruct1.Code === oTextStruct2.Code) {
                                    //let oGeomMorph = new CGeometryMorphObject(oTexturesCache, nRelH1, nRelH2,
                                    //    oTextStruct1.geometry, oTextStruct1.brush, oTextStruct1.pen, oTransform1,
                                    //    oTextStruct2.geometry, oTextStruct2.brush, oTextStruct2.pen, oTransform2);
                                    //if(oGeomMorph.isValid()) {
                                    //    this.addMorphObject(oGeomMorph);
                                    //}
                                    //else {
                                    let oWrapper1 = new CObjectForDrawWrapper(this.getMorph(), oTextStruct1, oTransform1, oContent1.Get_Theme(), oContent1.Get_ColorMap(), oContent1.Is_DrawingShape(true));
                                    let oWrapper2 = new CObjectForDrawWrapper(this.getMorph(), oTextStruct2, oTransform2, oContent2.Get_Theme(), oContent2.Get_ColorMap(), oContent2.Is_DrawingShape(true));
                                    this.addMorphObject(new CStretchTextureTransform(oTexturesCache, nRelH1 + 0.5, nRelH2 + 0.5, oWrapper1, oWrapper2));
                                    //}
                                }
                            }
                        }
                    }
                }
            }
            if(bTexture) {
                this.addMorphObject(new CStretchTextureTransform(oTexturesCache, nRelH1, nRelH2, oDrawWrapper1, oDrawWrapper2));
            }
        }
    }
    AscFormat.InitClassWithoutType(CContentMorphObject, CComplexMorphObject);

    function CMorphedFadeObject(oTexturesCache, oDrawing, nRelH, bNoText) {
        CMorphObjectBase.call(this, oTexturesCache, nRelH, null);
        this.drawing = oDrawing;
        this.bNoText = bNoText;
    }
    AscFormat.InitClassWithoutType(CMorphedFadeObject, CMorphObjectBase);
    CMorphedFadeObject.prototype.morph = function(dRelTime) {
        CMorphObjectBase.prototype.morph.call(this, dRelTime);
    };
    CMorphedFadeObject.prototype.drawWithAlpha = function(oGraphics, dAlpha) {
        const dScale = oGraphics.m_oCoordTransform.sx;
        const oOldTxBody = this.drawing.txBody;
        if(this.bNoText) {
            this.drawing.txBody = null;
        }
        const oTexture = this.cache.checkMorphTexture(this.drawing.GetId(), dScale);
        if(this.bNoText) {
            this.drawing.txBody = oOldTxBody;
        }
        if(!oTexture) {
            return;
        }
        const oFadeTexture = oTexture.createFadeIn(dAlpha);
        if(!oFadeTexture) {
            return;
        }
        oFadeTexture.drawWithoutSaveState(oGraphics, null);
    };

    function CMorphedAppearObject(oTexturesCache, oDrawing, nRelH, bNoText) {
        CMorphedFadeObject.call(this, oTexturesCache, oDrawing, nRelH, bNoText)
    }
    AscFormat.InitClassWithoutType(CMorphedAppearObject, CMorphedFadeObject);
    CMorphedAppearObject.prototype.draw = function(oGraphics) {
        let dAlpha;
        if(this.relTime < 0.5) {
            dAlpha = 0.0;
        }
        else {
            dAlpha = 2 * this.relTime - 1.0;
        }
        this.drawWithAlpha(oGraphics, dAlpha);
    };
    function CMorphedDisappearObject(oTexturesCache, oDrawing, nRelH, bNoText) {
        CMorphedFadeObject.call(this, oTexturesCache, oDrawing, nRelH, bNoText)
    }
    AscFormat.InitClassWithoutType(CMorphedDisappearObject, CMorphedFadeObject);
    CMorphedDisappearObject.prototype.draw = function(oGraphics) {
        let dAlpha;
        if(this.relTime < 0.5) {
            dAlpha = 1.0 - 2 * this.relTime;
        }
        else {
            dAlpha = 0.0;
        }
        this.drawWithAlpha(oGraphics, dAlpha);
    };


    function CStretchTextureTransform(oTexturesCache, nRelH1, nRelH2, oDrawing1, oDrawing2, bNoText) {
        CMorphObjectBase.call(this, oTexturesCache, nRelH1, nRelH2)
        this.drawing1 = oDrawing1;
        this.drawing2 = oDrawing2;
        this.bNoText = !!bNoText;
        this.texture = null;
    }
    AscFormat.InitClassWithoutType(CStretchTextureTransform, CMorphObjectBase);
    CStretchTextureTransform.prototype.getTextureData = function(oDrawing, oAnimPlayer, dScale) {

        function createParams(oTexture, oBounds, dOpacity) {
            if(!oTexture || !oBounds) {
                return null;
            }
            return {
                texture: oTexture,
                bounds: oBounds,
                opacity: dOpacity || 0
            };
        }

        let oOldTextBody = oDrawing.txBody;
        if(this.bNoText) {
            oOldTextBody = oDrawing.txBody;
            oDrawing.txBody = null;
        }
        let sId = oDrawing.GetId();
        let sFormatId = sId;
        let oAnimParams = oAnimPlayer.getDrawingParams(sFormatId, true);
        let oTexture, oBounds, dOpacity;
        if(!oAnimParams) {
            oTexture = this.cache.checkMorphTexture(sId, dScale);
            oBounds = oDrawing.getBoundsByDrawing(true);
        }
        else {

            dOpacity = oAnimParams.opacity;
            oBounds = oDrawing.getBoundsByDrawing(true);
            oBounds.transformRect(oAnimParams.transform);
            oTexture = this.cache.checkMorphTexture(sId, dScale, undefined, oAnimParams);
        }

        if(this.bNoText) {
            oDrawing.txBody = oOldTextBody;
        }
        return createParams(oTexture, oBounds, dOpacity);
    };
    CStretchTextureTransform.prototype.draw = function(oGraphics) {
        const oT = oGraphics.m_oCoordTransform;
        const dScale = oT.sx;

        const oPlayer1 = this.getPlayer1();
        const oPlayer2 = this.getPlayer2();
        const oTextureData1 = this.getTextureData(this.drawing1, oPlayer1, dScale);
        const oTextureData2 = this.getTextureData(this.drawing2, oPlayer2, dScale);

        if(!oTextureData1 || !oTextureData2) {
            return;
        }

        let oTexture1 = oTextureData1.texture, oTexture2 = oTextureData2.texture;
        let oBounds1 = oTextureData1.bounds, oBounds2 = oTextureData2.bounds;
        if(!oTexture1 || !oTexture2) {
            return;
        }
        const dAlpha1 = (1 - this.relTime) * (1 - oTextureData1.opacity);
        const dAlpha2 = this.relTime * (1 - oTextureData2.opacity);
        const nX = (this.getValBetween(oBounds1.x, oBounds2.x) * dScale) + oT.tx >> 0;
        const nY = (this.getValBetween(oBounds1.y, oBounds2.y) * dScale) + oT.ty >> 0;
        const nR = (this.getValBetween(oBounds1.r, oBounds2.r) * dScale) + oT.tx + 0.5 >> 0;
        const nB = (this.getValBetween(oBounds1.b, oBounds2.b) * dScale) + oT.ty + 0.5 >> 0;

        const nW = nR - nX;
        const nH = nB - nY;
        if(!this.texture) {
            const nTextureWidth = Math.max(oTexture1.canvas.width, oTexture2.canvas.width);
            const nTextureHeight = Math.max(oTexture1.canvas.height, oTexture2.canvas.height);
            this.texture = oTexture1.createTexture(nTextureWidth, nTextureHeight);
        }
        const oDrawCanvas = this.texture.canvas;
        if(!this.texture.context) {
            this.texture.context = oDrawCanvas.getContext("2d");
        }
        let oCtx = this.texture.context;
        oCtx.clearRect(0, 0, oDrawCanvas.width, oDrawCanvas.height);
        let sOldOperation = oCtx.globalCompositeOperation;
        oCtx.globalCompositeOperation = "lighter";

        oCtx.globalAlpha = dAlpha1;
        oCtx.drawImage(oTexture1.canvas, 0, 0, oDrawCanvas.width, oDrawCanvas.height);

        oCtx.globalAlpha = dAlpha2;
        oCtx.drawImage(oTexture2.canvas, 0, 0, oDrawCanvas.width, oDrawCanvas.height);
        oCtx.globalCompositeOperation = sOldOperation;
        if(!oGraphics.GetIntegerGrid()) {
            oGraphics.SetIntegerGrid(true);
        }
        oGraphics.m_oContext.drawImage(oDrawCanvas, nX, nY, nW, nH);
    };


    function CGroupComplexMorph(oTexturesCache, nRelH1, nRelH2, oGroup1, oGroup2) {
        CComplexMorphObject.call(this, oTexturesCache, nRelH1, nRelH2);
        let aDrawings1 = oGroup1.arrGraphicObjects;
        let aDrawings2 = oGroup2.arrGraphicObjects;
        this.getMorph().addComparisonMorph(aDrawings1, aDrawings2)
    }
    AscFormat.InitClassWithoutType(CGroupComplexMorph, CComplexMorphObject);

    function CWrapperBase(oMorph, oTransform, oTheme, oColorMap, oFormatDrawing) {
        this.morph = oMorph;
        this.theme = oTheme;
        this.colorMap = oColorMap;
        this.transform = oTransform;
        this.bounds = new AscFormat.CGraphicBounds(0, 0, 0, 0);
        this.Id = null;
        this.formatDrawing = oFormatDrawing;
        this.init();
    }
    CWrapperBase.prototype.GetId = function() {
        return this.Id;
    };
    CWrapperBase.prototype.registerId = function() {
        AscFormat.ExecuteNoHistory(function() {
            this.Id = AscCommon.g_oIdCounter.Get_NewId();
            AscCommon.g_oTableId.TurnOn();
            AscCommon.g_oTableId.Add(this, this.Id);
            AscCommon.g_oTableId.TurnOff();
        }, this, []);
        this.morph.registerWrapperObject(this);
    };
    CWrapperBase.prototype.recalculateBounds = function() {
        const oBoundsChecker = new AscFormat.CSlideBoundsChecker();
        this.draw(oBoundsChecker);
        const oBounds = oBoundsChecker.Bounds;
        this.bounds.reset(oBounds.min_x, oBounds.min_y, oBounds.max_x, oBounds.max_y);
    };
    CWrapperBase.prototype.init = function() {
        this.registerId();
        this.recalculateBounds();
    };
    CWrapperBase.prototype.draw = function(oGraphics) {
    };
    CWrapperBase.prototype.getAnimTexture = function (scale, bMorph, oAnimParams) {
        return AscFormat.CGraphicObjectBase.prototype.getAnimTexture.call(this, scale, bMorph, oAnimParams);
    };
    CWrapperBase.prototype.getBoundsByDrawing = function (bMorph) {
        return this.bounds;
    };
    CWrapperBase.prototype.compareForMorph = function(oDrawingToCheck, oCurCandidate) {
        return oCurCandidate;
    };
    CWrapperBase.prototype.isShape = function () {
        return false;
    };
    CWrapperBase.prototype.getObjectType = function () {
        return null;
    };

    function CTextDrawerStructureWrapper(oMorph, oTextDrawerStructure, oTransform, oTheme, oColorMap, oDrawing) {
        this.textDrawerStructure = oTextDrawerStructure;
        CWrapperBase.call(this, oMorph, oTransform, oTheme, oColorMap, oDrawing);
    }
    AscFormat.InitClassWithoutType(CTextDrawerStructureWrapper, CWrapperBase);
    CTextDrawerStructureWrapper.prototype.draw = function(oGraphics) {
        this.textDrawerStructure.draw(oGraphics, this.transform, this.theme, this.colorMap);
    };

    function CObjectForDrawWrapper(oMorph, oObjectForDraw, oTransform, oTheme, oColorMap, oDrawing) {
        this.objectForDraw = oObjectForDraw;
        CWrapperBase.call(this, oMorph, oTransform, oTheme, oColorMap, oDrawing);
    }
    AscFormat.InitClassWithoutType(CObjectForDrawWrapper, CWrapperBase);
    CObjectForDrawWrapper.prototype.draw = function(oGraphics) {
        this.objectForDraw.draw(oGraphics, undefined, this.transform, this.theme, this.colorMap);
    };
    CObjectForDrawWrapper.prototype.compareForMorph = function(oDrawingToCheck, oCurCandidate) {
        if(!(oDrawingToCheck instanceof CObjectForDrawWrapper)) {
            return oCurCandidate;
        }
        if(this.objectForDraw.compareForMorph(oDrawingToCheck.objectForDraw) !== oDrawingToCheck.objectForDraw) {
            return oCurCandidate;
        }
        return oDrawingToCheck;
    };

    function CObjectForDrawArrayWrapper(oMorph, aObjectForDraw, oTransform, oTheme, oColorMap, oDrawing) {
        this.objectsForDraw = aObjectForDraw;
        CWrapperBase.call(this, oMorph, oTransform, oTheme, oColorMap, oDrawing);
    }
    AscFormat.InitClassWithoutType(CObjectForDrawArrayWrapper, CWrapperBase);
    CObjectForDrawArrayWrapper.prototype.draw = function(oGraphics) {
        for(let nIdx = 0; nIdx < this.objectsForDraw.length; ++nIdx) {
            this.objectsForDraw[nIdx].draw(oGraphics, undefined, this.transform, this.theme, this.colorMap);
        }
    };
    CObjectForDrawArrayWrapper.prototype.compareForMorph = function(oDrawingToCheck, oCurCandidate) {
        if(!(oDrawingToCheck instanceof CObjectForDrawArrayWrapper)) {
            return oCurCandidate;
        }
        if(oDrawingToCheck.objectsForDraw.length !== this.objectsForDraw.length) {
            return oCurCandidate;
        }
        for(let nIdx = 0; nIdx < oDrawingToCheck.objectsForDraw.length; ++nIdx) {
            let oToCheck = oDrawingToCheck.objectsForDraw[nIdx];
            if(oToCheck !== this.objectsForDraw[nIdx].compareForMorph(oToCheck, null)) {
                return oCurCandidate;
            }
        }
        return oDrawingToCheck;
    };

    function CBackgroundWrapper(oMorph, oSlide) {
        this.slide = oSlide;
        CWrapperBase.call(this, oMorph, new AscCommon.CMatrix(), oSlide.getTheme(), oSlide.getColorMap(), null);
    }
    AscFormat.InitClassWithoutType(CBackgroundWrapper, CWrapperBase);
    CBackgroundWrapper.prototype.draw = function(oGraphics) {
        oGraphics.SaveGrState();
        oGraphics.transform3(this.transform);
        this.slide.drawBgMasterAndLayout(oGraphics, true, false);
        oGraphics.RestoreGrState();
    };
    CBackgroundWrapper.prototype.recalculateBounds = function() {
        this.bounds.reset(0, 0, this.slide.Width, this.slide.Height);
    };
    CBackgroundWrapper.prototype.isEqual = function(oWrapper) {
        return this.slide.isEqualBgMasterAndLayout(oWrapper.slide);
    };

    function CTableComplexMorph(oTexturesCache, nRelH1, nRelH2, oGrFrame1, oGrFrame2) {
        CStretchTextureTransform.call(this, oTexturesCache, nRelH1, nRelH2, oGrFrame1, oGrFrame2);
    }
    AscFormat.InitClassWithoutType(CTableComplexMorph, CStretchTextureTransform);

    function CSlideMorphEffect(oSlide1, oSlide2, nType) {
        this.slide1 = oSlide1;
        this.slide2 = oSlide2;
        this.type = c_oAscSlideTransitionParams.Morph_Objects;
        if(AscFormat.isRealNumber(nType)) {
            this.type = nType;
        }
        this.texturesCache = new AscCommon.CTexturesCache();
        this.texturesCache.parent = this;
        this.morphObjects = [];

        this.wrappersMap = {};
        this.init();
    }
    CSlideMorphEffect.prototype.draw = function(oCanvas, oRect) {
        if(!this.slide1 || !this.slide2) {
            return;
        }
        let wPix = oRect.w;
        let hPix = oRect.h;
        let wMM = this.slide1.Width;
        let hMM = this.slide1.Height;
        let oGraphics = new AscCommon.CGraphics();
        let oCtx = oCanvas.getContext('2d');

        oCtx.clearRect(oRect.x, oRect.y, oRect.w, oRect.h);
        oGraphics.init(oCtx, wPix, hPix, wMM, hMM);
        oGraphics.m_oCoordTransform.tx = oRect.x;
        oGraphics.m_oCoordTransform.ty = oRect.y;
        oGraphics.m_oFontManager = AscCommon.g_fontManager;
        oGraphics.transform(1, 0, 0, 1, 0, 0);
        oGraphics.IsNoDrawingEmptyPlaceholder = true;
        oGraphics.IsDemonstrationMode = true;

        oGraphics.SaveGrState();
        oGraphics.AddClipRect(0, 0, wMM, hMM);
        DrawBackground(oGraphics, AscFormat.CreateSolidFillRGBA(255, 255, 255, 255), wMM, hMM);
        for(let nIdx = 0; nIdx < this.morphObjects.length; ++nIdx) {
            this.morphObjects[nIdx].draw(oGraphics);
        }
        oGraphics.RestoreGrState();
    };
    CSlideMorphEffect.prototype.init = function() {
        if(!this.slide1 || !this.slide2) {
            return;
        }
        AscFormat.ExecuteNoHistory(function() {

            let oTiming = this.slide1.timing;

            if(oTiming) {
                this.slide1.timing = oTiming.createDuplicate();
            }
            this.player1 = new AscFormat.CAnimationPlayer(this.slide1, null);
            this.player1.goToEnd();
            this.player1.pause();
            this.slide1.timing = oTiming;

            oTiming = this.slide2.timing;
            if(oTiming) {
                this.slide2.timing = oTiming.createDuplicate();
            }
            this.player2 = new AscFormat.CAnimationPlayer(this.slide2, null);
            this.player2.start();
            this.player2.pause();

            this.slide2.timing = oTiming;
            switch(this.type) {
                case c_oAscSlideTransitionParams.Morph_Words: {
                    this.generateWordBasedMorphs();
                    break;
                }
                case c_oAscSlideTransitionParams.Morph_Letters: {
                    this.generateLetterBasedMorphs();
                    break;
                }
                default: {
                    this.generateObjectBasedMorphs();
                    break;
                }
            }
        }, this, []);
    };
    CSlideMorphEffect.prototype.pushMorphObject = function (oMorph) {
        this.morphObjects.push(oMorph);
    };
    CSlideMorphEffect.prototype.addShapeMorphs = function (oShape1, nRelH1, oShape2, nRelH2, bNoText) {
        this.pushMorphObject(new CShapeComplexMorph(this.texturesCache, nRelH1, nRelH2, oShape1, oShape2, bNoText));
    };
    CSlideMorphEffect.prototype.addTableMorphs = function (oGrFrame1, nRelH1, oGrFrame2, nRelH2) {
        this.pushMorphObject(new CTableComplexMorph(this.texturesCache, nRelH1, nRelH2, oGrFrame1, oGrFrame2));
    };
    CSlideMorphEffect.prototype.addGroupMorphs = function (oGroup1, nRelH1, oGroup2, nRelH2) {
        this.pushMorphObject(new CGroupComplexMorph(this.texturesCache, nRelH1, nRelH2, oGroup1, oGroup2));
    };
    CSlideMorphEffect.prototype.addObjectMorphs = function(oDrawing1, nRelH1, oDrawing2, nRelH2, bNoText) {
        if(!oDrawing1 || !oDrawing2) {
            return;
        }
        const nType1 = oDrawing1.getObjectType();
        const nType2 = oDrawing2.getObjectType();
        if(nType1 !== nType2) {
            return;
        }
        switch (nType1) {
            case AscDFH.historyitem_type_Shape:
            case AscDFH.historyitem_type_Cnx:
            case AscDFH.historyitem_type_ImageShape: {
                this.addShapeMorphs(oDrawing1, nRelH1, oDrawing2, nRelH2, bNoText)
                break;
            }
            case AscDFH.historyitem_type_GraphicFrame: {
                this.addTableMorphs(oDrawing1, nRelH1, oDrawing2, nRelH2);
                break;
            }
            case AscDFH.historyitem_type_GroupShape:
            case AscDFH.historyitem_type_SmartArt: {
                this.addGroupMorphs(oDrawing1, nRelH1, oDrawing2, nRelH2);
                break;
            }
            default: {
                this.pushMorphObject(new CStretchTextureTransform(this.texturesCache, nRelH1, nRelH2, oDrawing1, oDrawing2));
                break;
            }
        }
    };
    CSlideMorphEffect.prototype.isDrawingHidden = function(oDrawing, oPlayer) {
        if(!oPlayer) {
            return false;
        }
        return !oPlayer.isDrawingVisible(GetDrawingFormatId(oDrawing));
    };
    CSlideMorphEffect.prototype.addBackgroundMorph = function() {
        this.pushMorphObject(new CStretchTextureTransform(this.texturesCache, -1, -1, new CBackgroundWrapper(this, this.slide1), new CBackgroundWrapper(this, this.slide2)));
    };
    CSlideMorphEffect.prototype.generateObjectBasedMorphs = function() {

        //match objects
        this.addBackgroundMorph();
        const aDrawings1 = this.slide1.getDrawingObjects();
        const aDrawings2 = this.slide2.getDrawingObjects();
        const nDrawingsCount1 = aDrawings1.length;
        const nDrawingsCount2 = aDrawings2.length;
        const oMapPaired = {};
        let oPairedDrawing, oDrawing1;
        let oCandidates = {};
        for(let nDrawing1 = 0; nDrawing1 < nDrawingsCount1; ++nDrawing1) {
            oDrawing1 = aDrawings1[nDrawing1];
            if(this.isDrawingHidden(oDrawing1, this.player1)) {
                continue;
            }
            oCandidates[oDrawing1.Id] = [];
            oPairedDrawing = null;
            let nParedRelH = null;
            for(let nDrawing2 = 0; nDrawing2 < nDrawingsCount2; ++nDrawing2) {
                let oDrawing2 = aDrawings2[nDrawing2];
                if(this.isDrawingHidden(oDrawing2, this.player2)) {
                    continue;
                }
                oPairedDrawing = oDrawing1.compareForMorph(oDrawing2, oPairedDrawing, oMapPaired);
                if(oDrawing2 === oPairedDrawing) {
                    nParedRelH = nDrawing2;
                    oCandidates[oDrawing1.Id].push([oPairedDrawing, nDrawing1]);
                }
            }
            if(oPairedDrawing) {
                if(oMapPaired[oPairedDrawing.Id]) {
                    let oOldDrawing1 = oMapPaired[oPairedDrawing.Id].drawing;
                    let aCandidates = oCandidates[oOldDrawing1.Id];
                    for(let nCandidate = 0; nCandidate < aCandidates.length; ++nCandidate) {
                        let aPair = aCandidates[nCandidate];
                        let sPairedId = aPair[0].Id;
                        let nOld1RelH = aPair[1];
                        if(!oMapPaired[sPairedId]) {
                            oMapPaired[sPairedId] = {drawing: oOldDrawing1, relH: nOld1RelH};
                        }
                    }

                }
                oMapPaired[oPairedDrawing.Id] = {drawing: oDrawing1, relH: nDrawing1};
            }
        }
        for(let nDrawing2 = 0; nDrawing2 < nDrawingsCount2; ++nDrawing2) {
            let oDrawing2 = aDrawings2[nDrawing2];
            if(this.isDrawingHidden(oDrawing2, this.player2)) {
                continue;
            }
            let oParedObj = oMapPaired[oDrawing2.Id];
            if(oParedObj) {
                this.addObjectMorphs(oParedObj.drawing, oParedObj.relH, oDrawing2, nDrawing2);
            }
            else {
                this.pushMorphObject(new CMorphedAppearObject(this.texturesCache, oDrawing2, nDrawing2));
            }
        }
        for(let nDrawing1 = 0; nDrawing1 < nDrawingsCount1; ++nDrawing1) {
            oDrawing1 = aDrawings1[nDrawing1];
            if(this.isDrawingHidden(oDrawing1, this.player1)) {
                continue;
            }
            let bDisappear = true;
            for(let sKey in oMapPaired) {
                if(oMapPaired.hasOwnProperty(sKey)) {
                    if(oMapPaired[sKey].drawing === oDrawing1) {
                        bDisappear = false;
                    }
                }
            }
            if(bDisappear) {
                this.pushMorphObject(new CMorphedDisappearObject(this.texturesCache, oDrawing1, nDrawing1));
            }
        }
    };
    CSlideMorphEffect.prototype.generateWordBasedMorphs = function() {
        this.generateTextBasedMorph(false);
    };
    CSlideMorphEffect.prototype.generateLetterBasedMorphs = function() {
        this.generateTextBasedMorph(true);
    };
    CSlideMorphEffect.prototype.generateTextBasedMorph = function(bLetter) {
        this.addBackgroundMorph();
        const aDrawings1 = this.slide1.getDrawingObjects();
        const aDrawings2 = this.slide2.getDrawingObjects();
        const aMorphedDrawings1 = this.createMatchArray(aDrawings1, bLetter);
        const aMorphedDrawings2 = this.createMatchArray(aDrawings2, bLetter);
        this.addComparisonMorph(aMorphedDrawings1, aMorphedDrawings2);
        //------------------------

    };
    CSlideMorphEffect.prototype.addComparisonMorph = function (aMorphedDrawings1, aMorphedDrawings2) {
        const oCompareResult = compareDrawings(aMorphedDrawings1, aMorphedDrawings2);
        const aBaseNode = oCompareResult[0].children;
        const aCompareNode = oCompareResult[1].children;
        for(let nIdx = 0; nIdx < aBaseNode.length; ++nIdx) {
            let oBaseNode = aBaseNode[nIdx];
            let oPartner = oBaseNode.partner;
            if(oBaseNode.partner) {
                this.addObjectMorphs(oBaseNode.element, oBaseNode.idx, oPartner.element, oPartner.idx, true);
            }
            else {
                this.pushMorphObject(new CMorphedDisappearObject(this.texturesCache, oBaseNode.element, oBaseNode.idx, true));
            }
        }
        for(let nIdx = 0; nIdx < aCompareNode.length; ++nIdx) {
            let oNode = aCompareNode[nIdx];
            if(!oNode.partner) {
                this.pushMorphObject(new CMorphedAppearObject(this.texturesCache, oNode.element, oNode.idx, true));
            }
        }
    };
    CSlideMorphEffect.prototype.createMatchArray = function(aDrawings, bLetter) {
        let aRet = [];
        for(let nSp = 0; nSp < aDrawings.length; ++nSp) {
            const oSp = aDrawings[nSp];
            let nType = oSp.getObjectType();
            switch (nType) {
                case AscDFH.historyitem_type_Shape: {
                    aRet.push(oSp);
                    let oDocContent = oSp.getDocContent();
                    if(oDocContent && !oDocContent.HaveMaths()) {
                        const oTextDrawer = new AscFormat.CTextDrawer(oDocContent.XLimit, oDocContent.YLimit, false, oDocContent.Get_Theme(), true);
                        oDocContent.Draw(oDocContent.StartPage, oTextDrawer);
                        const oDocStruct = oTextDrawer.m_oDocContentStructure;
                        let nIdx;
                        let oTheme = oDocContent.Get_Theme();
                        let oColorMap = oDocContent.Get_ColorMap();
                        let oTransform = oSp.transformText;
                        let oObjectForDraw;
                        for(nIdx = 0; nIdx < oDocStruct.m_aParagraphBackgrounds.length; ++nIdx)
                        {
                            oObjectForDraw = oDocStruct.m_aParagraphBackgrounds[nIdx];
                            aRet.push(new CObjectForDrawWrapper(this, oObjectForDraw, oTransform, oTheme, oColorMap, oSp));
                        }
                        for(nIdx = 0;nIdx< oDocStruct.m_aBorders.length; ++nIdx)
                        {
                            oObjectForDraw = oDocStruct.m_aBorders[nIdx];
                            aRet.push(new CObjectForDrawWrapper(this, oObjectForDraw, oTransform, oTheme, oColorMap, oSp));
                        }
                        for(nIdx = 0; nIdx < oDocStruct.m_aBackgrounds.length; ++nIdx)
                        {
                            oObjectForDraw = oDocStruct.m_aBackgrounds[nIdx];
                            aRet.push(new CObjectForDrawWrapper(this, oObjectForDraw, oTransform, oTheme, oColorMap, oSp));
                        }

                        for(nIdx = 0; nIdx < oDocStruct.m_aContent.length; ++nIdx) {
                            let oPara = oDocStruct.m_aContent[nIdx];
                            let aWords = oPara.m_aWords;
                            for(let nWord = 0; nWord < aWords.length; ++nWord) {
                                let aWord = aWords[nWord];
                                if(bLetter) {
                                    for(let nLetter = 0; nLetter < aWord.length; ++nLetter) {
                                        aRet.push( new CObjectForDrawWrapper(this, aWord[nLetter], oTransform, oTheme, oColorMap, oSp));
                                    }
                                }
                                else {
                                    aRet.push( new CObjectForDrawArrayWrapper(this, aWord, oTransform, oTheme, oColorMap, oSp));
                                }
                            }
                        }
                    }
                    break;
                }
                default: {
                    aRet.push(oSp);
                    break;
                }
            }
        }
        return aRet;
    };
    CSlideMorphEffect.prototype.morph = function(t) {
        for(let nIdx = 0; nIdx < this.morphObjects.length; ++nIdx) {
            this.morphObjects[nIdx].morph(t);
        }
        this.morphObjects.sort(function (a, b) {
            return a.relHeight - b.relHeight;
        });
    };
    CSlideMorphEffect.prototype.end = function () {
        this.player1.stop();
        this.player2.stop();
        this.clearWrappers();
        this.texturesCache.clear();
        this.morphObjects.length = 0;
    };
    CSlideMorphEffect.prototype.registerWrapperObject = function(oWrapper) {
        if(!oWrapper) {
            return;
        }
        this.wrappersMap[oWrapper.GetId()] = oWrapper;
    };
    CSlideMorphEffect.prototype.clearWrappers = function() {
        for(let sId in this.wrappersMap) {
            if(this.wrappersMap.hasOwnProperty(sId)) {
                AscCommon.g_oTableId.Delete(sId);
            }
        }
        this.wrappersMap = {};
    };
    window['AscCommonSlide'] = window['AscCommonSlide'] || {};
    window['AscCommonSlide'].CSlideMorphEffect = CSlideMorphEffect;
}) (window);
