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

(function(){

    /**
	 * Class representing a Ink annotation.
	 * @constructor
    */
    function CAnnotationInk(sName, nPage, aRect, oDoc)
    {
        AscPDF.CPdfShape.call(this);
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Ink, nPage, aRect, oDoc);
        
        AscCommon.History.StartNoHistoryMode();
        initShape(this);
        AscCommon.History.EndNoHistoryMode();

        this._state         = undefined;
        this._stateModel    = undefined;
        this._gestures      = [];
        this._relativePaths = [];
        this._width         = 1;
    }
    
	CAnnotationInk.prototype.constructor = CAnnotationInk;
    AscFormat.InitClass(CAnnotationInk, AscPDF.CPdfShape, AscDFH.historyitem_type_Pdf_Annot_Ink);
    Object.assign(CAnnotationInk.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationInk.prototype.IsInk = function() {
        return true;
    };
    CAnnotationInk.prototype.GetDrawing = function() {
        return this.content.GetAllDrawingObjects()[0];
    };

    CAnnotationInk.prototype.SetInkPoints = function(aSourcePaths) {
        let oThis = this;
        aSourcePaths.forEach(function(aPath) {
            oThis.AddInkPath(aPath);
        });
    };
    CAnnotationInk.prototype.AddInkPath = function(aInkPath) {
        AscCommon.History.Add(new CChangesPDFInkPoints(this, this._gestures.length, aInkPath, true));
        this._gestures.push(aInkPath);

        let oViewer = Asc.editor.getDocumentRenderer();
        if (false == oViewer.IsOpenAnnotsInProgress) {
            this.SetRect(this.private_CalculateBoundingBox());
        }
        
        this.SetWasChanged(true);
        this.recalcGeometry();
        this.SetNeedRecalc(true);
    };
    CAnnotationInk.prototype.private_CalculateBoundingBox = function() {
        if (this._gestures.length === 0) {
            return null;
        }
    
        let nLineW = this.GetWidth();

        let minX = Infinity;
        let minY = Infinity;
        let maxX = -Infinity;
        let maxY = -Infinity;
    
        for (let i = 0; i < this._gestures.length; i++) {
            let path = this._gestures[i];
            for (let j = 0; j < path.length; j += 2) {
                let x = path[j];
                let y = path[j + 1];
                if (x < minX) {
                    minX = x;
                }
                if (y < minY) {
                    minY = y;
                }
                if (x > maxX) {
                    maxX = x;
                }
                if (y > maxY) {
                    maxY = y;
                }
            }
        }
    
        return [minX - nLineW, minY - nLineW, maxX + nLineW, maxY + nLineW]
    };
    CAnnotationInk.prototype.GetInkPoints = function() {
        return this._gestures;
    };
    CAnnotationInk.prototype.SetRect = function(aOrigRect) {
        let oViewer     = editor.getDocumentRenderer();
        let oDoc        = oViewer.getPDFDoc();

        oDoc.History.Add(new CChangesPDFAnnotRect(this, this.GetOrigRect(), aOrigRect));

        this._origRect = aOrigRect;

        let oXfrm = this.getXfrm();
        if (oXfrm) {
            AscCommon.History.StartNoHistoryMode();

            let nX1 = aOrigRect[0] * g_dKoef_pt_to_mm;
            let nX2 = aOrigRect[2] * g_dKoef_pt_to_mm;
            let nY1 = aOrigRect[1] * g_dKoef_pt_to_mm;
            let nY2 = aOrigRect[3] * g_dKoef_pt_to_mm;

            this.spPr.xfrm.setExtX(nX2 - nX1);
            this.spPr.xfrm.setExtY(nY2 - nY1);
            this.spPr.xfrm.setOffX(nX1);
            this.spPr.xfrm.setOffY(nY1);
            
            this.SetNeedRecalc(true);
            this.RefillGeometry(this.spPr.geometry, [nX1, nY1, nX2, nY2]);

            AscCommon.History.EndNoHistoryMode();
        }
        
        this.SetWasChanged(true);
    };
    CAnnotationInk.prototype.SetFlipV = function(bFlip) {
        let oDoc = this.GetDocument();
        return;
        
        if (this.flipV != bFlip) {
            oDoc.History.Add(new CChangesPDFInkFlipV(this, this.flipV, bFlip));
            this.changeFlipV(!this.flipV);
            this.recalculate();
        }
    };
    CAnnotationInk.prototype.SetFlipH = function(bFlip) {
        let oDoc = this.GetDocument();
        return;
        
        if (this.flipH != bFlip) {
            oDoc.History.Add(new CChangesPDFInkFlipH(this, this.flipV, bFlip));
            this.changeFlipH(!this.flipH);
            this.recalculate();
        }
    };
    CAnnotationInk.prototype.Recalculate = function(bForce) {
        if (true !== bForce && false == this.IsNeedRecalc()) {
            return;
        }

        if (this._relativePaths.length != this._gestures.length) {
            this.InitGeometry();
            return;
        }
        
        this.recalculateTransform();
        this.updateTransformMatrix();
        this.recalculate();
        this.SetNeedRecalc(false);
    };
    CAnnotationInk.prototype.InitGeometry = function() {
        let aSourcePaths = this._gestures;

        let aShapePaths = [];
        for (let nPath = 0; nPath < aSourcePaths.length; nPath++) {
            let aSourcePath = aSourcePaths[nPath];
            let aShapePath  = [];
            
            for (let i = 0; i < aSourcePath.length - 1; i += 2) {
                aShapePath.push({
                    x: aSourcePath[i] * g_dKoef_pt_to_mm,
                    y: (aSourcePath[i + 1]) * g_dKoef_pt_to_mm
                });
            }

            aShapePaths.push(aShapePath);
        }
        
        let aShapeRectInMM = this.GetOrigRect().map(function(measure) {
            return measure * g_dKoef_pt_to_mm;
        });

        fillShapeByPoints(aShapePaths, aShapeRectInMM, this);
        
        let aRelPointsPos   = [];
        let aAllPoints      = [];

        for (let i = 0; i < aShapePaths.length; i++)
            aAllPoints = aAllPoints.concat(aShapePaths[i]);

        let aMinRect = getMinRect(aAllPoints);
        let xMin = aMinRect[0];
        let yMin = aMinRect[1];
        let xMax = aMinRect[2];
        let yMax = aMinRect[3];

        // считаем относительное положение точек внутри фигуры
        for (let nPath = 0; nPath < aShapePaths.length; nPath++) {
            let aPoints         = aShapePaths[nPath]
            let aTmpRelPoints   = [];
            
            for (let nPoint = 0; nPoint < aPoints.length; nPoint++) {
                let oPoint = aPoints[nPoint];

                let nIndX = oPoint.x - xMin;
                let nIndY = oPoint.y - yMin;

                aTmpRelPoints.push({
                    relX: nIndX / (xMax - xMin),
                    relY: nIndY / (yMax - yMin)
                });
            }
            
            aRelPointsPos.push(aTmpRelPoints);
        }
        
        this._relativePaths = aRelPointsPos;

        this.Recalculate(true);
    };
    CAnnotationInk.prototype.RefillGeometry = function(oGeometry, aBounds) {
        if (!this._relativePaths) {
            return;
        }

        let aRelPointsPos   = this._relativePaths;
        let aShapePaths     = [];
        
        let nLineW = this.GetWidth() * g_dKoef_pt_to_mm;

        let xMin = aBounds[0] + nLineW;
        let yMin = aBounds[1] + nLineW;
        let xMax = aBounds[2] - nLineW;
        let yMax = aBounds[3] - nLineW;

        let nWidthMM    = (xMax - xMin);
        let nHeightMM   = (yMax - yMin);

        for (let nPath = 0; nPath < aRelPointsPos.length; nPath++) {
            let aPath       = aRelPointsPos[nPath];
            let aShapePath  = [];

            for (let nPoint = 0; nPoint < aPath.length; nPoint++) {
                aShapePath.push({
                    x: (nWidthMM) * aPath[nPoint].relX + xMin,
                    y: (nHeightMM) * aPath[nPoint].relY + yMin
                });
            }
            
            aShapePaths.push(aShapePath);
        }
        
        let geometry = generateGeometry(aShapePaths, aBounds, oGeometry);
        this.recalcTransform()
        var transform = this.getTransform();
        
        geometry.Recalculate(transform.extX, transform.extY);

        return geometry;
    };
    CAnnotationInk.prototype.LazyCopy = function() {
        let oDoc = this.GetDocument();
        oDoc.StartNoHistoryMode();

        let oNewInk = new CAnnotationInk(AscCommon.CreateGUID(), this.GetPage(), this.GetOrigRect().slice(), oDoc);

        oNewInk.lazyCopy = true;

        this.fillObject(oNewInk);

        let aStrokeColor = this.GetStrokeColor();

        oNewInk._apIdx = this._apIdx;
        oNewInk._originView = this._originView;
        oNewInk.SetOriginPage(this.GetOriginPage());
        oNewInk.SetAuthor(this.GetAuthor());
        oNewInk.SetModDate(this.GetModDate());
        oNewInk.SetCreationDate(this.GetCreationDate());
        aStrokeColor && oNewInk.SetStrokeColor(aStrokeColor.slice());
        oNewInk.SetWidth(this.GetWidth());
        oNewInk.SetOpacity(this.GetOpacity());
        oNewInk._relativePaths = this.GetRelativePaths().slice();
        oNewInk._gestures = this._gestures.slice();
        oNewInk.SetContents(this.GetContents());
        oNewInk.recalcGeometry();

        oDoc.EndNoHistoryMode();
        return oNewInk;
    };
    CAnnotationInk.prototype.GetRelativePaths = function() {
        return this._relativePaths;
    };
    
    CAnnotationInk.prototype.IsSelected = function() {
        let oViewer         = editor.getDocumentRenderer();
        let oDrawingObjects = oViewer.DrawingObjects;
        return oDrawingObjects.selectedObjects.includes(this);
    };
        
    CAnnotationInk.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        let aSourcePoints = this.GetInkPoints();

        memory.WriteLong(aSourcePoints.length);
        for (let i = 0; i < aSourcePoints.length; i++) {
            memory.WriteLong(aSourcePoints[i].length);

            for (let j = 0; j < aSourcePoints[i].length - 1; j+=2) {
                memory.WriteDouble(aSourcePoints[i][j]);
                memory.WriteDouble(aSourcePoints[i][j + 1]);
            }
        }

        let nEndPos = memory.GetCurPosition();
        memory.Seek(memory.posForFlags);
        memory.WriteLong(memory.annotFlags);
        
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };
    
    function fillShapeByPoints(arrOfArrPoints, aShapeRect, oParentAnnot) {
        let xMax = aShapeRect[2];
        let xMin = aShapeRect[0];
        let yMin = aShapeRect[1];
        let yMax = aShapeRect[3];

        let geometry = generateGeometry(arrOfArrPoints, [xMin, yMin, xMax, yMax], oParentAnnot.spPr.geometry);
        if (!oParentAnnot.spPr.geometry) {
            oParentAnnot.spPr.setGeometry(geometry);
        }

        let oXfrm = oParentAnnot.getXfrm();
        oXfrm.offX = xMin;
        oXfrm.offY = yMin;

        return oParentAnnot;
    }

    function initShape(oParentAnnot) {
        let aOrigRect = oParentAnnot.GetOrigRect();
        let aShapeRectInMM = aOrigRect ? aOrigRect.map(function(measure) {
            return measure * g_dKoef_pt_to_mm;
        }) : [];
        let xMin = aShapeRectInMM[0];
        let xMax = aShapeRectInMM[2];
        let yMin = aShapeRectInMM[1];
        let yMax = aShapeRectInMM[3];

        oParentAnnot.setSpPr(new AscFormat.CSpPr());
        oParentAnnot.spPr.setLn(new AscFormat.CLn());
        oParentAnnot.spPr.ln.setFill(AscFormat.CreateNoFillUniFill());
        oParentAnnot.spPr.setFill(AscFormat.CreateNoFillUniFill());
        oParentAnnot.spPr.setParent(oParentAnnot);
        oParentAnnot.spPr.setXfrm(new AscFormat.CXfrm());
        oParentAnnot.spPr.xfrm.setParent(oParentAnnot.spPr);
        
        oParentAnnot.spPr.xfrm.setOffX(xMin);
        oParentAnnot.spPr.xfrm.setOffY(yMin);
        oParentAnnot.spPr.xfrm.setExtX(Math.abs(xMax - xMin));
        oParentAnnot.spPr.xfrm.setExtY(Math.abs(yMax - yMin));
        oParentAnnot.setStyle(AscFormat.CreateDefaultShapeStyle());
        oParentAnnot.setBDeleted(false);
        oParentAnnot.recalcInfo.recalculateGeometry = false;
        oParentAnnot.recalculate();
        oParentAnnot.brush = AscFormat.CreateNoFillUniFill();
    }
    function generateGeometry(arrOfArrPoints, aBounds, oGeometry) {
        let xMin = aBounds[0];
        let yMin = aBounds[1];
        let xMax = aBounds[2];
        let yMax = aBounds[3];

        let geometry = oGeometry ? oGeometry : new AscFormat.Geometry();
        if (oGeometry) {
            oGeometry.pathLst = [];
        }

        for (let nPath = 0; nPath < arrOfArrPoints.length; nPath++) {
            let bClosed     = false;
            let aPoints     = arrOfArrPoints[nPath];
            let min_dist    = editor.WordControl.m_oDrawingDocument.GetMMPerDot(3);
            let oLastPoint  = aPoints[aPoints.length-1];
            let nLastIndex  = aPoints.length-1;
            if(oLastPoint.bTemporary) {
                nLastIndex--;
            }
            if(nLastIndex > 1)
            {
                let dx = aPoints[0].x - aPoints[nLastIndex].x;
                let dy = aPoints[0].y - aPoints[nLastIndex].y;
                if(Math.sqrt(dx*dx +dy*dy) < min_dist)
                {
                    bClosed = true;
                }
            }
            let nMaxPtIdx = bClosed ? (nLastIndex - 1) : nLastIndex;

            let w = xMax - xMin, h = yMax-yMin;
            let kw, kh, pathW, pathH;
            if(w > 0)
            {
                pathW = 43200;
                kw = 43200/ w;
            }
            else
            {
                pathW = 0;
                kw = 0;
            }
            if(h > 0)
            {
                pathH = 43200;
                kh = 43200 / h;
            }
            else
            {
                pathH = 0;
                kh = 0;
            }
            geometry.AddPathCommand(0, undefined, bClosed ? "norm": "none", undefined, pathW, pathH);
            geometry.AddRect("l", "t", "r", "b");
            geometry.AddPathCommand(1, (((aPoints[0].x - xMin) * kw) >> 0) + "", (((aPoints[0].y - yMin) * kh) >> 0) + "");
            let i = 1;
            let aRanges = [[0, aPoints.length - 1]];
            let aRange, nRange;
            let nEnd;
            let nPtsCount = aPoints.length;
            let oPt1, oPt2, oPt3, nPt;
            for(nRange = 0; nRange < aRanges.length; ++nRange)
            {
                aRange = aRanges[nRange];
                if(aRange[0] + 1 > nMaxPtIdx) {
                    break;
                }
                nPt = aRange[0] + 1;
                nEnd = Math.min(aRange[1], nMaxPtIdx);
                while(nPt <= nEnd)
                {
                    if(nPt + 2 <= nEnd)
                    {
                        //cubic bezier curve
                        oPt1 = aPoints[nPt++];
                        oPt2 = aPoints[nPt++];
                        oPt3 = aPoints[nPt++];
                        geometry.AddPathCommand(5,
                            (((oPt1.x - xMin) * kw) >> 0) + "", (((oPt1.y - yMin) * kh) >> 0) + "",
                            (((oPt2.x - xMin) * kw) >> 0) + "", (((oPt2.y - yMin) * kh) >> 0) + "",
                            (((oPt3.x - xMin) * kw) >> 0) + "", (((oPt3.y - yMin) * kh) >> 0) + ""
                        );
                    }
                    else if(nPt + 1 <= nEnd)
                    {
                        //quad bezier curve
                        oPt1 = aPoints[nPt++];
                        oPt2 = aPoints[nPt++];
                        geometry.AddPathCommand(4,
                            (((oPt1.x - xMin) * kw) >> 0) + "", (((oPt1.y - yMin) * kh) >> 0) + "",
                            (((oPt2.x - xMin) * kw) >> 0) + "", (((oPt2.y - yMin) * kh) >> 0) + ""
                        );
                    }
                    else
                    {
                        //lineTo
                        oPt1 = aPoints[nPt++];
                        geometry.AddPathCommand(2,
                            (((oPt1.x - xMin) * kw) >> 0) + "", (((oPt1.y - yMin) * kh) >> 0) + ""
                        );
                    }
                }
            }
            if(bClosed)
            {
                geometry.AddPathCommand(6);
            }
        }
        

        return geometry;
    }

    function getMinRect(aPoints) {
        let xMax = aPoints[0].x, yMax = aPoints[0].y, xMin = xMax, yMin = yMax;
        for(let i = 1; i < aPoints.length; i++)
        {
            if(aPoints[i].x > xMax)
            {
                xMax = aPoints[i].x;
            }
            if(aPoints[i].y > yMax)
            {
                yMax = aPoints[i].y;
            }

            if(aPoints[i].x < xMin)
            {
                xMin = aPoints[i].x;
            }

            if(aPoints[i].y < yMin)
            {
                yMin = aPoints[i].y;
            }
        }

        return [xMin, yMin, xMax, yMax];
    }

    window["AscPDF"].fillShapeByPoints  = fillShapeByPoints;
    window["AscPDF"].initShape  = initShape;
    window["AscPDF"].CAnnotationInk     = CAnnotationInk;
})();

