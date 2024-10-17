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
	 * Class representing a stamp annotation.
	 * @constructor
    */
    function CAnnotationStamp(sName, nPage, aRect, oDoc)
    {
        AscPDF.CPdfShape.call(this);
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Stamp, nPage, aRect, oDoc);
        
        this.Init();
    }
    
	CAnnotationStamp.prototype.constructor = CAnnotationStamp;
    AscFormat.InitClass(CAnnotationStamp, AscPDF.CPdfShape, AscDFH.historyitem_type_Pdf_Annot_Stamp);
    Object.assign(CAnnotationStamp.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationStamp.prototype.IsStamp = function() {
        return true;
    };
    CAnnotationStamp.prototype.SetInRect = function(aInRect) {
        this.inRect = aInRect;
        
        function getDistance(x1, y1, x2, y2) {
            return Math.sqrt(Math.pow(x2 - x1, 2) + Math.pow(y2 - y1, 2));
        }

        let nSourceW = getDistance(aInRect[0], aInRect[1], aInRect[6], aInRect[7]);
        let nSourceH = getDistance(aInRect[0], aInRect[1], aInRect[2], aInRect[3]);

        this.spPr.xfrm.setExtX(nSourceW * g_dKoef_pt_to_mm);
        this.spPr.xfrm.setExtY(nSourceH * g_dKoef_pt_to_mm);

        let aRect = this.GetRect();
        
        let nAnnotW = aRect[2] - aRect[0];
        let nAnnotH = aRect[3] - aRect[1];
        
        let nOffX = (aRect[0] - (nSourceW - nAnnotW) / 2);
        let nOffY = (aRect[1] - (nSourceH - nAnnotH) / 2);
        
        this.spPr.xfrm.setOffX(nOffX * g_dKoef_pt_to_mm);
        this.spPr.xfrm.setOffY(nOffY * g_dKoef_pt_to_mm);
    };
    CAnnotationStamp.prototype.GetInRect = function() {
        return this.inRect;
    };
    CAnnotationStamp.prototype.GetDrawing = function() {
        return this.content.GetAllDrawingObjects()[0];
    };
    CAnnotationStamp.prototype.SetWasChanged = function(isChanged) {
        let oViewer = Asc.editor.getDocumentRenderer();

        if (this._wasChanged !== isChanged && oViewer.IsOpenAnnotsInProgress == false) {
            this._wasChanged = isChanged;
        }
    };
    CAnnotationStamp.prototype.GetOriginViewScale = function() {
        let aInRect = this.GetInRect();
        if (!aInRect) {
            return 1;
        }

        function getDistance(x1, y1, x2, y2) {
            return Math.sqrt(Math.pow(x2 - x1, 2) + Math.pow(y2 - y1, 2));
        }

        let nSourceH = getDistance(aInRect[0], aInRect[1], aInRect[2], aInRect[3]) * g_dKoef_pt_to_mm;
        let nCurrentH = this.extY;

        return nCurrentH / nSourceH;
    };
    CAnnotationStamp.prototype.DrawFromStream = function(oGraphicsPDF) {
        if (this.IsHidden() == true)
            return;
            
        let nScale      = this.GetOriginViewScale();
        let originView  = this.GetOriginView(oGraphicsPDF.GetDrawingPageW() * nScale, oGraphicsPDF.GetDrawingPageH() * nScale);
        let nRot        = this.GetRot();

        if (originView) {
            let oXfrm = this.getXfrm();
            
            let X = oXfrm.offX / g_dKoef_pt_to_mm >> 0;
            let Y = oXfrm.offY / g_dKoef_pt_to_mm >> 0;

            if (this.IsHighlight())
                AscPDF.startMultiplyMode(oGraphicsPDF.GetContext());
            
            oGraphicsPDF.DrawImageXY(originView, X, Y, nRot);
            AscPDF.endMultiplyMode(oGraphicsPDF.GetContext());
        }

        oGraphicsPDF.SetLineWidth(1);
        let aOringRect  = this.GetRect();
        let X       = aOringRect[0];
        let Y       = aOringRect[1];
        let nWidth  = aOringRect[2] - aOringRect[0];
        let nHeight = aOringRect[3] - aOringRect[1];

        Y += 1 / 2;
        X += 1 / 2;
        nWidth  -= 1;
        nHeight -= 1;

        oGraphicsPDF.SetStrokeStyle(0, 255, 255);
        oGraphicsPDF.SetLineDash([]);
        oGraphicsPDF.BeginPath();
        oGraphicsPDF.Rect(X, Y, nWidth, nHeight);
        oGraphicsPDF.Stroke();
    };
    CAnnotationStamp.prototype.GetOriginViewInfo = function(nPageW, nPageH) {
        let oViewer     = editor.getDocumentRenderer();
        let oFile       = oViewer.file;
        let nPage       = this.GetOriginPage();

        if (this.APInfo == null || this.APInfo.size.w != nPageW || this.APInfo.size.h != nPageH) {
            this.APInfo = {
                info: oFile.nativeFile["getAnnotationsAP"](nPage, nPageW, nPageH, undefined, this.GetApIdx()),
                size: {
                    w: nPageW,
                    h: nPageH
                }
            }
        }
        
        for (let i = 0; i < this.APInfo.info.length; i++) {
            if (this.APInfo.info[i]["i"] == this._apIdx)
                return this.APInfo.info[i];
        }

        return null;
    };
    CAnnotationStamp.prototype.SetRect = function(aRect) {
        let oViewer     = editor.getDocumentRenderer();
        let oDoc        = oViewer.getPDFDoc();
        let aCurRect    = this.GetRect();

        let bCalcRect = this._origRect.length != 0 && false == AscCommon.History.UndoRedoInProgress;

        this._origRect = aRect;

        if (bCalcRect) {
            AscCommon.History.StartNoHistoryMode();

            this.Recalculate(true);
            this.recalcBounds();
            this.recalcGeometry();
            
            AscCommon.History.EndNoHistoryMode();
            
            let oGrBounds = this.bounds;

            let nX1 = aRect[0] * g_dKoef_pt_to_mm;
            let nX2 = aRect[2] * g_dKoef_pt_to_mm;
            let nY1 = aRect[1] * g_dKoef_pt_to_mm;
            let nY2 = aRect[3] * g_dKoef_pt_to_mm;

            this._origRect[0] = Math.round(oGrBounds.l) * g_dKoef_mm_to_pt;
            this._origRect[1] = Math.round(oGrBounds.t) * g_dKoef_mm_to_pt;
            this._origRect[2] = Math.round(oGrBounds.r) * g_dKoef_mm_to_pt;
            this._origRect[3] = Math.round(oGrBounds.b) * g_dKoef_mm_to_pt;

            this.spPr.xfrm.setExtX(nX2 - nX1);
            this.spPr.xfrm.setExtY(nY2 - nY1);
            this.spPr.xfrm.setOffX(nX1);
            this.spPr.xfrm.setOffY(nY1);
            
            oDoc.History.Add(new CChangesPDFAnnotRect(this, aCurRect, aRect));
        }

        this.SetWasChanged(true);
    };
    // CAnnotationStamp.prototype.SetRect = function(aRect) {
    //     let oViewer     = editor.getDocumentRenderer();
    //     let oDoc        = oViewer.getPDFDoc();

    //     oDoc.History.Add(new CChangesPDFAnnotRect(this, this.GetRect(), aRect));

    //     this._origRect = aRect;

    //     let oXfrm = this.getXfrm();
    //     if (oXfrm) {
    //         AscCommon.History.StartNoHistoryMode();

    //         let nX1 = aRect[0] * g_dKoef_pt_to_mm;
    //         let nX2 = aRect[2] * g_dKoef_pt_to_mm;
    //         let nY1 = aRect[1] * g_dKoef_pt_to_mm;
    //         let nY2 = aRect[3] * g_dKoef_pt_to_mm;

    //         this.spPr.xfrm.setExtX(nX2 - nX1);
    //         this.spPr.xfrm.setExtY(nY2 - nY1);
    //         this.spPr.xfrm.setOffX(nX1);
    //         this.spPr.xfrm.setOffY(nY1);
            
    //         this.SetNeedRecalc(true);
    //         this.RefillGeometry(this.spPr.geometry, [nX1, nY1, nX2, nY2]);

    //         AscCommon.History.EndNoHistoryMode();
    //     }
        
    //     this.SetWasChanged(true);
    // };
    CAnnotationStamp.prototype.canRotate = function() {
        return true;
    };
    CAnnotationStamp.prototype.Recalculate = function(bForce) {
        if (true !== bForce && false == this.IsNeedRecalc()) {
            return;
        }

        this.recalculateTransform();
        this.updateTransformMatrix();
        this.recalculate();
        this.SetNeedRecalc(false);
    };
    CAnnotationStamp.prototype.InitGeometry = function() {
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
    CAnnotationStamp.prototype.RefillGeometry = function(oGeometry, aBounds) {
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
    CAnnotationStamp.prototype.LazyCopy = function() {
        let oDoc = this.GetDocument();
        oDoc.StartNoHistoryMode();

        let oNewStamp = new CAnnotationStamp(AscCommon.CreateGUID(), this.GetPage(), this.GetOrigRect().slice(), oDoc);

        oNewStamp.sourceRect = this.sourceRect;
        oNewStamp.lazyCopy = true;

        this.fillObject(oNewStamp);

        let aStrokeColor    = this.GetStrokeColor();
        let aFillColor      = this.GetFillColor();

        oNewStamp._apIdx = this._apIdx;
        oNewStamp._originView = this._originView;
        oNewStamp.SetOriginPage(this.GetOriginPage());
        oNewStamp.SetAuthor(this.GetAuthor());
        oNewStamp.SetModDate(this.GetModDate());
        oNewStamp.SetCreationDate(this.GetCreationDate());
        aStrokeColor && oNewStamp.SetStrokeColor(aStrokeColor.slice());
        aFillColor && oNewStamp.SetFillColor(aFillColor.slice());
        oNewStamp.SetWidth(this.GetWidth());
        oNewStamp.SetOpacity(this.GetOpacity());
        oNewStamp.recalcGeometry()
        oNewStamp.Recalculate(true);

        oDoc.EndNoHistoryMode();
        return oNewStamp;
    };
    
    CAnnotationStamp.prototype.IsSelected = function() {
        let oViewer         = editor.getDocumentRenderer();
        let oDrawingObjects = oViewer.DrawingObjects;
        return oDrawingObjects.selectedObjects.includes(this);
    };
    
    CAnnotationStamp.prototype.SetIconType = function(sType) {
        if (sType == this._stampType) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFAnnotName(this, this._stampType, sType));
        this._stampType = sType;
        this.SetWasChanged(true);
    };
    CAnnotationStamp.prototype.GetIconType = function() {
        return this._stampType;
    };
    CAnnotationStamp.prototype.SetRotate = function(nAngle) {
        AscCommon.History.Add(new CChangesPDFAnnotRotate(this, this._rotate, nAngle));
		AscCommon.ExecuteNoHistory(function() {
			let oXfrm = this.getXfrm();
            oXfrm.setRot(-nAngle * (Math.PI / 180));
		}, undefined, this);
		
        this._rotate = nAngle;
        this.SetNeedRecalc(true);
    };
    CAnnotationStamp.prototype.GetRotate = function() {
        return this.GetRot() * (180 / Math.PI);
    };
    CAnnotationStamp.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        memory.WriteString(this.GetIconType());

        let nAngle = this.GetRotate();
        memory.WriteLong(nAngle);
        
        let nEndPos = memory.GetCurPosition();
        memory.Seek(memory.posForFlags);
        memory.WriteLong(memory.annotFlags);
        
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };
    
    CAnnotationStamp.prototype.getNoChangeAspect = function() {
        return true;
    };
    CAnnotationStamp.prototype.Init = function() {
        AscCommon.History.StartNoHistoryMode();

        let aOrigRect = this.GetRect();
        let aAnnotRectMM = aOrigRect ? aOrigRect.map(function(measure) {
            return measure * g_dKoef_pt_to_mm;
        }) : [];

        let nOffX = aAnnotRectMM[0];
        let nOffY = aAnnotRectMM[1];
        let nExtX = aAnnotRectMM[2] - aAnnotRectMM[0];
        let nExtY = aAnnotRectMM[3] - aAnnotRectMM[1];

        this.setSpPr(new AscFormat.CSpPr());
        this.spPr.setLn(new AscFormat.CLn());
        this.spPr.ln.setFill(AscFormat.CreateNoFillUniFill());
        this.spPr.setFill(AscFormat.CreateSolidFillRGBA(255, 255, 255, 255));
        this.spPr.setParent(this);
        this.spPr.setXfrm(new AscFormat.CXfrm());
        this.spPr.xfrm.setParent(this.spPr);
        
        this.spPr.xfrm.setOffX(nOffX);
        this.spPr.xfrm.setOffY(nOffY);
        this.spPr.xfrm.setExtX(nExtX);
        this.spPr.xfrm.setExtY(nExtY);
        
        this.spPr.setGeometry(AscFormat.CreateGeometry("rect"));

        this.setStyle(AscFormat.CreateDefaultShapeStyle());
        this.setBDeleted(false);
        this.recalculate();
        
        AscCommon.History.EndNoHistoryMode();
    };
    window["AscPDF"].CAnnotationStamp = CAnnotationStamp;
})();

