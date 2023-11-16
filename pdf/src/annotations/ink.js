/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Ink, nPage, aRect, oDoc);
        AscFormat.CShape.call(this);
        initShape(this);

        this._point         = undefined;
        this._popupOpen     = false;
        this._popupRect     = undefined;
        this._richContents  = undefined;
        this._rotate        = undefined;
        this._state         = undefined;
        this._stateModel    = undefined;
        this._width         = 1;
    }
    
	CAnnotationInk.prototype.constructor = CAnnotationInk;
    AscFormat.InitClass(CAnnotationInk, AscFormat.CShape, AscDFH.historyitem_type_Shape);
    Object.assign(CAnnotationInk.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationInk.prototype.IsInk = function() {
        return true;
    };
    CAnnotationInk.prototype.GetDrawing = function() {
        return this.content.GetAllDrawingObjects()[0];
    };
    CAnnotationInk.prototype.GetShapeInkMargins = function(bInMM) {
        let aRect = this.GetRect();
        let nScale = bInMM ? g_dKoef_pix_to_mm : 1;

        return {
            left:   (this._shapeBounds[0] - aRect[0]) * nScale,
            top:    (this._shapeBounds[1] - aRect[1]) * nScale,
            right:  (aRect[2] - this._shapeBounds[2]) * nScale,
            bottom: (aRect[3] - this._shapeBounds[3]) * nScale
        }
    };

    CAnnotationInk.prototype.onMouseDown = function(e) {
        let oViewer         = editor.getDocumentRenderer();
        let oDrawingObjects = oViewer.DrawingObjects;
        let oDoc            = this.GetDocument();
        let oDrDoc          = oDoc.GetDrawingDocument();

        this.selectStartPage = this.GetPage();
        let oPos    = oDrDoc.ConvertCoordsFromCursor2(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
        let X       = oPos.X;
        let Y       = oPos.Y;

        let pageObject = oViewer.getPageByCoords3(AscCommon.global_mouseEvent.X - oViewer.x, AscCommon.global_mouseEvent.Y - oViewer.y);

        oDrawingObjects.OnMouseDown(e, X, Y, pageObject.index);

        if (this.IsSelected()) {
            oDrawingObjects.handleEventMode = HANDLE_EVENT_MODE_CURSOR;
        }
        else {
            oDrawingObjects.handleEventMode = HANDLE_EVENT_MODE_HANDLE;
        }

        oDrawingObjects.OnMouseDown(e, X, Y, pageObject.index);
    };
    CAnnotationInk.prototype.SetInkPoints = function(aSourcePaths) {
        let oViewer = editor.getDocumentRenderer();
        let oDoc    = oViewer.getPDFDoc();
        let oDrDoc  = oDoc.GetDrawingDocument();
        let nPage   = this.GetPage();

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        let aShapePaths = [];
        for (let nPath = 0; nPath < aSourcePaths.length; nPath++) {
            let aSourcePath = aSourcePaths[nPath];
            let aShapePath  = [];
            
            for (let i = 0; i < aSourcePath.length - 1; i += 2) {
                aShapePath.push({
                    x: aSourcePath[i] * g_dKoef_pix_to_mm * nScaleX,
                    y: (aSourcePath[i + 1])* g_dKoef_pix_to_mm * nScaleY
                });
            }

            aShapePaths.push(aShapePath);
        }
        
        let aShapeRectInMM = this.GetRect().map(function(measure) {
            return measure * g_dKoef_pix_to_mm;
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
        this._gestures = aShapePaths;
    };
    CAnnotationInk.prototype.FillShapeByPoints = function(aPoints, oPen) {
        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = oViewer.getPDFDoc();
        let oDrDoc          = oDoc.GetDrawingDocument();

        let aShapeRectInMM = this.GetRect().map(function(measure) {
            return measure * g_dKoef_pix_to_mm;
        });
        let shape = fillShapeByPoints([aPoints], aShapeRectInMM, this);

        let oRGBPen = oPen.Fill.getRGBAColor();
        shape.spPr.setLn(oPen);
        shape.spPr.setFill(AscFormat.CreateNoFillUniFill());
        shape.recalculate();
        this.SetStrokeColor([oRGBPen.R / 255, oRGBPen.G / 255, oRGBPen.B / 255]);

        let aRelPointsPos = [];
        let aMinRect = getMinRect(aPoints);
        let xMin = aMinRect[0];
        let yMin = aMinRect[1];
        let xMax = aMinRect[2];
        let yMax = aMinRect[3];

        // считаем относительное положение точек внутри фигуры
        for (let nPoint = 0; nPoint < aPoints.length; nPoint++) {
            let oPoint = aPoints[nPoint];

            let nIndX = oPoint.x - xMin;
            let nIndY = oPoint.y - yMin;

            let oRelPos = {
                relX: nIndX / (xMax - xMin),
                relY: nIndY / (yMax - yMin)
            };

            if (isNaN(oRelPos.relX))
                oRelPos.relX = 0;
            if (isNaN(oRelPos.relY)) {
                oRelPos.relY = 0;
            }

            aRelPointsPos.push(oRelPos);
        }
        
        this._relativePaths = [aRelPointsPos];
        this._gestures = [aPoints];

        return shape;
    };

    CAnnotationInk.prototype.SetRect = function(aRect) {
        let oViewer     = editor.getDocumentRenderer();
        let oDoc        = oViewer.getPDFDoc();
        let nPage       = this.GetPage();

        oDoc.History.Add(new CChangesPDFAnnotRect(this, this.GetRect(), aRect));

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        this._rect = aRect;

        this._pagePos = {
            x: aRect[0],
            y: aRect[1],
            w: (aRect[2] - aRect[0]),
            h: (aRect[3] - aRect[1])
        };

        this._origRect[0] = this._rect[0] / nScaleX;
        this._origRect[1] = this._rect[1] / nScaleY;
        this._origRect[2] = this._rect[2] / nScaleX;
        this._origRect[3] = this._rect[3] / nScaleY;

        oDoc.TurnOffHistory();
        this.spPr.xfrm.extX = this._pagePos.w * g_dKoef_pix_to_mm;
        this.spPr.xfrm.extY = this._pagePos.h * g_dKoef_pix_to_mm;
        
        this.Recalculate();
        this.AddToRedraw();
        this.RefillGeometry(this.spPr.geometry, [aRect[0] * g_dKoef_pix_to_mm, aRect[1] * g_dKoef_pix_to_mm, aRect[2] * g_dKoef_pix_to_mm, aRect[3] * g_dKoef_pix_to_mm]);

        this.SetWasChanged(true);
        this.SetDrawFromStream(false);
    };
    CAnnotationInk.prototype.SetFlipV = function(bFlip) {
        let oDoc = this.GetDocument();
        
        if (this.flipV != bFlip) {
            oDoc.History.Add(new CChangesPDFInkFlipV(this, this.flipV, bFlip));
            this.changeFlipV(!this.flipV);
            this.recalculate();
        }
    };
    CAnnotationInk.prototype.SetFlipH = function(bFlip) {
        let oDoc = this.GetDocument();

        if (this.flipH != bFlip) {
            oDoc.History.Add(new CChangesPDFInkFlipH(this, this.flipV, bFlip));
            this.changeFlipH(!this.flipH);
            this.recalculate();
        }
    };
    CAnnotationInk.prototype.AddPath = function(aNewPath) {
        let oDoc = this.GetDocument();
        if (oDoc.History.UndoRedoInProgress == false) {
            oDoc.CreateNewHistoryPoint();
            oDoc.History.Add(new CChangesPDFInkPoints(this, this._gestures.length, aNewPath));
            oDoc.TurnOffHistory();
        }

        let nLineW = this.GetWidth() * g_dKoef_pt_to_mm * g_dKoef_mm_to_pix;

        let aCurAllPoints = [];
        for (let i = 0; i < this._gestures.length; i++)
            aCurAllPoints = aCurAllPoints.concat(this._gestures[i]);
    
        let aNewAllPoints = aCurAllPoints.slice();
        aNewAllPoints = aNewAllPoints.concat(aNewPath);
        
        let aCurMinRect = getMinRect(aCurAllPoints);
        let curMinX = aCurMinRect[0];
        let curMinY = aCurMinRect[1];
        let curMaxX = aCurMinRect[2];
        let curMaxY = aCurMinRect[3];

        let aNewMinRect = getMinRect(aNewAllPoints);
        let newMinX = aNewMinRect[0];
        let newMinY = aNewMinRect[1];
        let newMaxX = aNewMinRect[2];
        let newMaxY = aNewMinRect[3]; 
        
        let aNewAnnotRect = [(newMinX * g_dKoef_mm_to_pix - nLineW), (newMinY * g_dKoef_mm_to_pix - nLineW), (newMaxX * g_dKoef_mm_to_pix + nLineW), (newMaxY * g_dKoef_mm_to_pix + nLineW)];

        // пересчет текущих точек в новом ректе
        let aCurRelPaths = this.GetRelativePaths();

        for (let nPath = 0; nPath < aCurRelPaths.length; nPath++) {
            let aPath = aCurRelPaths[nPath];
            for (let nPoint = 0; nPoint < aPath.length; nPoint++) {
                let oPoint = aPath[nPoint];

                // если новый рект начинается левее старого
                if (newMinX < curMinX)
                    oPoint.relX = ((curMaxX - curMinX) * oPoint.relX + (curMinX - newMinX)) / (newMaxX - newMinX);
                // если ширина нового ректа больше старого
                else if (curMaxX - curMinX < newMaxX - newMinX)
                    oPoint.relX = ((curMaxX - curMinX) * oPoint.relX) / (newMaxX - newMinX);


                // если новый рект начинается выше старого
                if (newMinY < curMinY)
                    oPoint.relY = ((curMaxY - curMinY) * oPoint.relY + (curMinY - newMinY)) / (newMaxY - newMinY);
                // если высота нового ректа больше старого
                else if (curMaxY - curMinY < newMaxY - newMinY)
                    oPoint.relY = ((curMaxY - curMinY) * oPoint.relY) / (newMaxY - newMinY);

                if (isNaN(oPoint.relX))
                    oPoint.relX = 0;
                if (isNaN(oPoint.relY))
                    oPoint.relY = 0;

                continue;
            }
        }

        // считаем новые точки
        let aRelPointsPos = [];
        for (let nPoint = 0; nPoint < aNewPath.length; nPoint++) {
            let oPoint = aNewPath[nPoint];

            let nIndX = oPoint.x - newMinX;
            let nIndY = oPoint.y - newMinY;

            aRelPointsPos.push({
                relX: nIndX / (newMaxX - newMinX),
                relY: nIndY / (newMaxY - newMinY)
            });

            continue;
        }

        aCurRelPaths.push(aRelPointsPos);
        this._gestures.push(aNewPath);
        
        this.SetRect(aNewAnnotRect, true);
    };
    
    CAnnotationInk.prototype.RemoveLastAddedPath = function() {
        let nLineW = this.GetWidth() * g_dKoef_pt_to_mm * g_dKoef_mm_to_pix;

        let aCurAllPoints = [];
        for (let i = 0; i < this._gestures.length; i++)
            aCurAllPoints = aCurAllPoints.concat(this._gestures[i]);

        let aNewAllPoints = [];
        for (let i = 0; i < this._gestures.length - 1; i++)
            aNewAllPoints = aNewAllPoints.concat(this._gestures[i]);

        let aCurMinRect = getMinRect(aCurAllPoints);
        let curMinX = aCurMinRect[0];
        let curMinY = aCurMinRect[1];
        let curMaxX = aCurMinRect[2];
        let curMaxY = aCurMinRect[3];

        let aNewMinRect = getMinRect(aNewAllPoints);
        let newMinX = aNewMinRect[0];
        let newMinY = aNewMinRect[1];
        let newMaxX = aNewMinRect[2];
        let newMaxY = aNewMinRect[3];
        
        let aNewAnnotRect = [(newMinX * g_dKoef_mm_to_pix - nLineW), (newMinY * g_dKoef_mm_to_pix - nLineW), (newMaxX * g_dKoef_mm_to_pix + nLineW), (newMaxY * g_dKoef_mm_to_pix + nLineW)];

        // пересчет текущих точек в новом ректе
        let aCurRelPaths = this.GetRelativePaths();
        aCurRelPaths.splice(aCurRelPaths.length - 1, 1);
        this._gestures.splice(this._gestures.length - 1, 1);

        for (let nPath = 0; nPath < aCurRelPaths.length; nPath++) {
            let aPath = aCurRelPaths[nPath];
            for (let nPoint = 0; nPoint < aPath.length; nPoint++) {
                let oPoint = aPath[nPoint];

                // если предыдущий рект начинается левее старого
                if (newMinX > curMinX)
                    oPoint.relX = ((curMaxX - curMinX) * oPoint.relX + (curMinX - newMinX)) / (newMaxX - newMinX);
                // если ширина нового ректа больше старого
                else if (curMaxX - curMinX > newMaxX - newMinX)
                    oPoint.relX = ((curMaxX - curMinX) * oPoint.relX) / (newMaxX - newMinX);


                // если новый рект начинается выше старого
                if (newMinY > curMinY)
                    oPoint.relY = ((curMaxY - curMinY) * oPoint.relY + (curMinY - newMinY)) / (newMaxY - newMinY);
                // если высота нового ректа больше старого
                else if (curMaxY - curMinY > newMaxY - newMinY)
                    oPoint.relY = ((curMaxY - curMinY) * oPoint.relY) / (newMaxY - newMinY);

                if (isNaN(oPoint.relX))
                    oPoint.relX = 0;
                if (isNaN(oPoint.relY))
                    oPoint.relY = 0;
                        
                continue;
            }
        }

        this.spPr.xfrm.extX = (newMaxX + nLineW * g_dKoef_pix_to_mm) - (newMinX - nLineW * g_dKoef_pix_to_mm);
        this.spPr.xfrm.extY = (newMaxY + nLineW * g_dKoef_pix_to_mm) - (newMinY - nLineW * g_dKoef_pix_to_mm);

        this.SetRect(aNewAnnotRect, true);
    };
    CAnnotationInk.prototype.RefillGeometry = function(oGeometry, aBounds) {
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
        oDoc.TurnOffHistory();

        let oNewInk = new CAnnotationInk(AscCommon.CreateGUID(), this.GetPage(), this.GetRect().slice(), oDoc);

        oNewInk._pagePos = {
            x: this._pagePos.x,
            y: this._pagePos.y,
            w: this._pagePos.w,
            h: this._pagePos.h
        }
        oNewInk._origRect = this._origRect.slice();

        this.fillObject(oNewInk);

        oNewInk.pen = new AscFormat.CLn();
        oNewInk._apIdx = this._apIdx;
        oNewInk._originView = this._originView;
        oNewInk.SetOriginPage(this.GetOriginPage());
        oNewInk.SetAuthor(this.GetAuthor());
        oNewInk.SetModDate(this.GetModDate());
        oNewInk.SetCreationDate(this.GetCreationDate());
        oNewInk.SetWidth(this.GetWidth());
        oNewInk.SetStrokeColor(this.GetStrokeColor().slice());
        oNewInk._relativePaths = this.GetRelativePaths().slice();
        oNewInk._gestures = this._gestures.slice();
        oNewInk.SetContents(this.GetContents());
        oNewInk.recalcInfo.recalculatePen = false;
        oNewInk.recalcInfo.recalculateGeometry = true;

        return oNewInk;
    };
    CAnnotationInk.prototype.GetRelativePaths = function() {
        return this._relativePaths;
    };
    CAnnotationInk.prototype.SetStrokeColor = function(aColor) {
        this._strokeColor = aColor;

        let oRGB    = this.GetRGBColor(aColor);
        let oFill   = AscFormat.CreateSolidFillRGBA(oRGB.r, oRGB.g, oRGB.b, 255);
        let oLine   = this.pen;
        oLine.setFill(oFill);
    };
    CAnnotationInk.prototype.SetOpacity = function(value) {
        this._opacity = value;
        this.SetWasChanged(true);

        this.pen.Fill.transparent = value * 100 * 2.55;
    };
    CAnnotationInk.prototype.GetStrokeColor = function() {
        return this._strokeColor;
    };
    CAnnotationInk.prototype.SetWidth = function(nWidthPt) {
        this._width = nWidthPt; 

        nWidthPt = nWidthPt > 0 ? nWidthPt : 0.5;
        let oLine = this.pen;
        oLine.setW(nWidthPt * g_dKoef_pt_to_mm * 36000.0);
    };
    CAnnotationInk.prototype.GetWidth = function() {
        return this._width;
    };
    
    CAnnotationInk.prototype.IsSelected = function() {
        let oViewer         = editor.getDocumentRenderer();
        let oDrawingObjects = oViewer.DrawingObjects;
        return oDrawingObjects.selectedObjects.includes(this);
    };
    CAnnotationInk.prototype.SetDrawing = function(oDrawing) {
        let oRun = this.content.GetElement(0).GetElement(0);
        oRun.Add_ToContent(oRun.Content.length, oDrawing);
    };
    
    CAnnotationInk.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        // считаем точки
        let aRelPointsPos   = this._relativePaths;
        let aShapePaths     = [];
        let nLineW          = this.GetWidth() * g_dKoef_pt_to_mm * g_dKoef_mm_to_pix;
        let aBounds         = this.GetOrigRect();

        let xMin;
        let yMin;
        let xMax;
        let yMax;
        if (this.IsNeedDrawFromStream() == false) {
            xMin = aBounds[0] + nLineW * 0.75;
            yMin = aBounds[1] + nLineW * 0.75;
            xMax = aBounds[2] - nLineW * 0.75;
            yMax = aBounds[3] - nLineW * 0.75;

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
        }
        else {
            let oViewer = editor.getDocumentRenderer();
            let nScaleY = oViewer.drawingPages[this.GetPage()].H / oViewer.file.pages[this.GetPage()].H / oViewer.zoom;
            let nScaleX = oViewer.drawingPages[this.GetPage()].W / oViewer.file.pages[this.GetPage()].W / oViewer.zoom;

            for (let nPath = 0; nPath < this._gestures.length; nPath++) {
                let aPath       = this._gestures[nPath];
                let aShapePath  = [];

                for (let nPoint = 0; nPoint < aPath.length; nPoint++) {
                    aShapePath.push({
                        x: aPath[nPoint].x * g_dKoef_mm_to_pix / nScaleX,
                        y: aPath[nPoint].y * g_dKoef_mm_to_pix / nScaleY
                    });
                }
                
                aShapePaths.push(aShapePath);
            }
        }

        memory.WriteLong(aShapePaths.length);
        for (let i = 0; i < aShapePaths.length; i++) {
            memory.WriteLong(aShapePaths[i].length * 2);

            for (let j = 0; j < aShapePaths[i].length; j++) {
                memory.WriteDouble(aShapePaths[i][j].x);
                memory.WriteDouble(aShapePaths[i][j].y);
            }
        }

        let nEndPos = memory.GetCurPosition();
        memory.Seek(memory.posForFlags);
        memory.WriteLong(memory.annotFlags);
        
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);

        this._replies.forEach(function(reply) {
            reply.WriteToBinary(memory); 
        });
    };

    function fillShapeByPoints(arrOfArrPoints, aShapeRect, oParentAnnot) {
        let xMax = aShapeRect[2];
        let xMin = aShapeRect[0];
        let yMin = aShapeRect[1];
        let yMax = aShapeRect[3];

        let geometry = generateGeometry(arrOfArrPoints, [xMin, yMin, xMax, yMax]);
        
        oParentAnnot.spPr.setGeometry(geometry);
        oParentAnnot.updatePosition(xMin, yMin);

        oParentAnnot.x = xMin;
        oParentAnnot.y = yMin;
        return oParentAnnot;
    }

    function initShape(oParentAnnot) {
        let aShapeRectInMM = oParentAnnot.GetRect().map(function(measure) {
            return measure * g_dKoef_pix_to_mm;
        });
        let xMax = aShapeRectInMM[2];
        let xMin = aShapeRectInMM[0];
        let yMin = aShapeRectInMM[1];
        let yMax = aShapeRectInMM[3];

        oParentAnnot.setSpPr(new AscFormat.CSpPr());
        oParentAnnot.spPr.setParent(oParentAnnot);
        oParentAnnot.spPr.setXfrm(new AscFormat.CXfrm());
        oParentAnnot.spPr.xfrm.setParent(oParentAnnot.spPr);
        oParentAnnot.setWordShape(true);
        
        oParentAnnot.spPr.xfrm.setOffX(0);
        oParentAnnot.spPr.xfrm.setOffY(0);
        oParentAnnot.spPr.xfrm.setExtX(Math.abs(xMax - xMin));
        oParentAnnot.spPr.xfrm.setExtY(Math.abs(yMax - yMin));
        oParentAnnot.setStyle(AscFormat.CreateDefaultShapeStyle());
        oParentAnnot.setBDeleted(false);
        oParentAnnot.recalcInfo.recalculateGeometry = false;
        oParentAnnot.recalculate();
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

    function TurnOffHistory() {
        if (AscCommon.History.IsOn() == true)
            AscCommon.History.TurnOff();
    }

    window["AscPDF"].fillShapeByPoints  = fillShapeByPoints;
    window["AscPDF"].initShape  = initShape;
    window["AscPDF"].CAnnotationInk     = CAnnotationInk;
})();

