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

    let LINE_CAP_STYLES = {
        Butt:       0,
        Round:      1,
        Projecting: 2
    }

    let LINE_INTENT_TYPE = {
        Dimension:  0,
        Arrow:      1
    }

    let LINE_END_TYPE = {
        Square:         0,
        Circle:         1,
        Diamond:        2,
        OpenArrow:      3,
        ClosedArrow:    4,
        None:           5,
        Butt:           6,
        ROpenArrow:     7,
        RClosedArrow:   8,
        Slash:          9
    }

    let CAPTION_POSITIONING = {
        Inline: 0,
        Top:    1
    }
    /**
	 * Class representing a Ink annotation.
	 * @constructor
    */
    function CAnnotationLine(sName, nPage, aRect, oDoc)
    {
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Line, nPage, aRect, oDoc);
        AscFormat.CShape.call(this);
        AscPDF.initShape(this);

        this._popupOpen     = false;
        this._popupRect     = undefined;
        this._richContents  = undefined;
        this._rotate        = undefined;
        this._state         = undefined;
        this._stateModel    = undefined;
        this._width         = undefined;
        this._points        = undefined;
        this._doCaption     = undefined;
        this._intent        = undefined;
        this._lineStart     = undefined;
        this._lineEnd       = undefined;
        this._leaderLength  = undefined; // LL
        this._leaderExtend  = undefined; // LLE
        this._leaderLineOffset  = undefined; // LLO
        this._captionPos        = CAPTION_POSITIONING.Inline; // CP
        this._captionOffset     = undefined;  // CO
        this._needRecalc        = false;

        // internal
        TurnOffHistory();
    }
	CAnnotationLine.prototype.constructor = CAnnotationLine;
    AscFormat.InitClass(CAnnotationLine, AscFormat.CShape, AscDFH.historyitem_type_Shape);
    Object.assign(CAnnotationLine.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationLine.prototype.SetCaptionOffset = function(array) {
        this._captionOffset = array;
    };
    CAnnotationLine.prototype.GetCaptionOffset = function() {
        return this._captionOffset;
    };

    CAnnotationLine.prototype.SetCaptionPos = function(nPosType) {
        this._captionPos = nPosType;
    };
    CAnnotationLine.prototype.GetCaptionPos = function() {
        return this._captionPos;
    };
    CAnnotationLine.prototype.RefillGeometry = function() {
        let oViewer = editor.getDocumentRenderer();
        let oDoc    = oViewer.getPDFDoc();
        
        let aPoints = this.GetLinePoints();
        let nScaleY = oViewer.drawingPages[this.GetPage()].H / oViewer.file.pages[this.GetPage()].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[this.GetPage()].W / oViewer.file.pages[this.GetPage()].W / oViewer.zoom;

        let aLinePoints = [];
        for (let i = 0; i < aPoints.length - 1; i += 2) {
            aLinePoints.push({
                x: aPoints[i] * g_dKoef_pix_to_mm * nScaleX,
                y: (aPoints[i + 1])* g_dKoef_pix_to_mm * nScaleY
            });
        }
        
        let aShapeRectInMM = this.GetRect().map(function(measure) {
            return measure * g_dKoef_pix_to_mm;
        });

        oDoc.TurnOffHistory();
        AscPDF.fillShapeByPoints([aLinePoints], aShapeRectInMM, this);
        
        this._points = aPoints;
    }
    CAnnotationLine.prototype.SetLeaderLineOffset = function(nValue) {
        this._leaderLineOffset = nValue;
    };
    CAnnotationLine.prototype.GetLeaderLineOffset = function() {
        return this._leaderLineOffset;
    };
    CAnnotationLine.prototype.SetLeaderLength = function(nValue) {
        this._leaderLength = nValue;
    };
    CAnnotationLine.prototype.GetLeaderLength = function() {
        return this._leaderLength;
    };
    CAnnotationLine.prototype.SetLeaderExtend = function(nValue) {
        this._leaderExtend = nValue;
    };
    CAnnotationLine.prototype.GetLeaderExtend = function() {
        return this._leaderExtend;
    };
    CAnnotationLine.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        let oViewer     = editor.getDocumentRenderer();
        let nPage       = this.GetPage();
        let aOrigRect   = this.GetOrigRect();

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;
        
        this.handleUpdatePosition();
        if (this.recalcInfo.recalculateGeometry)
            this.RefillGeometry();

        this.recalculate();
        this.updatePosition(aOrigRect[0] * g_dKoef_pix_to_mm * nScaleX, aOrigRect[1] * g_dKoef_pix_to_mm * nScaleY)
    };
    CAnnotationLine.prototype.SetNeedRecalc = function(bRecalc) {
        this._needRecalc = bRecalc;
    };
    CAnnotationLine.prototype.IsNeedRecalc = function() {
        return this._needRecalc;
    };
    CAnnotationLine.prototype.SetLinePoints = function(aPoints) {
        let oViewer = editor.getDocumentRenderer();
        let oDoc    = oViewer.getPDFDoc();
        
        this.recalcGeometry();
        oDoc.History.Add(new CChangesPDFLinePoints(this, this.GetLinePoints(), aPoints));

        this._points = aPoints;
    };
    CAnnotationLine.prototype.GetLinePoints = function() {
        return this._points;
    };
    CAnnotationLine.prototype.onMouseDown = function(e) {
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
        oDrawingObjects.startEditGeometry();
    }
    CAnnotationLine.prototype.LazyCopy = function() {
        let oDoc = this.GetDocument();
        oDoc.TurnOffHistory();

        let oLine = new CAnnotationLine(AscCommon.CreateGUID(), this.GetPage(), this.GetRect().slice(), oDoc);

        oLine._pagePos = {
            x: this._pagePos.x,
            y: this._pagePos.y,
            w: this._pagePos.w,
            h: this._pagePos.h
        }
        oLine._origRect = this._origRect.slice();

        this.fillObject(oLine);

        oLine.pen = new AscFormat.CLn();
        oLine._apIdx = this._apIdx;
        oLine._originView = this._originView;
        oLine.SetOriginPage(this.GetOriginPage());
        oLine.SetAuthor(this.GetAuthor());
        oLine.SetModDate(this.GetModDate());
        oLine.SetCreationDate(this.GetCreationDate());
        oLine.SetWidth(this.GetWidth());
        oLine.SetStrokeColor(this.GetStrokeColor().slice());
        oLine.SetLineStart(this.GetLineStart());
        oLine.SetLineEnd(this.GetLineEnd());
        oLine.SetContents(this.GetContents());
        oLine.SetFillColor(this.GetFillColor());
        oLine.recalcInfo.recalculatePen = false;
        oLine.recalcInfo.recalculateGeometry = true;
        oLine._points = this._points.slice();
        oLine.recalculate();

        return oLine;
    };
    CAnnotationLine.prototype.IsLine = function() {
        return true;
    };
    CAnnotationLine.prototype.GetMinShapeRect = function() {
        let oViewer     = editor.getDocumentRenderer();
        let nLineWidth  = this.GetWidth() * g_dKoef_pt_to_mm * g_dKoef_mm_to_pix;
        let aPoints     = this.GetLinePoints();
        let nPage       = this.GetPage();

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        let shapeAtStart    = getFigureSize(this.GetLineStart(), nLineWidth);
        let shapeAtEnd      = getFigureSize(this.GetLineEnd(), nLineWidth);

        function calculateBoundingRectangle(line, figure1, figure2) {
            const {x1, y1, x2, y2} = line;
        
            // Расчет угла поворота в радианах
            const angle = Math.atan2(y2 - y1, x2 - x1);
        
            function rotatePoint(cx, cy, angle, px, py) {
                var cos = Math.cos(angle),
                    sin = Math.sin(angle),
                    nx = (sin * (px - cx)) + (cos * (py - cy)) + cx,
                    ny = (sin * (py - cy)) - (cos * (px - cx)) + cy;
                return {x: nx, y: ny};
            }
            
            function getRectangleCorners(cx, cy, width, height, angle) {
                var halfWidth = width / 2;
                var halfHeight = height / 2;
            
                // Углы прямоугольника до поворота
                var corners = [
                    {x: cx - halfWidth, y: cy - halfHeight}, // верхний левый
                    {x: cx + halfWidth, y: cy - halfHeight}, // верхний правый
                    {x: cx + halfWidth, y: cy + halfHeight}, // нижний правый
                    {x: cx - halfWidth, y: cy + halfHeight}  // нижний левый
                ];
            
                // Поворачиваем каждую точку
                return corners.map(function(point) {
                    return rotatePoint(cx, cy, angle, point.x, point.y);
                });
            }
        
            const cornersFigure1 = getRectangleCorners(x1, y1, figure1.width, figure1.height, angle);
            const cornersFigure2 = getRectangleCorners(x2, y2, figure2.width, figure2.height, angle);

            // Находим минимальные и максимальные координаты
            let minX = Math.min(x1, x2);
            let maxX = Math.max(x1, x2);
            let minY = Math.min(y1, y2);
            let maxY = Math.max(y1, y2);

            [...cornersFigure1, ...cornersFigure2].forEach(point => {
                minX = Math.min(minX, point.x);
                maxX = Math.max(maxX, point.x);
                minY = Math.min(minY, point.y);
                maxY = Math.max(maxY, point.y);
            });
        
            // Возвращаем координаты прямоугольника
            return [minX * nScaleX, minY * nScaleY, maxX * nScaleX, maxY * nScaleY];
        }

        return calculateBoundingRectangle({x1: aPoints[0], y1: aPoints[1], x2: aPoints[2], y2: aPoints[3]}, shapeAtStart, shapeAtEnd);
    };
    CAnnotationLine.prototype.SetRect = function(aRect) {
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
        
        this.AddToRedraw();
        this.SetWasChanged(true);
        this.SetDrawFromStream(false);
    };
    CAnnotationLine.prototype.SetDrawing = function(oDrawing) {
        let oRun = this.content.GetElement(0).GetElement(0);
        oRun.Add_ToContent(oRun.Content.length, oDrawing);
    };
    CAnnotationLine.prototype.SetStrokeColor = function(aColor) {
        this._strokeColor = aColor;

        let oRGB    = this.GetRGBColor(aColor);
        let oFill   = AscFormat.CreateSolidFillRGBA(oRGB.r, oRGB.g, oRGB.b, 255);
        let oLine   = this.pen;
        oLine.setFill(oFill);
    };
    CAnnotationLine.prototype.SetOpacity = function(value) {
        this._opacity = value;
        this.SetWasChanged(true);

        this.pen.Fill.transparent = value * 100 * 2.55;
    };
    CAnnotationLine.prototype.GetDrawing = function() {
        return this.content.GetAllDrawingObjects()[0];
    };
    CAnnotationLine.prototype.SetWidth = function(nWidthPt) {
        this._width = nWidthPt; 

        nWidthPt = nWidthPt > 0 ? nWidthPt : 0.5;
        let oLine = this.pen;
        oLine.setW(nWidthPt * g_dKoef_pt_to_mm * 36000.0);
    };
    CAnnotationLine.prototype.GetLinePoints = function() {
        return this._points;
    };
    CAnnotationLine.prototype.SetDoCaption = function(nType) {
        this._doCaption = nType;
    };
    CAnnotationLine.prototype.GetDoCaption = function() {
        return this._doCaption;
    };
    CAnnotationLine.prototype.SetLineStart = function(nType) {
        this._lineStart = nType;

        this.SetWasChanged(true);
        let oLine = this.pen;
        oLine.setHeadEnd(new AscFormat.EndArrow());
        let nLineEndType;
        switch (nType) {
            case LINE_END_TYPE.None:
                nLineEndType = AscFormat.LineEndType.None;
                break;
            case LINE_END_TYPE.OpenArrow:
                nLineEndType = AscFormat.LineEndType.Arrow;
                break;
            case LINE_END_TYPE.Diamond:
                nLineEndType = AscFormat.LineEndType.Diamond;
                break;
            case LINE_END_TYPE.Circle:
                nLineEndType = AscFormat.LineEndType.Oval;
                break;
            case LINE_END_TYPE.ClosedArrow:
                nLineEndType = AscFormat.LineEndType.Triangle;
                break;
            case LINE_END_TYPE.ROpenArrow:
                nLineEndType = AscFormat.LineEndType.ReverseArrow;
                break;
            case LINE_END_TYPE.RClosedArrow:
                nLineEndType = AscFormat.LineEndType.ReverseTriangle;
                break;
            case LINE_END_TYPE.Butt:
                nLineEndType = AscFormat.LineEndType.Butt;
                break;
            case LINE_END_TYPE.Square:
                nLineEndType = AscFormat.LineEndType.Square;
                break;
            case LINE_END_TYPE.Slash:
                nLineEndType = AscFormat.LineEndType.Slash;
                break;
            default:
                nLineEndType = AscFormat.LineEndType.Arrow;
                break;
        }

        oLine.headEnd.setType(nLineEndType);
        oLine.headEnd.setLen(AscFormat.LineEndSize.Mid);
    };
    CAnnotationLine.prototype.SetLineEnd = function(nType) {
        this._lineEnd = nType;
        
        this.SetWasChanged(true);
        let oLine = this.pen;
        oLine.setTailEnd(new AscFormat.EndArrow());
        let nLineEndType;
        switch (nType) {
            case LINE_END_TYPE.None:
                nLineEndType = AscFormat.LineEndType.None;
                break;
            case LINE_END_TYPE.OpenArrow:
                nLineEndType = AscFormat.LineEndType.Arrow;
                break;
            case LINE_END_TYPE.Diamond:
                nLineEndType = AscFormat.LineEndType.Diamond;
                break;
            case LINE_END_TYPE.Circle:
                nLineEndType = AscFormat.LineEndType.Oval;
                break;
            case LINE_END_TYPE.ClosedArrow:
                nLineEndType = AscFormat.LineEndType.Triangle;
                break;
            case LINE_END_TYPE.ROpenArrow:
                nLineEndType = AscFormat.LineEndType.ReverseArrow;
                break;
            case LINE_END_TYPE.RClosedArrow:
                nLineEndType = AscFormat.LineEndType.ReverseTriangle;
                break;
            case LINE_END_TYPE.Butt:
                nLineEndType = AscFormat.LineEndType.Butt;
                break;
            case LINE_END_TYPE.Square:
                nLineEndType = AscFormat.LineEndType.Square;
                break;
            case LINE_END_TYPE.Slash:
                nLineEndType = AscFormat.LineEndType.Slash;
                break;
            default:
                nLineEndType = AscFormat.LineEndType.Arrow;
                break;
        }

        oLine.tailEnd.setType(nLineEndType);
        oLine.tailEnd.setLen(AscFormat.LineEndSize.Mid);
    };
    CAnnotationLine.prototype.GetLineStart = function() {
        return this._lineStart;
    };
    CAnnotationLine.prototype.GetLineEnd = function() {
        return this._lineEnd;
    };

    CAnnotationLine.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        // line points
        let aLinePoints = this.GetLinePoints();
        for (let i = 0; i < aLinePoints.length; i++) {
            memory.WriteDouble(aLinePoints[i]);
        }

        // line ending
        let nLS = this.GetLineStart();
        let nLE = this.GetLineEnd();
        if (nLE != null && nLS != null) {
            memory.annotFlags |= (1 << 15);
            memory.WriteByte(nLS);
            memory.WriteByte(nLE);
        }

        // fill
        let aFill = this.GetFillColor();
        if (aFill != null) {
            memory.annotFlags |= (1 << 16);
            memory.WriteLong(aFill.length);
            for (let i = 0; i < aFill.length; i++)
                memory.WriteDouble(aFill[i]);
        }

        // leader leader
        let nLL = this.GetLeaderLength();
        if (nLL) {
            memory.annotFlags |= (1 << 17);
            memory.WriteDouble(nLL);
        }

        // leader extend
        let nLLE = this.GetLeaderExtend();
        if (nLLE) {
            memory.annotFlags |= (1 << 18);
            memory.WriteDouble(nLLE);
        }

        // do caption
        let bDoCaption = this.GetDoCaption();
        if (bDoCaption) {
            memory.annotFlags |= (1 << 19);
        }
        
        // intent
        let nIntent = this.GetIntent();
        if (nIntent != null) {
            memory.annotFlags |= (1 << 20);
            memory.WriteDouble(nIntent);
        }

        // leader Line Offset
        let nLLO = this.GetLeaderLineOffset();
        if (nLLO != null) {
            memory.annotFlags |= (1 << 21);
            memory.WriteDouble(nLLO);
        }

        // caption positioning
        let nCP = this.GetCaptionPos();
        if (nCP != null) {
            memory.annotFlags |= (1 << 22);
            memory.WriteByte(nCP);
        }

        // caption offset
        let aCO = this.GetCaptionOffset();
        if (aCO != null) {
            memory.annotFlags |= (1 << 23);
            memory.WriteDouble(aCO[0]);
            memory.WriteDouble(aCO[1]);
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

    function generateGeometry(arrOfArrPoints, aBounds, oGeometry) {
        let xMin = aBounds[0];
        let yMin = aBounds[1];
        let xMax = aBounds[2];
        let yMax = aBounds[3];
        // let xMin = Math.min(arrOfArrPoints[0][0].x, arrOfArrPoints[0][1].x);
        // let yMin = Math.min(arrOfArrPoints[0][0].y, arrOfArrPoints[0][1].y);
        // let xMax = Math.max(arrOfArrPoints[0][0].x, arrOfArrPoints[0][1].x);
        // let yMax = Math.max(arrOfArrPoints[0][0].y, arrOfArrPoints[0][1].y);

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

    function getFigureSize(nType, nLineW) {
        let oSize = {width: 0, height: 0};

        switch (nType) {
            case LINE_END_TYPE.None:
                oSize.width = nLineW;
                oSize.height = nLineW;
            case LINE_END_TYPE.OpenArrow:
            case LINE_END_TYPE.ClosedArrow:
                oSize.width = 4 * nLineW;
                oSize.height = 2 * nLineW;
                break;
            case LINE_END_TYPE.Diamond:
            case LINE_END_TYPE.Square:
                oSize.width = 4 * nLineW;
                oSize.height = 4 * nLineW;
                break;
            case LINE_END_TYPE.Circle:
                oSize.width = 4 * nLineW;
                oSize.height = 4 * nLineW;
                break;
            case LINE_END_TYPE.RClosedArrow:
                oSize.width = 6 * nLineW;
                oSize.height = 6 * nLineW;
                break;
            case LINE_END_TYPE.ROpenArrow:
                oSize.width = 5 * nLineW;
                oSize.height = 5 * nLineW;
                break;
            case LINE_END_TYPE.Butt:
                oSize.width = 5 * nLineW;
                oSize.height = 1.5 * nLineW;
                break;
            case LINE_END_TYPE.Slash:
                oSize.width = 4 * nLineW;
                oSize.height = 3.5 * nLineW;
                break;
            
        }

        return oSize;
    }

    function TurnOffHistory() {
        if (AscCommon.History.IsOn() == true)
            AscCommon.History.TurnOff();
    }

    window["AscPDF"].CAnnotationLine = CAnnotationLine;
})();

