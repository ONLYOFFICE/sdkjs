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

    let POLYLINE_INTENT_TYPE = {
        PolygonCloud:       0,
        PolyLineDimension:  1,
        PolygonDimension:   2
    }

    /**
	 * Class representing a Ink annotation.
	 * @constructor
    */
    function CAnnotationPolyLine(sName, nPage, aRect, oDoc)
    {
        AscPDF.CPdfShape.call(this);
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.PolyLine, nPage, aRect, oDoc);
        
        AscPDF.initShape(this);

        this._point         = undefined;
        this._popupOpen     = false;
        this._popupRect     = undefined;
        this._richContents  = undefined;
        this._rotate        = undefined;
        this._state         = undefined;
        this._stateModel    = undefined;
        this._width         = undefined;
        this._lineStart     = undefined;
        this._lineEnd       = undefined;
        this._vertices      = undefined;
        this._width         = undefined;
    }
    CAnnotationPolyLine.prototype.constructor = CAnnotationPolyLine;
    AscFormat.InitClass(CAnnotationPolyLine, AscPDF.CPdfShape, AscDFH.historyitem_type_Pdf_Annot_Polyline);
    Object.assign(CAnnotationPolyLine.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationPolyLine.prototype.SetVertices = function(aVertices) {
        let oViewer = editor.getDocumentRenderer();
        let oDoc    = oViewer.getPDFDoc();
        
        this.recalcGeometry();
        oDoc.History.Add(new CChangesPDFAnnotVertices(this, this.GetVertices(), aVertices));

        this._vertices = aVertices;
    };
    CAnnotationPolyLine.prototype.GetVertices = function() {
        return this._vertices;
    };

    CAnnotationPolyLine.prototype.Recalculate = function(bForce) {
        if (true !== bForce && false == this.IsNeedRecalc()) {
            return;
        }

        if (this.recalcInfo.recalculateGeometry)
            this.RefillGeometry();

        this.recalculateTransform();
        this.updateTransformMatrix();
        this.recalculate();
        this.SetNeedRecalc(false);
    };
    CAnnotationPolyLine.prototype.RefillGeometry = function() {
        let oViewer = editor.getDocumentRenderer();
        let oDoc    = oViewer.getPDFDoc();
        
        let aPoints = this.GetVertices();
        let aPolygonPoints = [];

        for (let i = 0; i < aPoints.length - 1; i += 2) {
            aPolygonPoints.push({
                x: aPoints[i] * g_dKoef_pt_to_mm,
                y: (aPoints[i + 1])* g_dKoef_pt_to_mm
            });
        }
        
        let aShapeRectInMM = this.GetOrigRect().map(function(measure) {
            return measure * g_dKoef_pt_to_mm;
        });

        oDoc.StartNoHistoryMode();
        fillShapeByPoints([aPolygonPoints], aShapeRectInMM, this);
        oDoc.EndNoHistoryMode();
    };
    CAnnotationPolyLine.prototype.SetRect = function(aOrigRect) {
        let oViewer     = editor.getDocumentRenderer();
        let oDoc        = oViewer.getPDFDoc();

        oDoc.History.Add(new CChangesPDFAnnotRect(this, this.GetOrigRect(), aOrigRect));

        this._origRect = aOrigRect;

        let oXfrm = this.getXfrm();
        if (oXfrm) {
            AscCommon.History.StartNoHistoryMode();
            oXfrm.setOffX(aOrigRect[0] * g_dKoef_pt_to_mm);
            oXfrm.setOffY(aOrigRect[1] * g_dKoef_pt_to_mm);
            oXfrm.setExtX((aOrigRect[2] - aOrigRect[0]) * g_dKoef_pt_to_mm);
            oXfrm.setExtY((aOrigRect[3] - aOrigRect[1]) * g_dKoef_pt_to_mm);
            AscCommon.History.EndNoHistoryMode();
        }
        
        this.AddToRedraw();
        this.SetWasChanged(true);
    };
    CAnnotationPolyLine.prototype.LazyCopy = function() {
        let oDoc = this.GetDocument();
        oDoc.StartNoHistoryMode();

        let oPolyline = new CAnnotationPolyLine(AscCommon.CreateGUID(), this.GetPage(), this.GetOrigRect().slice(), oDoc);
        oPolyline.lazyCopy = true;

        this.fillObject(oPolyline);

        let aStrokeColor    = this.GetStrokeColor();
        let aFillColor      = this.GetFillColor();
        let aVertices       = this.GetVertices();

        oPolyline._apIdx = this._apIdx;
        oPolyline._originView = this._originView;
        oPolyline.SetOriginPage(this.GetOriginPage());
        oPolyline.SetAuthor(this.GetAuthor());
        oPolyline.SetModDate(this.GetModDate());
        oPolyline.SetCreationDate(this.GetCreationDate());
        oPolyline.SetContents(this.GetContents());
        aStrokeColor && oPolyline.SetStrokeColor(aStrokeColor.slice());
        aFillColor && oPolyline.SetFillColor(aFillColor.slice());
        oPolyline.SetWidth(this.GetWidth());
        oPolyline.SetLineStart(this.GetLineStart());
        oPolyline.SetLineEnd(this.GetLineEnd());
        oPolyline.SetOpacity(this.GetOpacity());
        aVertices && oPolyline.SetVertices(aVertices.slice());
        oPolyline.SetWasChanged(oPolyline.IsChanged());
        oPolyline.recalcInfo.recalculateGeometry = true;
        
        oDoc.EndNoHistoryMode();

        return oPolyline;
    };
    CAnnotationPolyLine.prototype.onMouseDown = function(x, y, e) {
        let oViewer         = Asc.editor.getDocumentRenderer();
        let oDrawingObjects = oViewer.DrawingObjects;

        this.selectStartPage = this.GetPage();

        let pageObject = oViewer.getPageByCoords2(x, y);
        if (!pageObject)
            return false;

        let X = pageObject.x;
        let Y = pageObject.y;

        oDrawingObjects.OnMouseDown(e, X, Y, pageObject.index);
        oDrawingObjects.startEditGeometry();
    };
    CAnnotationPolyLine.prototype.IsPolyLine = function() {
        return true;
    };
    CAnnotationPolyLine.prototype.SetLineStart = function(nType) {
        this._lineStart = nType;

        this.SetWasChanged(true);
        let oLine = this.spPr.ln;
        oLine.setHeadEnd(new AscFormat.EndArrow());
        let nLineEndType;
        switch (nType) {
            case AscPDF.LINE_END_TYPE.None:
                nLineEndType = AscFormat.LineEndType.None;
                break;
            case AscPDF.LINE_END_TYPE.OpenArrow:
                nLineEndType = AscFormat.LineEndType.Arrow;
                break;
            case AscPDF.LINE_END_TYPE.Diamond:
                nLineEndType = AscFormat.LineEndType.Diamond;
                break;
            case AscPDF.LINE_END_TYPE.Circle:
                nLineEndType = AscFormat.LineEndType.Oval;
                break;
            case AscPDF.LINE_END_TYPE.ClosedArrow:
                nLineEndType = AscFormat.LineEndType.Triangle;
                break;
            case AscPDF.LINE_END_TYPE.ROpenArrow:
                nLineEndType = AscFormat.LineEndType.ReverseArrow;
                break;
            case AscPDF.LINE_END_TYPE.RClosedArrow:
                nLineEndType = AscFormat.LineEndType.ReverseTriangle;
                break;
            case AscPDF.LINE_END_TYPE.Butt:
                nLineEndType = AscFormat.LineEndType.Butt;
                break;
            case AscPDF.LINE_END_TYPE.Square:
                nLineEndType = AscFormat.LineEndType.Square;
                break;
            case AscPDF.LINE_END_TYPE.Slash:
                nLineEndType = AscFormat.LineEndType.Slash;
                break;
            default:
                nLineEndType = AscFormat.LineEndType.Arrow;
                break;
        }

        oLine.headEnd.setType(nLineEndType);
        oLine.headEnd.setLen(AscFormat.LineEndSize.Mid);
        this.handleUpdateLn();
    };
    CAnnotationPolyLine.prototype.SetLineEnd = function(nType) {
        this._lineEnd = nType;
        
        this.SetWasChanged(true);
        let oLine = this.spPr.ln;
        oLine.setTailEnd(new AscFormat.EndArrow());
        let nLineEndType;
        switch (nType) {
            case AscPDF.LINE_END_TYPE.None:
                nLineEndType = AscFormat.LineEndType.None;
                break;
            case AscPDF.LINE_END_TYPE.OpenArrow:
                nLineEndType = AscFormat.LineEndType.Arrow;
                break;
            case AscPDF.LINE_END_TYPE.Diamond:
                nLineEndType = AscFormat.LineEndType.Diamond;
                break;
            case AscPDF.LINE_END_TYPE.Circle:
                nLineEndType = AscFormat.LineEndType.Oval;
                break;
            case AscPDF.LINE_END_TYPE.ClosedArrow:
                nLineEndType = AscFormat.LineEndType.Triangle;
                break;
            case AscPDF.LINE_END_TYPE.ROpenArrow:
                nLineEndType = AscFormat.LineEndType.ReverseArrow;
                break;
            case AscPDF.LINE_END_TYPE.RClosedArrow:
                nLineEndType = AscFormat.LineEndType.ReverseTriangle;
                break;
            case AscPDF.LINE_END_TYPE.Butt:
                nLineEndType = AscFormat.LineEndType.Butt;
                break;
            case AscPDF.LINE_END_TYPE.Square:
                nLineEndType = AscFormat.LineEndType.Square;
                break;
            case AscPDF.LINE_END_TYPE.Slash:
                nLineEndType = AscFormat.LineEndType.Slash;
                break;
            default:
                nLineEndType = AscFormat.LineEndType.Arrow;
                break;
        }

        oLine.tailEnd.setType(nLineEndType);
        oLine.tailEnd.setLen(AscFormat.LineEndSize.Mid);
        this.handleUpdateLn();
    };
    CAnnotationPolyLine.prototype.GetLineStart = function() {
        return this._lineStart;
    };
    CAnnotationPolyLine.prototype.GetLineEnd = function() {
        return this._lineEnd;
    };
    CAnnotationPolyLine.prototype.GetMinShapeRect = function() {
        let nLineWidth  = this.GetWidth();
        let aVertices   = this.GetVertices();

        let shapeSizeAtStart    = getFigureSize(this.GetLineStart(), nLineWidth);
        let shapeSizeAtEnd      = getFigureSize(this.GetLineEnd(), nLineWidth);

        function calculateBoundingRectangle(line, figure1, figure2) {
            const x1 = line.x1, y1 = line.y1, x2 = line.x2, y2 = line.y2;
        
            // Calculate the rotation angle in radians
            const angle = Math.atan2(y2 - y1, x2 - x1);
        
            function rotatePoint(cx, cy, angle, px, py) {
                let cos = Math.cos(angle),
                    sin = Math.sin(angle),
                    nx = (sin * (px - cx)) + (cos * (py - cy)) + cx,
                    ny = (sin * (py - cy)) - (cos * (px - cx)) + cy;
                return {x: nx, y: ny};
            }
        
            function getRectangleCorners(cx, cy, width, height, angle) {
                let halfWidth = width / 2;
                let halfHeight = height / 2;
        
                // Corners of the rectangle before rotation
                let corners = [
                    {x: cx - halfWidth, y: cy - halfHeight}, // top left
                    {x: cx + halfWidth, y: cy - halfHeight}, // top right
                    {x: cx + halfWidth, y: cy + halfHeight}, // bottom right
                    {x: cx - halfWidth, y: cy + halfHeight}  // bottom left
                ];
        
                // Rotate each point
                let rotatedCorners = [];
                for (let i = 0; i < corners.length; i++) {
                    rotatedCorners.push(rotatePoint(cx, cy, angle, corners[i].x, corners[i].y));
                }
                return rotatedCorners;
            }
        
            let cornersFigure1 = getRectangleCorners(x1, y1, figure1.width, figure1.height, angle);
            let cornersFigure2 = getRectangleCorners(x2, y2, figure2.width, figure2.height, angle);
        
            // Find minimum and maximum coordinates
            let minX = Math.min(x1, x2);
            let maxX = Math.max(x1, x2);
            let minY = Math.min(y1, y2);
            let maxY = Math.max(y1, y2);
        
            for (let i = 0; i < cornersFigure1.length; i++) {
                let point = cornersFigure1[i];
                minX = Math.min(minX, point.x);
                maxX = Math.max(maxX, point.x);
                minY = Math.min(minY, point.y);
                maxY = Math.max(maxY, point.y);
            }
        
            for (let i = 0; i < cornersFigure2.length; i++) {
                let point = cornersFigure2[i];
                minX = Math.min(minX, point.x);
                maxX = Math.max(maxX, point.x);
                minY = Math.min(minY, point.y);
                maxY = Math.max(maxY, point.y);
            }
        
            // Return the coordinates of the rectangle
            return [minX, minY, maxX, maxY];
        }

        let oStartLine = {
            x1: aVertices[0],
            y1: aVertices[1],
            x2: aVertices[2],
            y2: aVertices[3]
        }
        let oEndLine = {
            x1: aVertices[aVertices.length - 4],
            y1: aVertices[aVertices.length - 3],
            x2: aVertices[aVertices.length - 2],
            y2: aVertices[aVertices.length - 1]
        }

        function findBoundingRectangle(points) {
            let x_min = points[0], y_min = points[1];
            let x_max = points[0], y_max = points[1];
        
            for (let i = 2; i < points.length; i += 2) {
                x_min = Math.min(x_min, points[i]);
                x_max = Math.max(x_max, points[i]);
                y_min = Math.min(y_min, points[i + 1]);
                y_max = Math.max(y_max, points[i + 1]);
            }
        
            return [x_min, y_min, x_max, y_max];
        }

        // находим ректы исходных точек. Стартовой линии учитывая lineStart фигуру, и такую же для конца
        // далее нахоим объединения всех прямоугольников для получения результирующего
        let aSourceRect     = findBoundingRectangle(aVertices);
        let aStartLineRect  = calculateBoundingRectangle(oStartLine, shapeSizeAtStart, {width: 0, height: 0});
        let aEndLineRect    = calculateBoundingRectangle(oEndLine, {width: 0, height: 0} , shapeSizeAtEnd);

        return unionRectangles([aSourceRect, aStartLineRect, aEndLineRect]);
    };
    CAnnotationPolyLine.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        // vertices
        let aVertices = this.GetVertices();
        if (aVertices) {
            memory.WriteLong(aVertices.length);
            for (let i = 0; i < aVertices.length; i++) {
                memory.WriteDouble(aVertices[i]);
            }
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

        // intent
        let nIntent = this.GetIntent();
        if (nIntent != null) {
            memory.annotFlags |= (1 << 20);
            memory.WriteByte(nIntent);
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

        let geometry = generateGeometry(arrOfArrPoints, [xMin, yMin, xMax, yMax]);
        oParentAnnot.spPr.setGeometry(geometry);

        let oXfrm = oParentAnnot.getXfrm();
        oXfrm.offX = xMin;
        oXfrm.offY = yMin;

        return oParentAnnot;
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
            
            geometry.AddPathCommand(0,undefined, "none", undefined, pathW, pathH);
            geometry.AddPathCommand(1, (((aPoints[0].x - xMin) * kw) >> 0) + "", (((aPoints[0].y - yMin) * kh) >> 0) + "");

            let oPt, nPt;
            for(nPt = 1; nPt < aPoints.length; nPt++) {
                oPt = aPoints[nPt];

                geometry.AddPathCommand(2,
                    (((oPt.x - xMin) * kw) >> 0) + "", (((oPt.y - yMin) * kh) >> 0) + ""
                );
            }
        }

        geometry.preset = null;
        geometry.rectS = null;
        return geometry;
    }

    function getFigureSize(nType, nLineW) {
        let oSize = {width: 0, height: 0};

        switch (nType) {
            case AscPDF.LINE_END_TYPE.None:
                oSize.width = nLineW;
                oSize.height = nLineW;
            case AscPDF.LINE_END_TYPE.OpenArrow:
            case AscPDF.LINE_END_TYPE.ClosedArrow:
                oSize.width = 4 * nLineW;
                oSize.height = 2 * nLineW;
                break;
            case AscPDF.LINE_END_TYPE.Diamond:
            case AscPDF.LINE_END_TYPE.Square:
                oSize.width = 4 * nLineW;
                oSize.height = 4 * nLineW;
                break;
            case AscPDF.LINE_END_TYPE.Circle:
                oSize.width = 4 * nLineW;
                oSize.height = 4 * nLineW;
                break;
            case AscPDF.LINE_END_TYPE.RClosedArrow:
                oSize.width = 6 * nLineW;
                oSize.height = 6 * nLineW;
                break;
            case AscPDF.LINE_END_TYPE.ROpenArrow:
                oSize.width = 5 * nLineW;
                oSize.height = 5 * nLineW;
                break;
            case AscPDF.LINE_END_TYPE.Butt:
                oSize.width = 5 * nLineW;
                oSize.height = 1.5 * nLineW;
                break;
            case AscPDF.LINE_END_TYPE.Slash:
                oSize.width = 4 * nLineW;
                oSize.height = 3.5 * nLineW;
                break;
            
        }

        return oSize;
    }

    function unionRectangles(rects) {
        let minX = Infinity;
        let minY = Infinity;
        let maxX = -Infinity;
        let maxY = -Infinity;
    
        rects.forEach(function(rect) {
            if (!rect)
                return;

            minX = Math.min(minX, rect[0]);
            minY = Math.min(minY, rect[1]);
            maxX = Math.max(maxX, rect[2]);
            maxY = Math.max(maxY, rect[3]);
        });
    
        return [minX, minY, maxX, maxY];
    }

    window["AscPDF"].CAnnotationPolyLine = CAnnotationPolyLine;
    window["AscPDF"].unionRectangles = unionRectangles;
})();

