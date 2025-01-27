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
        AscPDF.CPdfShape.call(this);
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Line, nPage, aRect, oDoc);
        
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
    }
	CAnnotationLine.prototype.constructor = CAnnotationLine;
    AscFormat.InitClass(CAnnotationLine, AscPDF.CPdfShape, AscDFH.historyitem_type_Pdf_Annot_Line);
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
        let aLinePoints = [];
        for (let i = 0; i < aPoints.length - 1; i += 2) {
            aLinePoints.push({
                x: aPoints[i] * g_dKoef_pt_to_mm,
                y: (aPoints[i + 1])* g_dKoef_pt_to_mm
            });
        }
        
        let aShapeRectInMM = this.GetOrigRect().map(function(measure) {
            return measure * g_dKoef_pt_to_mm;
        });

        oDoc.StartNoHistoryMode();
        AscPDF.fillShapeByPoints([aLinePoints], aShapeRectInMM, this);
        oDoc.EndNoHistoryMode();
    };
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
    CAnnotationLine.prototype.Recalculate = function(bForce) {
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
    CAnnotationLine.prototype.onMouseDown = function(x, y, e) {
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
    CAnnotationLine.prototype.LazyCopy = function() {
        let oDoc = this.GetDocument();
        oDoc.StartNoHistoryMode();

        let oLine = new CAnnotationLine(AscCommon.CreateGUID(), this.GetPage(), this.GetOrigRect().slice(), oDoc);

        oLine.lazyCopy = true;

        this.fillObject(oLine);

        let aStrokeColor    = this.GetStrokeColor();
        let aFillColor      = this.GetFillColor();
        let aLinePoints     = this.GetLinePoints();

        oLine._apIdx = this._apIdx;
        oLine._originView = this._originView;
        oLine.SetOriginPage(this.GetOriginPage());
        oLine.SetAuthor(this.GetAuthor());
        oLine.SetModDate(this.GetModDate());
        oLine.SetCreationDate(this.GetCreationDate());
        aStrokeColor && oLine.SetStrokeColor(aStrokeColor.slice());
        oLine.SetWidth(this.GetWidth());
        oLine.SetLineStart(this.GetLineStart());
        oLine.SetLineEnd(this.GetLineEnd());
        oLine.SetContents(this.GetContents());
        aFillColor && oLine.SetFillColor(aFillColor.slice());
        oLine.SetOpacity(this.GetOpacity());
        aLinePoints && oLine.SetLinePoints(aLinePoints.slice());
        oLine.recalcInfo.recalculateGeometry = true;
        oLine.recalculate();

        oDoc.EndNoHistoryMode();
        return oLine;
    };
    CAnnotationLine.prototype.IsLine = function() {
        return true;
    };
    CAnnotationLine.prototype.GetMinShapeRect = function() {
        let nLineWidth  = this.GetWidth();
        let aPoints     = this.GetLinePoints();

        let shapeSizeAtStart    = getFigureSize(this.GetLineStart(), nLineWidth);
        let shapeSizeAtEnd      = getFigureSize(this.GetLineEnd(), nLineWidth);

        function calculateBoundingRectangle(line, figure1, figure2) {
            let x1 = line.x1, y1 = line.y1, x2 = line.x2, y2 = line.y2;
        
            // Расчет угла поворота в радианах
            let angle = Math.atan2(y2 - y1, x2 - x1);
        
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
            
                let topLeft = {x: cx - halfWidth, y: cy - halfHeight},
                    topRight = {x: cx + halfWidth, y: cy - halfHeight},
                    bottomRight = {x: cx + halfWidth, y: cy + halfHeight},
                    bottomLeft = {x: cx - halfWidth, y: cy + halfHeight};
            
                let corners = [topLeft, topRight, bottomRight, bottomLeft];
            
                let rotatedCorners = [];
                for (let i = 0; i < corners.length; i++) {
                    rotatedCorners.push(rotatePoint(cx, cy, angle, corners[i].x, corners[i].y));
                }
                return rotatedCorners;
            }
        
            let cornersFigure1 = getRectangleCorners(x1, y1, figure1.width, figure1.height, angle);
            let cornersFigure2 = getRectangleCorners(x2, y2, figure2.width, figure2.height, angle);
        
            let minX = Math.min(x1, x2);
            let maxX = Math.max(x1, x2);
            let minY = Math.min(y1, y2);
            let maxY = Math.max(y1, y2);
        
            let allCorners = cornersFigure1.concat(cornersFigure2);
            for (let i = 0; i < allCorners.length; i++) {
                let point = allCorners[i];
                minX = Math.min(minX, point.x);
                maxX = Math.max(maxX, point.x);
                minY = Math.min(minY, point.y);
                maxY = Math.max(maxY, point.y);
            }
        
            // Возвращаем координаты прямоугольника
            return [minX, minY, maxX, maxY];
        }

        return calculateBoundingRectangle({x1: aPoints[0], y1: aPoints[1], x2: aPoints[2], y2: aPoints[3]}, shapeSizeAtStart, shapeSizeAtEnd);
    };
    CAnnotationLine.prototype.SetRect = function(aOrigRect) {
        let oViewer     = editor.getDocumentRenderer();
        let oDoc        = oViewer.getPDFDoc();

        oDoc.History.Add(new CChangesPDFAnnotRect(this, this.GetOrigRect(), aOrigRect));

        this._origRect = aOrigRect;

        let oXfrm = this.getXfrm();
        if (oXfrm) {
            oDoc.StartNoHistoryMode();
            this.spPr.xfrm.setExtX([aOrigRect[2] - aOrigRect[0]] * g_dKoef_pt_to_mm);
            this.spPr.xfrm.setExtY((aOrigRect[3] - aOrigRect[1]) * g_dKoef_pt_to_mm);
            oDoc.EndNoHistoryMode();
        }

        this.SetNeedRecalc(true);
        this.SetWasChanged(true);
    };
    CAnnotationLine.prototype.SetStrokeColor = function(aColor) {
        AscCommon.History.Add(new CChangesPDFAnnotStroke(this, this.GetStrokeColor(), aColor));

        this._strokeColor = aColor;

        let oRGB    = this.GetRGBColor(aColor);
        let oFill   = AscFormat.CreateSolidFillRGBA(oRGB.r, oRGB.g, oRGB.b, 255);
        let oLine   = this.spPr.ln;
        oLine.setFill(oFill);
        this.handleUpdateLn();
    };
    CAnnotationLine.prototype.SetOpacity = function(value) {
        this._opacity = value;
        this.SetWasChanged(true);

        let oLine = this.spPr.ln;
        oLine.Fill.transparent = value * 100 * 2.55;
        this.handleUpdateLn();
    };
    CAnnotationLine.prototype.GetDrawing = function() {
        return this.content.GetAllDrawingObjects()[0];
    };
    CAnnotationLine.prototype.SetWidth = function(nWidthPt) {
        this._width = nWidthPt; 

        nWidthPt = nWidthPt > 0 ? nWidthPt : 0.5;
        let oLine = this.spPr.ln;
        oLine.setW(nWidthPt * g_dKoef_pt_to_mm * 36000.0);
        this.handleUpdateLn();
    };
    CAnnotationLine.prototype.GetLinePoints = function() {
        return this._points;
    };
    CAnnotationLine.prototype.SetDoCaption = function(nType) {
        this._doCaption = nType;
    };
    CAnnotationLine.prototype.IsDoCaption = function() {
        return this._doCaption;
    };
    CAnnotationLine.prototype.SetLineStart = function(nType) {
        this._lineStart = nType;

        this.SetWasChanged(true);
        let oLine = this.spPr.ln;
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
        }

        oLine.headEnd.setType(nLineEndType);
        oLine.headEnd.setLen(AscFormat.LineEndSize.Mid);
        this.handleUpdateLn();
    };
    CAnnotationLine.prototype.SetLineEnd = function(nType) {
        this._lineEnd = nType;
        
        this.SetWasChanged(true);
        let oLine = this.spPr.ln;
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
        }

        oLine.tailEnd.setType(nLineEndType);
        oLine.tailEnd.setLen(AscFormat.LineEndSize.Mid);
        this.handleUpdateLn();
    };
    CAnnotationLine.prototype.GetLineStart = function() {
        return this._lineStart;
    };
    CAnnotationLine.prototype.GetLineEnd = function() {
        return this._lineEnd;
    };

    // for work with comments
    CAnnotationLine.prototype.RemoveComment = function() {
        if (!this.IsUseContentAsComment()) {
            AscPDF.CAnnotationFreeText.prototype.RemoveComment.call(this);
        }
        else {
            AscPDF.CAnnotationBase.prototype.RemoveComment.call(this);
        }
    };
    CAnnotationLine.prototype.SetContents = function(contents) {
        if (!this.IsUseContentAsComment()) {
            AscPDF.CAnnotationFreeText.prototype.SetContents.call(this, contents);
        }
        else {
            AscPDF.CAnnotationBase.prototype.SetContents.call(this, contents);
        }
    };
    CAnnotationLine.prototype.GetAscCommentData = function() {
        if (!this.IsUseContentAsComment()) {
            return AscPDF.CAnnotationFreeText.prototype.GetAscCommentData.call(this);
        }
        else {
            return AscPDF.CAnnotationBase.prototype.GetAscCommentData.call(this);
        }
    };
    //////////////////////////////////////

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
        let bDoCaption = this.IsDoCaption();
        if (bDoCaption) {
            memory.annotFlags |= (1 << 19);
        }
        
        // intent
        let nIntent = this.GetIntent();
        if (nIntent != null) {
            memory.annotFlags |= (1 << 20);
            memory.WriteByte(nIntent);
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
    };

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
            case AscPDF.LINE_END_TYPE.None:
                oSize.width = nLineW;
                oSize.height = nLineW;
            case AscPDF.LINE_END_TYPE.OpenArrow:
            case AscPDF.LINE_END_TYPE.ClosedArrow:
                oSize.width = 6 * nLineW;
                oSize.height = 3 * nLineW;
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
                oSize.width = 6 * nLineW;
                oSize.height = 6 * nLineW;
                break;
            case AscPDF.LINE_END_TYPE.Butt:
                oSize.width = 5 * nLineW;
                oSize.height = 1.5 * nLineW;
                break;
            case AscPDF.LINE_END_TYPE.Slash:
                oSize.width = 6 * nLineW;
                oSize.height = 3 * nLineW;
                break;
        }

        return oSize;
    }

    window["AscPDF"].CAnnotationLine    = CAnnotationLine;
    window["AscPDF"].LINE_END_TYPE      = LINE_END_TYPE;
})();

