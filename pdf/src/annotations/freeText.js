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

    let FREE_TEXT_INTENT_TYPE = {
        FreeText:           0,
        FreeTextCallout:    1,
        FreeTextTypeWriter: 2
    }

    let CALLOUT_EXIT_POS = {
        none:  -1,
        left:   0,
        top:    1,
        right:  2,
        bottom: 3
    }

    /**
	 * Class representing a free text annotation.
	 * @constructor
    */
    function CAnnotationFreeText(sName, nPage, aRect, oDoc)
    {
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.FreeText, nPage, aRect, oDoc);
        AscFormat.CGroupShape.call(this);
        initGroupShape(this);
        
        this.GraphicObj     = this;
        
        this._popupOpen     = false;
        this._popupRect     = undefined;
        this._richContents  = undefined;
        this._rotate        = undefined;
        this._state         = undefined;
        this._stateModel    = undefined;
        this._width         = undefined;
        this._points        = undefined;
        this._intent        = undefined;
        this._lineEnd       = undefined;
        this._callout       = undefined;
        this._alignment     = undefined;
        this._defaultStyle  = undefined;

        this.recalcInfo.recalculateGeometry = true;
        this.defaultPerpLength = 12; // длина выступающего перпендикуляра callout по умолчанию

        // internal
        TurnOffHistory();
    }
    CAnnotationFreeText.prototype.constructor = CAnnotationFreeText;
    AscFormat.InitClass(CAnnotationFreeText, AscFormat.CGroupShape, AscDFH.historyitem_type_GroupShape);
    Object.assign(CAnnotationFreeText.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationFreeText.prototype.GetCalloutExitPos = function() {
        let aCallout = this.GetCallout();
        if (!aCallout)
            return CALLOUT_EXIT_POS.none;
        
        let aTextBoxRect = this.GetTextBoxRect();

        // x1, y1 линии
        let oArrowEndPt = {
            x: aCallout[1 * 2],
            y: (aCallout[1 * 2 + 1])
        };

        if (oArrowEndPt.x < aTextBoxRect[0]) {
            return CALLOUT_EXIT_POS.left;
        }
        else if (oArrowEndPt.x > aTextBoxRect[2]) {
            return CALLOUT_EXIT_POS.right;
        }
        else if (oArrowEndPt.x >= aTextBoxRect[0] && oArrowEndPt.x <= aTextBoxRect[2]) {
            if (oArrowEndPt.y < aTextBoxRect[1]) {
                return CALLOUT_EXIT_POS.top;
            }
            else if (oArrowEndPt.y > aTextBoxRect[3]) {
                return CALLOUT_EXIT_POS.bottom;
            }
        }
    };

    CAnnotationFreeText.prototype.IsNeedDrawFromStream = function() {
        return false;
    };

    CAnnotationFreeText.prototype.IsFreeText = function() {
        return true;
    };
    CAnnotationFreeText.prototype.SetDefaultStyle = function(sStyle) {
        this._defaultStyle = sStyle;
    };
    CAnnotationFreeText.prototype.GetDefaultStyle = function() {
        return this._defaultStyle;
    };
    CAnnotationFreeText.prototype.SetAlign = function(nType) {
        this._alignment = nType;
    };
    CAnnotationFreeText.prototype.GetAlign = function() {
        return this._alignment;
    };
    CAnnotationFreeText.prototype.SetLineEnd = function(nType) {
        this._lineEnd = nType;
        
        this.SetWasChanged(true);

        if (this.spTree.length == 3) {
            let oTargetSp = this.spTree[1];
            let oLine = oTargetSp.pen;
            oLine.setTailEnd(new AscFormat.EndArrow());
            let nLineEndType = getInnerLineEndType(nType);
            

            oLine.tailEnd.setType(nLineEndType);
            oLine.tailEnd.setLen(AscFormat.LineEndSize.Mid);
        }
    };
    CAnnotationFreeText.prototype.GetLineEnd = function() {
        return this._lineEnd;
    };
    /**
	 * Выставлят настройки ширины линии, цвета и тд для внутренних фигур.
	 * @constructor
    */
    CAnnotationFreeText.prototype.CheckInnerShapesProps = function() {
        let oStrokeColor    = this.GetStrokeColor();
        if (oStrokeColor) {
            let oRGB            = this.GetRGBColor(oStrokeColor);
            let oFill           = AscFormat.CreateSolidFillRGBA(oRGB.r, oRGB.g, oRGB.b, 255);
            for (let i = 0; i < this.spTree.length; i++) {
                let oLine = this.spTree[i].pen;
                oLine.setFill(oFill);
            }
        }
        
        let oFillColor = this.GetFillColor();
        if (oFillColor) {
            let oRGB    = this.GetRGBColor(oFillColor);
            let oFill   = AscFormat.CreateSolidFillRGBA(oRGB.r, oRGB.g, oRGB.b, 255);
            for (let i = 0; i < this.spTree.length; i++) {
                this.spTree[i].setFill(oFill);
            }
        }

        let nWidthPt = this.GetWidth();
        nWidthPt = nWidthPt > 0 ? nWidthPt : 0.5;
        for (let i = 0; i < this.spTree.length; i++) {
            let oLine = this.spTree[i].pen;
            oLine.setW(nWidthPt * g_dKoef_pt_to_mm * 36000.0);
        }

        let nLineEndType = this.GetLineEnd();
        if (this.spTree[1] && this.spTree[1].getPresetGeom() == "line") {
            let oTargetSp = this.spTree[1];
            let oLine = oTargetSp.pen;
            oLine.setTailEnd(new AscFormat.EndArrow());
            let nInnerType = getInnerLineEndType(nLineEndType);
            
            oLine.tailEnd.setType(nInnerType);
            oLine.tailEnd.setLen(AscFormat.LineEndSize.Mid);
        }

        this.recalculate();
    };
    CAnnotationFreeText.prototype.GetMinShapeRect = function() {
        let oViewer     = editor.getDocumentRenderer();
        let nLineWidth  = this.GetWidth() * g_dKoef_pt_to_mm * g_dKoef_mm_to_pix;
        let aVertices     = this.GetVertices();
        let nPage       = this.GetPage();

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        let shapeAtStart    = getFigureSize(this.GetLineStart(), nLineWidth);
        let shapeAtEnd      = getFigureSize(this.GetLineEnd(), nLineWidth);

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
            
                let corners = [
                    {x: cx - halfWidth, y: cy - halfHeight},
                    {x: cx + halfWidth, y: cy - halfHeight},
                    {x: cx + halfWidth, y: cy + halfHeight},
                    {x: cx - halfWidth, y: cy + halfHeight}
                ];
            
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
        
            for (let i = 0; i < cornersFigure1.length; i++) {
                let point = cornersFigure1[i];
                minX = Math.min(minX, point.x);
                maxX = Math.max(maxX, point.x);
                minY = Math.min(minY, point.y);
                maxY = Math.max(maxY, point.y);
            }
        
            for (let j = 0; j < cornersFigure2.length; j++) {
                let point = cornersFigure2[j];
                minX = Math.min(minX, point.x);
                maxX = Math.max(maxX, point.x);
                minY = Math.min(minY, point.y);
                maxY = Math.max(maxY, point.y);
            }
        
            // Возвращаем координаты прямоугольника
            return [minX * nScaleX, minY * nScaleY, maxX * nScaleX, maxY * nScaleY];
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
        
            return [x_min * nScaleX, y_min * nScaleY, x_max * nScaleX, y_max * nScaleY];
        }

        // находим ректы исходных точек. Стартовой линии учитывая lineStart фигуру, и такую же для конца
        // далее нахоим объединения всех прямоугольников для получения результирующего
        let aSourceRect     = findBoundingRectangle(aVertices);
        let aStartLineRect  = calculateBoundingRectangle(oStartLine, shapeAtStart, {width: 0, height: 0});
        let aEndLineRect    = calculateBoundingRectangle(oEndLine, {width: 0, height: 0} , shapeAtEnd);

        return unionRectangles([aSourceRect, aStartLineRect, aEndLineRect]);
    };
    CAnnotationFreeText.prototype.SetCallout = function(aCallout) {
        this.recalcGeometry();
        
        this._callout = aCallout;
    };
    CAnnotationFreeText.prototype.GetCallout = function(bScaled) {
        if (bScaled != true || !this._callout)
            return this._callout;

        let oViewer = Asc.editor.getDocumentRenderer();
        let nPage   = this.GetPage();
        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        return this._callout.slice().map(function(measure, idx) {
            return idx % 2 == 0 ? measure * nScaleX : measure * nScaleY;
        });
    };
    CAnnotationFreeText.prototype.SetWidth = function(nWidthPt) {
        this._width = nWidthPt; 

        nWidthPt = nWidthPt > 0 ? nWidthPt : 0.5;
        for (let i = 0; i < this.spTree.length; i++) {
            let oLine = this.spTree[i].pen;
            oLine.setW(nWidthPt * g_dKoef_pt_to_mm * 36000.0);
        }
    };
    CAnnotationFreeText.prototype.SetStrokeColor = function(aColor) {
        this._strokeColor = aColor;

        let oRGB    = this.GetRGBColor(aColor);
        let oFill   = AscFormat.CreateSolidFillRGBA(oRGB.r, oRGB.g, oRGB.b, 255);

        for (let i = 0; i < this.spTree.length; i++) {
            let oLine = this.spTree[i].pen;
            oLine.setFill(oFill);
        }
    };
    CAnnotationFreeText.prototype.SetFillColor = function(aColor) {
        this._fillColor = aColor;

        let oRGB    = this.GetRGBColor(aColor);
        let oFill   = AscFormat.CreateSolidFillRGBA(oRGB.r, oRGB.g, oRGB.b, 255);
        for (let i = 0; i < this.spTree.length; i++) {
            this.spTree[i].setFill(oFill);
        }
    };
    CAnnotationFreeText.prototype.SetRect = function(aRect) {
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
        this.spPr.xfrm.setOffX(aRect[0] * g_dKoef_pix_to_mm);
        this.spPr.xfrm.setOffY(aRect[1] * g_dKoef_pix_to_mm);
        this.updateTransformMatrix();

        this.recalcGeometry();
        this.SetNeedRecalc(true);
        this.SetWasChanged(true);
    };
    CAnnotationFreeText.prototype.GetTextBoxRect = function(bScale) {
        let oViewer     = Asc.editor.getDocumentRenderer();
        let aOrigRect   = this.GetOrigRect();
        let aRD         = this.GetRectangleDiff() || [0, 0, 0, 0]; // отступ координат фигуры с текстом от ректа аннотации
        let nPage       = this.GetPage();

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        let xMin = bScale ? nScaleX * (aOrigRect[0] + aRD[0]) : (aOrigRect[0] + aRD[0]);
        let yMin = bScale ? nScaleY * (aOrigRect[1] + aRD[1]) : (aOrigRect[1] + aRD[1]);
        let xMax = bScale ? nScaleX * (aOrigRect[2] - aRD[2]) : (aOrigRect[2] - aRD[2]);
        let yMax = bScale ? nScaleY * (aOrigRect[3] - aRD[3]) : (aOrigRect[3] - aRD[3]);

        return [xMin, yMin, xMax, yMax]
    };
    CAnnotationFreeText.prototype.LazyCopy = function() {
        let oDoc = this.GetDocument();
        oDoc.TurnOffHistory();

        let oFreeText = new CAnnotationFreeText(AscCommon.CreateGUID(), this.GetPage(), this.GetOrigRect().slice(), oDoc);

        oFreeText._pagePos = {
            x: this._pagePos.x,
            y: this._pagePos.y,
            w: this._pagePos.w,
            h: this._pagePos.h
        }
        oFreeText._origRect = this._origRect.slice();

        this.copy2(oFreeText);
        oFreeText.recalculate();

        oFreeText.pen = new AscFormat.CLn();
        oFreeText._apIdx = this._apIdx;
        oFreeText._originView = this._originView;
        oFreeText.SetOriginPage(this.GetOriginPage());
        oFreeText.SetAuthor(this.GetAuthor());
        oFreeText.SetModDate(this.GetModDate());
        oFreeText.SetCreationDate(this.GetCreationDate());
        oFreeText.SetWidth(this.GetWidth());
        oFreeText.SetStrokeColor(this.GetStrokeColor().slice());
        oFreeText.SetContents(this.GetContents());
        oFreeText.SetFillColor(this.GetFillColor());
        oFreeText.SetLineEnd(this.GetLineEnd());
        oFreeText.recalcInfo.recalculatePen = false;
        oFreeText.recalcInfo.recalculateGeometry = false;
        oFreeText._callout = this._callout.slice();
        oFreeText._rectDiff = this._rectDiff.slice();
        oFreeText.SetWasChanged(oFreeText.IsChanged());
        
        return oFreeText;
    };
    CAnnotationFreeText.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        if (this.recalcInfo.recalculateGeometry)
            this.RefillGeometry();

        this.recalculate();
        this.SetNeedRecalc(false);
    };
    CAnnotationFreeText.prototype.RefillGeometry = function() {
        let oViewer = editor.getDocumentRenderer();
        let oDoc    = oViewer.getPDFDoc();
        
        let aOrigRect   = this.GetOrigRect();
        let aCallout    = this.GetCallout(); // координаты выходящей стрелки
        let aRD         = this.GetRectangleDiff() || [0, 0, 0, 0]; // отступ координат фигуры с текстом от ректа аннотации
        let nPage       = this.GetPage();

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom * g_dKoef_pix_to_mm;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom * g_dKoef_pix_to_mm;

        let aFreeTextPoints = [];
        let aFreeTextRect   = []; // прямоугольник
        let aFreeTextLine90 = []; // перпендикуляр к прямоуольнику (x3, y3 - x2, y2) точки из callout

        // левый верхний
        aFreeTextRect.push({
            x: (aOrigRect[0] + aRD[0]) * nScaleX,
            y: (aOrigRect[1] + aRD[1]) * nScaleY
        });
        // правый верхний
        aFreeTextRect.push({
            x: (aOrigRect[2] - aRD[2]) * nScaleX,
            y: (aOrigRect[1] + aRD[1]) * nScaleY
        });
        // правый нижний
        aFreeTextRect.push({
            x: (aOrigRect[2] - aRD[2]) * nScaleX,
            y: (aOrigRect[3] - aRD[3]) * nScaleY
        });
        // левый нижний
        aFreeTextRect.push({
            x: (aOrigRect[0] + aRD[0]) * nScaleX,
            y: (aOrigRect[3] - aRD[3]) * nScaleY
        });

        if (aCallout && aCallout.length == 6) {
            // точка выхода callout
            aFreeTextLine90.push({
                x: (aCallout[2 * 2]) * nScaleX,
                y: (aCallout[2 * 2 + 1]) * nScaleY
            });
            aFreeTextLine90.push({
                x: (aCallout[2 * 1]) * nScaleX,
                y: (aCallout[2 * 1 + 1]) * nScaleY
            });
        }
        
        let aCalloutLine = [];
        if (aCallout) {
            // x2, y2 линии
            aCalloutLine.push({
                x: aCallout[1 * 2] * nScaleX,
                y: (aCallout[1 * 2 + 1]) * nScaleY
            });
            // x1, y1 линии
            aCalloutLine.push({
                x: aCallout[0 * 2] * nScaleX,
                y: (aCallout[0 * 2 + 1]) * nScaleY
            });
        }

        aFreeTextPoints.push(aFreeTextRect);
        if (aCalloutLine.length != 0)
            aFreeTextPoints.push(aCalloutLine);
        if (aFreeTextLine90.length != 0)
            aFreeTextPoints.push(aFreeTextLine90);

        let aShapeRectInMM = this.GetRect().map(function(measure) {
            return measure * g_dKoef_pix_to_mm;
        });

        oDoc.TurnOffHistory();
        fillShapeByPoints(aFreeTextPoints, aShapeRectInMM, this);

        this.recalcInfo.recalculateGeometry = false;
        this.CheckInnerShapesProps();
    };
    CAnnotationFreeText.prototype.recalcGeometry = function () {
        this.recalcInfo.recalculateGeometry = true;
	};
    CAnnotationFreeText.prototype.RemoveComment = function() {
        let oDoc = this.GetDocument();

        oDoc.CreateNewHistoryPoint();
        this.SetReplies([]);
        oDoc.TurnOffHistory();
    };
    CAnnotationFreeText.prototype.SetContents = function(contents) {
        if (this.GetContents() == contents)
            return;

        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.GetDocument();
        let oCurContents    = this.GetContents();

        this._contents  = contents;
        
        if (oDoc.History.UndoRedoInProgress == false && oViewer.IsOpenAnnotsInProgress == false) {
            oDoc.History.Add(new CChangesPDFAnnotContents(this, oCurContents, contents));
        }
        
        this.SetWasChanged(true);
    };
    CAnnotationFreeText.prototype.SetReplies = function(aReplies) {
        let oDoc = this.GetDocument();
        let oViewer = editor.getDocumentRenderer();

        if (oDoc.History.UndoRedoInProgress == false && oViewer.IsOpenAnnotsInProgress == false) {
            oDoc.History.Add(new CChangesPDFAnnotReplies(this, this._replies, aReplies));
        }
        this._replies = aReplies;

        let oThis = this;
        aReplies.forEach(function(reply) {
            reply.SetReplyTo(oThis);
        });

        if (aReplies.length != 0)
            oDoc.CheckComment(this);
        else
            editor.sync_RemoveComment(this.GetId());
    };
    CAnnotationFreeText.prototype.hitInPath = function(x,y) {
        for (let i = 0; i < this.spTree.length; i++) {
            if (this.spTree[i].hitInPath(x,y))
                return true;
        }

        return false;
    };
    CAnnotationFreeText.prototype.hitInInnerArea = function(x, y) {
        for (let i = 0; i < this.spTree.length; i++) {
            if (this.spTree[i].hitInInnerArea(x,y))
                return true;
        }

        return false;
    };
    CAnnotationFreeText.prototype.GetAscCommentData = function() {
        let oAscCommData = new Asc["asc_CCommentDataWord"](null);
        if (this._replies.length == 0)
            return oAscCommData;

        let oMainComm = this._replies[0];
        oAscCommData.asc_putText(oMainComm.GetContents());
        oAscCommData.asc_putOnlyOfficeTime(oMainComm.GetModDate().toString());
        oAscCommData.asc_putUserId(editor.documentUserId);
        oAscCommData.asc_putUserName(oMainComm.GetAuthor());
        oAscCommData.asc_putSolved(false);
        oAscCommData.asc_putQuoteText("");
        oAscCommData.m_sUserData = oMainComm.GetApIdx();

        this._replies.forEach(function(reply, index) {
            if (index == 0)
                return;
            
            oAscCommData.m_aReplies.push(reply.GetAscCommentData());
        });

        return oAscCommData;
    };
    CAnnotationFreeText.prototype.onMouseDown = function(e) {
        let oViewer         = editor.getDocumentRenderer();
        let oDrawingObjects = oViewer.DrawingObjects;
        
        let _t = this;
        // селектим все фигуры в группе если до сих пор не заселекчены
        oDrawingObjects.selection.groupSelection = this;
        this.selectedObjects.length = 0;

        this.spTree.forEach(function(sp) {
            if (!(sp instanceof AscFormat.CConnectionShape)) {
                sp.selectStartPage = 0;
                _t.selectedObjects.push(sp);
            }
        });

        this.selectStartPage = this.GetPage();
    };
    CAnnotationFreeText.prototype.onAfterMove = function() {
        this.onMouseDown();
        this.isInMove = false;
    };
    CAnnotationFreeText.prototype.onPreMove = function(e) {
        if (this.isInMove)
            return;

        this.isInMove = true; // происходит ли resize/move действие

        let oViewer         = editor.getDocumentRenderer();
        let oDrawingObjects = oViewer.DrawingObjects;
        let oDoc            = this.GetDocument();
        let oDrDoc          = oDoc.GetDrawingDocument();

        this.selectStartPage = this.GetPage();
        let oPos    = oDrDoc.ConvertCoordsFromCursor2(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
        let X       = oPos.X;
        let Y       = oPos.Y;

        let pageObject = oViewer.getPageByCoords3(AscCommon.global_mouseEvent.X - oViewer.x, AscCommon.global_mouseEvent.Y - oViewer.y);

        let oCursorInfo = oDrawingObjects.getGraphicInfoUnderCursor(oPos.DrawPage, X, Y);
        if (oCursorInfo.cursorType == null) {
            this.isInMove = false;
            return;
        }

        let isResize    = oCursorInfo.cursorType.indexOf("resize") != -1 ? true : false;
        let sShapeId    = oCursorInfo.objectId;

        // если фигуры в селекте группы, тогда смотрим в какую попали
        this.selectedObjects.length = 0;

        let _t = this;
        // если в handles то телектим внутри группы нужную фигуру
        if (isResize) {
            this.spTree.forEach(function(sp) {
                if (sp.GetId() == sShapeId) {
                    sp.selectStartPage = 0;
                    _t.selectedObjects.push(sp);
                }
            });
        }
        // иначе move 
        else {
            // если попали в стрелку, тогда селектим группу, т.к. будем перемещать всю аннотацию целиком
            if (this.spTree[1] && sShapeId == this.spTree[1].GetId() && this.spTree[1].getPresetGeom() == "line") {
                _t.selectedObjects.length = 0;
                oDrawingObjects.selection.groupSelection = null;
                oDrawingObjects.selectedObjects.length = 0;
                oDrawingObjects.selectedObjects.push(this.spTree[1]);
            }
            // если попали в textbox, тогда селектим textbox фигуру внутри группы, т.к. будем перемещать только её
            else if (this.spTree[0] && sShapeId == this.spTree[0].GetId()) {
                this.selectedObjects.length                 = 0;
                oDrawingObjects.selection.groupSelection    = this;
                this.selectedObjects.push(this.spTree[0]);
            }
        }

        oDrawingObjects.OnMouseDown(e, X, Y, pageObject.index);
    };
    CAnnotationFreeText.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        // alignment
        let nAlign = this.GetAlign();
        if (nAlign != null)
            memory.WriteByte(nAlign);

        // rectangle diff
        let aRD = this.GetRectangleDiff();
        if (aRD) {
            memory.annotFlags |= (1 << 15);
            for (let i = 0; i < 4; i++) {
                memory.WriteDouble(aRD[i]);
            }
        }

        // callout
        let aCallout = this.GetCallout();
        if (aCallout != null) {
            memory.annotFlags |= (1 << 16);
            memory.WriteLong(aCallout.length);
            for (let i = 0; i < aCallout.length; i++)
                memory.WriteDouble(aCallout[i]);
        }

        // default style
        let sDefaultStyle = this.GetDefaultStyle();
        if (sDefaultStyle != null) {
            memory.annotFlags |= (1 << 17);
            memory.WriteString(sDefaultStyle);
        }

        // line end
        let nLE = this.GetLineEnd();
        if (nLE != null) {
            memory.annotFlags |= (1 << 18);
            memory.WriteByte(nLE);
        }
            
        // intent
        let nIntent = this.GetIntent();
        if (nIntent != null) {
            memory.annotFlags |= (1 << 20);
            memory.WriteDouble(nIntent);
        }

        let nEndPos = memory.GetCurPosition();
        memory.Seek(memory.posForFlags);
        memory.WriteLong(memory.annotFlags);
        
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };

    // shape methods
    CAnnotationFreeText.canRotate = function() {
        return false;
    };

    function fillShapeByPoints(arrOfArrPoints, aShapeRect, oParentAnnot) {
        let xMin = aShapeRect[0];
        let yMin = aShapeRect[1];

        let oRectShape = createInnerShape(arrOfArrPoints[0], oParentAnnot.spTree[0], oParentAnnot);
        oRectShape.createTextBoxContent();
        oRectShape.textBoxContent.GetElement(0).GetElement(0).AddText('1234');
        let oPara       = oRectShape.textBoxContent.GetElement(0);
        let oApiPara    = editor.private_CreateApiParagraph(oPara);

        oApiPara.SetColor(55, 55, 55, false);
        oPara.RecalcCompiledPr(true);

        // oRectShape.textBoxContent.Recalculate_Page(0, true);
        // прямоугольнику добавляем cnx точки

        oRectShape.setGroup(oParentAnnot);
        if (!oParentAnnot.spTree[0])
            oParentAnnot.addToSpTree(0, oRectShape);
        
        let oLineShape;
        // координаты стрелки
        if (arrOfArrPoints[1]) {
            // флипаем стрелку если соблюдаются условия (зачем? Чтобы handles были с нужной стороны - так уж устроены CShape)
            let aArrowPts   = arrOfArrPoints[1].slice();
            let bFlipH      = false;
            let bFlipV      = false;
            if (aArrowPts[0].x < aArrowPts[1].x) {
                let nTmpX = aArrowPts[0].x;
                aArrowPts[0].x = aArrowPts[1].x;
                aArrowPts[1].x = nTmpX;
                bFlipH = true;
            }
            if (aArrowPts[0].y < aArrowPts[1].y) {
                let nTmpY = aArrowPts[0].y;
                aArrowPts[0].y = aArrowPts[1].y;
                aArrowPts[1].y = nTmpY;
                bFlipV = true;
            }

            oLineShape = createInnerShape(aArrowPts, oParentAnnot.spTree[1], oParentAnnot);
            oLineShape.spPr.xfrm.setFlipH(bFlipH);
            oLineShape.spPr.xfrm.setFlipV(bFlipV);
            if (bFlipH || bFlipV) {
                oLineShape.recalculateTransform();
                oLineShape.updateTransformMatrix();
            }

            if (!oParentAnnot.spTree[1])
                oParentAnnot.addToSpTree(1, oLineShape);
        }

        if (arrOfArrPoints[2]) {
            let oConnShape = createConnectorShape(arrOfArrPoints[2], oParentAnnot.spTree[2], oParentAnnot);
            if (!oParentAnnot.spTree[2])
                oParentAnnot.addToSpTree(2, oConnShape);
        }
        
        oParentAnnot.x = xMin;
        oParentAnnot.y = yMin;
        return oParentAnnot;
    }

    function generateGeometry(aPoints, aBounds, oGeometry) {
        let xMin = aBounds[0];
        let yMin = aBounds[1];
        let xMax = aBounds[2];
        let yMax = aBounds[3];

        let geometry = oGeometry ? oGeometry : new AscFormat.Geometry();
        if (oGeometry) {
            oGeometry.pathLst = [];
        }

        let bClosed     = false;
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
        
        geometry.AddPathCommand(0,undefined, undefined, undefined, pathW, pathH);
        geometry.AddPathCommand(1, (((aPoints[0].x - xMin) * kw) >> 0) + "", (((aPoints[0].y - yMin) * kh) >> 0) + "");

        let oPt, nPt;
        for(nPt = 1; nPt < aPoints.length; nPt++) {
            oPt = aPoints[nPt];

            geometry.AddPathCommand(2,
                (((oPt.x - xMin) * kw) >> 0) + "", (((oPt.y - yMin) * kh) >> 0) + ""
            );
        }

        geometry.preset = null;
        geometry.rectS = null;

        if (aPoints.length > 2)
            geometry.AddPathCommand(6);
        else
            geometry.preset = "line";
        
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
            minX = Math.min(minX, rect[0]);
            minY = Math.min(minY, rect[1]);
            maxX = Math.max(maxX, rect[2]);
            maxY = Math.max(maxY, rect[3]);
        });
    
        return [minX, minY, maxX, maxY];
    }

    function initGroupShape(oParentFreeText) {
        let aShapeRectInMM = oParentFreeText.GetRect().map(function(measure) {
            return measure * g_dKoef_pix_to_mm;
        });
        let xMax = aShapeRectInMM[2];
        let xMin = aShapeRectInMM[0];
        let yMin = aShapeRectInMM[1];
        let yMax = aShapeRectInMM[3];

        oParentFreeText.setSpPr(new AscFormat.CSpPr());
        oParentFreeText.spPr.setParent(oParentFreeText);
        oParentFreeText.spPr.setXfrm(new AscFormat.CXfrm());
        oParentFreeText.spPr.xfrm.setParent(oParentFreeText.spPr);
        
        oParentFreeText.spPr.xfrm.setOffX(xMin);
        oParentFreeText.spPr.xfrm.setOffY(yMin);
        oParentFreeText.spPr.xfrm.setExtX(Math.abs(xMax - xMin));
        oParentFreeText.spPr.xfrm.setExtY(Math.abs(yMax - yMin));
        oParentFreeText.setBDeleted(false);
        oParentFreeText.recalculate();
        oParentFreeText.updateTransformMatrix();

        oParentFreeText.brush = AscFormat.CreateNoFillUniFill();
    }

    function createInnerShape(aPoints, oExistShape, oParentAnnot) {
        function findMinRect(points) {
            let x_min = points[0].x, y_min = points[0].y;
            let x_max = points[0].x, y_max = points[0].y;
        
            for (let i = 1; i < points.length; i ++) {
                x_min = Math.min(x_min, points[i].x);
                x_max = Math.max(x_max, points[i].x);
                y_min = Math.min(y_min, points[i].y);
                y_max = Math.max(y_max, points[i].y);
            }
        
            return [x_min, y_min, x_max, y_max];
        }

        
        let aShapeBounds = findMinRect(aPoints);

        let xMax = aShapeBounds[2];
        let xMin = aShapeBounds[0];
        let yMin = aShapeBounds[1];
        let yMax = aShapeBounds[3];

        let oShape = oExistShape ? oExistShape : new AscFormat.CShape();
        if (!oExistShape) {
            oShape.setSpPr(new AscFormat.CSpPr());
            oShape.spPr.setParent(oShape);
            oShape.spPr.setXfrm(new AscFormat.CXfrm());
            oShape.spPr.xfrm.setParent(oShape.spPr);
        }
        
        let aAnnotRect = oParentAnnot.GetRect().map(function(measure) {
            return measure * g_dKoef_pix_to_mm;
        });
        oShape.spPr.xfrm.setOffX(Math.abs(xMin - aAnnotRect[0]));
        oShape.spPr.xfrm.setOffY(Math.abs(yMin - aAnnotRect[1]));
        oShape.spPr.xfrm.setExtX(Math.abs(xMax - xMin));
        oShape.spPr.xfrm.setExtY(Math.abs(yMax - yMin));
        oShape.setStyle(AscFormat.CreateDefaultShapeStyle());
        oShape.setBDeleted(false);
        oShape.recalcInfo.recalculateGeometry = false;
        oShape.setGroup(oParentAnnot);
        oShape.recalculate();
        oShape.updateTransformMatrix();
        oShape.brush = AscFormat.CreateNoFillUniFill();

        let geometry = generateGeometry(aPoints, [xMin, yMin, xMax, yMax]);
        oShape.spPr.setGeometry(geometry);

        return oShape;
    }

    function createConnectorShape(aPoints, oExistShape, oParentAnnot) {
        function findMinRect(points) {
            let x_min = points[0].x, y_min = points[0].y;
            let x_max = points[0].x, y_max = points[0].y;
        
            for (let i = 1; i < points.length; i ++) {
                x_min = Math.min(x_min, points[i].x);
                x_max = Math.max(x_max, points[i].x);
                y_min = Math.min(y_min, points[i].y);
                y_max = Math.max(y_max, points[i].y);
            }
        
            return [x_min, y_min, x_max, y_max];
        }

        let oShape = oExistShape || new AscFormat.CConnectionShape();
        let aShapeBounds = findMinRect(aPoints);

        let xMax = aShapeBounds[2];
        let xMin = aShapeBounds[0];
        let yMin = aShapeBounds[1];
        let yMax = aShapeBounds[3];

        if (!oExistShape) {
            oShape.setSpPr(new AscFormat.CSpPr());
            oShape.spPr.setParent(oShape);
            oShape.spPr.setXfrm(new AscFormat.CXfrm());
            oShape.spPr.xfrm.setParent(oShape.spPr);
        }
        
        let aAnnotRect = oParentAnnot.GetRect().map(function(measure) {
            return measure * g_dKoef_pix_to_mm;
        });
        
        oShape.spPr.xfrm.setOffX(Math.abs(xMin - aAnnotRect[0]));
        oShape.spPr.xfrm.setOffY(Math.abs(yMin - aAnnotRect[1]));
        oShape.spPr.xfrm.setExtX(Math.abs(xMax - xMin));
        oShape.spPr.xfrm.setExtY(Math.abs(yMax - yMin));
        oShape.setStyle(AscFormat.CreateDefaultShapeStyle());
        oShape.setBDeleted(false);
        oShape.recalcInfo.recalculateGeometry = false;
        oShape.setGroup(oParentAnnot);
        oShape.recalculate();
        oShape.updateTransformMatrix();
        oShape.brush = AscFormat.CreateNoFillUniFill();

        let nv_sp_pr = new AscFormat.UniNvPr();
        nv_sp_pr.cNvPr.setId(0);
        oShape.setNvSpPr(nv_sp_pr);

        let nvUniSpPr = new AscFormat.CNvUniSpPr();
        if(oParentAnnot.spTree[1])
        {
            // nvUniSpPr.stCnxIdx = this.startConnectionInfo.idx;
            nvUniSpPr.stCnxIdx = 1;
            nvUniSpPr.stCnxId  = oParentAnnot.spTree[0].Id;
        }
        if(oParentAnnot.spPr[0])
        {
            // nvUniSpPr.endCnxIdx = this.endConnectionInfo.idx;
            nvUniSpPr.endCnxIdx = 1;
            nvUniSpPr.endCnxId  = oParentAnnot.spPr[1].Id;
        }
        oShape.nvSpPr.setUniSpPr(nvUniSpPr);

        let geometry = generateGeometry(aPoints, [xMin, yMin, xMax, yMax]);
        geometry.preset = "line";
        oShape.spPr.setGeometry(geometry);

        return oShape;
    }

    function getInnerLineEndType(nPdfType) {
        let nInnerType;
        switch (nPdfType) {
            case AscPDF.LINE_END_TYPE.None:
                nInnerType = AscFormat.LineEndType.None;
                break;
            case AscPDF.LINE_END_TYPE.OpenArrow:
                nInnerType = AscFormat.LineEndType.Arrow;
                break;
            case AscPDF.LINE_END_TYPE.Diamond:
                nInnerType = AscFormat.LineEndType.Diamond;
                break;
            case AscPDF.LINE_END_TYPE.Circle:
                nInnerType = AscFormat.LineEndType.Oval;
                break;
            case AscPDF.LINE_END_TYPE.ClosedArrow:
                nInnerType = AscFormat.LineEndType.Triangle;
                break;
            case AscPDF.LINE_END_TYPE.ROpenArrow:
                nInnerType = AscFormat.LineEndType.ReverseArrow;
                break;
            case AscPDF.LINE_END_TYPE.RClosedArrow:
                nInnerType = AscFormat.LineEndType.ReverseTriangle;
                break;
            case AscPDF.LINE_END_TYPE.Butt:
                nInnerType = AscFormat.LineEndType.Butt;
                break;
            case AscPDF.LINE_END_TYPE.Square:
                nInnerType = AscFormat.LineEndType.Square;
                break;
            case AscPDF.LINE_END_TYPE.Slash:
                nInnerType = AscFormat.LineEndType.Slash;
                break;
            default:
                nInnerType = AscFormat.LineEndType.Arrow;
                break;
        }

        return nInnerType;
    }
    
    function TurnOffHistory() {
        if (AscCommon.History.IsOn() == true)
            AscCommon.History.TurnOff();
    }

    window["AscPDF"].CAnnotationFreeText = CAnnotationFreeText;
    window["AscPDF"].CALLOUT_EXIT_POS = CALLOUT_EXIT_POS;
})();

