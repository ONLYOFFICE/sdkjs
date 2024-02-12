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
        // internal
        TurnOffHistory();
    }
    CAnnotationFreeText.prototype.constructor = CAnnotationFreeText;
    AscFormat.InitClass(CAnnotationFreeText, AscFormat.CGroupShape, AscDFH.historyitem_type_GroupShape);
    Object.assign(CAnnotationFreeText.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationFreeText.prototype.getObjectType = function() {
        return -1;
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
    }
    CAnnotationFreeText.prototype.GetAlign = function() {
        return this._alignment;
    }
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
        if (this.spTree.length == 2) {
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
    CAnnotationFreeText.prototype.GetCallout = function() {
        return this._callout;
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
        
        this.AddToRedraw();
        this.SetWasChanged(true);
        this.SetDrawFromStream(false);
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

        let oViewer     = editor.getDocumentRenderer();
        let nPage       = this.GetPage();
        let aOrigRect   = this.GetOrigRect();

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;
        
        this.handleUpdatePosition();
        if (this.recalcInfo.recalculateGeometry)
            this.RefillGeometry();

        this.recalculate();
        this.updatePosition(aOrigRect[0] * g_dKoef_pix_to_mm * nScaleX, aOrigRect[1] * g_dKoef_pix_to_mm * nScaleY);
    };
    CAnnotationFreeText.prototype.RefillGeometry = function() {
        let oViewer = editor.getDocumentRenderer();
        let oDoc    = oViewer.getPDFDoc();
        
        let aOrigRect   = this.GetOrigRect();
        let aCallout    = this.GetCallout(); // координаты выходящей стрелки
        let aRD         = this.GetRectangleDiff() || [0, 0, 0, 0]; // отступ координат фигуры с текстом от ректа аннотации

        let nScaleY = oViewer.drawingPages[this.GetPage()].H / oViewer.file.pages[this.GetPage()].H / oViewer.zoom * g_dKoef_pix_to_mm;
        let nScaleX = oViewer.drawingPages[this.GetPage()].W / oViewer.file.pages[this.GetPage()].W / oViewer.zoom * g_dKoef_pix_to_mm;

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
            // перпендикуляр из точки выхода линии callout
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
        aFreeTextPoints.push(aFreeTextLine90);
        if (aCalloutLine.length != 0) {
            aFreeTextPoints.push(aCalloutLine);
        }

        let aShapeRectInMM = this.GetRect().map(function(measure) {
            return measure * g_dKoef_pix_to_mm;
        });

        oDoc.TurnOffHistory();
        fillShapeByPoints(aFreeTextPoints, aShapeRectInMM, this);

        this.recalcInfo.recalculateGeometry = false;
        this.CheckInnerShapesProps();
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

    function fillShapeByPoints(arrOfArrPoints, aShapeRect, oParentAnnot) {
        let xMin = aShapeRect[0];
        let yMin = aShapeRect[1];

        let oShapeRect = createInnerShape([arrOfArrPoints[0], arrOfArrPoints[1]]);
        // прямоугольнику добавляем cnx точки
        oShapeRect.spPr.geometry.AddCnx('_3cd4', 'hc', 't');
        oShapeRect.spPr.geometry.AddCnx('cd2', 'l', 'vc');
        oShapeRect.spPr.geometry.AddCnx('cd4', 'hc', 'b');
        oShapeRect.spPr.geometry.AddCnx('0', 'r', 'vc');
        oShapeRect.spPr.geometry.AddRect('l', 't', 'r', 'b');

        oShapeRect.spPr.xfrm.setOffX(oShapeRect.x - xMin);
        oShapeRect.spPr.xfrm.setOffY(oShapeRect.y - yMin);
        oShapeRect.setGroup(oParentAnnot);
        oParentAnnot.addToSpTree(0, oShapeRect);

        if (arrOfArrPoints[2]) {
            let oLineShape = createConnectorShape(arrOfArrPoints[2], oShapeRect);
            oLineShape.spPr.xfrm.setOffX(oLineShape.x - xMin);
            oLineShape.spPr.xfrm.setOffY(oLineShape.y - yMin);
            oLineShape.setGroup(oParentAnnot);
            oParentAnnot.addToSpTree(1, oLineShape);
        }
        
        oParentAnnot.x = xMin;
        oParentAnnot.y = yMin;
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
            let arrPoints = arrOfArrPoints[nPath];
            if (arrPoints.length == 0)
                continue;

            let bClosed     = false;
            let min_dist    = editor.WordControl.m_oDrawingDocument.GetMMPerDot(3);
            let oLastPoint  = arrPoints[arrPoints.length-1];
            let nLastIndex  = arrPoints.length-1;
            if(oLastPoint.bTemporary) {
                nLastIndex--;
            }
            if(nLastIndex > 1)
            {
                let dx = arrPoints[0].x - arrPoints[nLastIndex].x;
                let dy = arrPoints[0].y - arrPoints[nLastIndex].y;
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
            geometry.AddPathCommand(1, (((arrPoints[0].x - xMin) * kw) >> 0) + "", (((arrPoints[0].y - yMin) * kh) >> 0) + "");

            let oPt, nPt;
            for(nPt = 1; nPt < arrPoints.length; nPt++) {
                oPt = arrPoints[nPt];

                geometry.AddPathCommand(2,
                    (((oPt.x - xMin) * kw) >> 0) + "", (((oPt.y - yMin) * kh) >> 0) + ""
                );
            }

            if (arrPoints.length > 2)
                geometry.AddPathCommand(6);
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
            minX = Math.min(minX, rect[0]);
            minY = Math.min(minY, rect[1]);
            maxX = Math.max(maxX, rect[2]);
            maxY = Math.max(maxY, rect[3]);
        });
    
        return [minX, minY, maxX, maxY];
    }

    function initGroupShape(oParentAnnot) {
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
        
        oParentAnnot.spPr.xfrm.setOffX(0);
        oParentAnnot.spPr.xfrm.setOffY(0);
        oParentAnnot.spPr.xfrm.setExtX(Math.abs(xMax - xMin));
        oParentAnnot.spPr.xfrm.setExtY(Math.abs(yMax - yMin));
        oParentAnnot.setBDeleted(false);
        oParentAnnot.recalculate();
        oParentAnnot.brush = AscFormat.CreateNoFillUniFill();
    }

    function createInnerShape(aPaths) {
        function findMinRect(aPaths) {
            let x_min = aPaths[0][0].x, y_min = aPaths[0][0].y;
            let x_max = aPaths[0][0].x, y_max = aPaths[0][0].y;
        
            for (let nPath = 0; nPath < aPaths.length; nPath++) {
                let points = aPaths[nPath];
                for (let i = 1; i < points.length; i ++) {
                    x_min = Math.min(x_min, points[i].x);
                    x_max = Math.max(x_max, points[i].x);
                    y_min = Math.min(y_min, points[i].y);
                    y_max = Math.max(y_max, points[i].y);
                }
            }
            
        
            return [x_min, y_min, x_max, y_max];
        }

        let oShape = new AscFormat.CShape();

        let aShapeBounds = findMinRect(aPaths);

        let xMax = aShapeBounds[2];
        let xMin = aShapeBounds[0];
        let yMin = aShapeBounds[1];
        let yMax = aShapeBounds[3];

        oShape.setSpPr(new AscFormat.CSpPr());
        oShape.spPr.setParent(oShape);
        oShape.spPr.setXfrm(new AscFormat.CXfrm());
        oShape.spPr.xfrm.setParent(oShape.spPr);
        oShape.setWordShape(true);
        
        oShape.spPr.xfrm.setOffX(0);
        oShape.spPr.xfrm.setOffY(0);
        oShape.spPr.xfrm.setExtX(Math.abs(xMax - xMin));
        oShape.spPr.xfrm.setExtY(Math.abs(yMax - yMin));
        oShape.setStyle(AscFormat.CreateDefaultShapeStyle());
        oShape.setBDeleted(false);
        oShape.recalcInfo.recalculateGeometry = false;
        oShape.recalculate();
        oShape.brush = AscFormat.CreateNoFillUniFill();

        let geometry = generateGeometry(aPaths, [xMin, yMin, xMax, yMax]);
        oShape.spPr.setGeometry(geometry);
        oShape.updatePosition(xMin, yMin);

        return oShape;
    }

    function createConnectorShape(aPoints, oStartShape, oEndShape) {
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

        let oShape = new AscFormat.CConnectionShape();

        let aShapeBounds = findMinRect(aPoints);

        let xMax = aShapeBounds[2];
        let xMin = aShapeBounds[0];
        let yMin = aShapeBounds[1];
        let yMax = aShapeBounds[3];

        oShape.setSpPr(new AscFormat.CSpPr());
        oShape.spPr.setParent(oShape);
        oShape.spPr.setXfrm(new AscFormat.CXfrm());
        oShape.spPr.xfrm.setParent(oShape.spPr);
        oShape.setWordShape(true);
        
        oShape.spPr.xfrm.setOffX(0);
        oShape.spPr.xfrm.setOffY(0);
        oShape.spPr.xfrm.setExtX(Math.abs(xMax - xMin));
        oShape.spPr.xfrm.setExtY(Math.abs(yMax - yMin));
        oShape.setStyle(AscFormat.CreateDefaultShapeStyle());
        oShape.setBDeleted(false);
        oShape.recalcInfo.recalculateGeometry = false;
        oShape.recalculate();
        oShape.brush = AscFormat.CreateNoFillUniFill();

        let nv_sp_pr = new AscFormat.UniNvPr();
        nv_sp_pr.cNvPr.setId(0);
        oShape.setNvSpPr(nv_sp_pr);

        let nvUniSpPr = new AscFormat.CNvUniSpPr();
        if(oStartShape)
        {
            // nvUniSpPr.stCnxIdx = this.startConnectionInfo.idx;
            nvUniSpPr.stCnxIdx = 1;
            nvUniSpPr.stCnxId  = oStartShape.Id;
        }
        if(oEndShape)
        {
            // nvUniSpPr.endCnxIdx = this.endConnectionInfo.idx;
            nvUniSpPr.endCnxIdx = 1;
            nvUniSpPr.endCnxId  = oEndShape.Id;
        }
        oShape.nvSpPr.setUniSpPr(nvUniSpPr);

        let geometry = generateGeometry([aPoints], [xMin, yMin, xMax, yMax]);
        geometry.preset = "line";
        oShape.spPr.setGeometry(geometry);
        oShape.updatePosition(xMin, yMin);

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
})();

