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
        this.isInTextBox                    = false; // флаг, что внутри текстбокса
        this.defaultPerpLength              = 12; // длина выступающего перпендикуляра callout по умолчанию
        
        // internal
        TurnOffHistory();
    }
    CAnnotationFreeText.prototype.constructor = CAnnotationFreeText;
    AscFormat.InitClass(CAnnotationFreeText, AscFormat.CGroupShape, AscDFH.historyitem_type_GroupShape);
    Object.assign(CAnnotationFreeText.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationFreeText.prototype.GetCalloutExitPos = function(aTxBoxRect) {
        let aCallout = this.GetCallout();
        if (!aCallout)
            return CALLOUT_EXIT_POS.none;
        
        let aTextBoxRect = aTxBoxRect || this.GetTextBoxRect();

        // x1, y1 линии
        let oArrowEndPt = {
            x: aCallout[0 * 2],
            y: (aCallout[0 * 2 + 1])
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
    CAnnotationFreeText.prototype.GetArrowRect = function(aArrowPts) {
        let aCallout = this.GetCallout();
        if (!aCallout && !aArrowPts)
            return undefined;

        let nLineWidth = this.GetWidth() * g_dKoef_pt_to_mm * g_dKoef_mm_to_pix;

        let oLine = {
            x1: aArrowPts ? aArrowPts[0] : aCallout[1 * 2],
            y1: aArrowPts ? aArrowPts[1] : aCallout[1 * 2 + 1],
            x2: aArrowPts ? aArrowPts[2] : aCallout[0 * 2],
            y2: aArrowPts ? aArrowPts[3] : aCallout[0 * 2 + 1]
        };

        let oShapeEndSize = getFigureSize(this.GetLineEnd(), nLineWidth);
        
        function calculateBoundingRectangle(line, figure) {
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
        
            let cornersFigure = getRectangleCorners(x2, y2, figure.width, figure.height, angle);
        
            // Find minimum and maximum coordinates
            let minX = Math.min(x1, x2);
            let maxX = Math.max(x1, x2);
            let minY = Math.min(y1, y2);
            let maxY = Math.max(y1, y2);
        
            for (let i = 0; i < cornersFigure.length; i++) {
                let point = cornersFigure[i];
                minX = Math.min(minX, point.x);
                maxX = Math.max(maxX, point.x);
                minY = Math.min(minY, point.y);
                maxY = Math.max(maxY, point.y);
            }
        
            // Return the coordinates of the rectangle
            return [minX, minY, maxX, maxY];
        }

        return calculateBoundingRectangle(oLine, oShapeEndSize);
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
    CAnnotationFreeText.prototype.SetRectangleDiff = function(aDiff) {
        let oDoc = this.GetDocument();
        oDoc.History.Add(new CChangesPDFAnnotRD(this, this.GetRectangleDiff(), aDiff));

        this._rectDiff = aDiff;
        this.recalcGeometry();
        this.SetWasChanged(true);
        this.AddToRedraw();
    };
    /**
	 * Выставлят настройки ширины линии, цвета и тд для внутренних фигур.
	 * @constructor
    */
    CAnnotationFreeText.prototype.CheckInnerShapesProps = function() {
        let oStrokeColor = this.GetStrokeColor();
        if (oStrokeColor) {
            let oRGB    = this.GetRGBColor(oStrokeColor);
            let oFill   = AscFormat.CreateSolidFillRGBA(oRGB.r, oRGB.g, oRGB.b, 255);

            let oTxBoxShape = this.GetTextBoxShape();
            let oLine = oTxBoxShape.spPr.ln;
            if (this.GetWidth() == 0) {
                oLine.setFill(AscFormat.CreateNoFillUniFill());
            }
            else {
                oLine.setFill(oFill);
            }

            for (let i = 1; i < this.spTree.length; i++) {
                let oLine = this.spTree[i].spPr.ln;
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
        for (let i = 1; i < this.spTree.length; i++) {
            let oLine = this.spTree[i].spPr.ln;
            oLine.setW(nWidthPt * g_dKoef_pt_to_mm * 36000.0);
        }

        let nLineEndType = this.GetLineEnd();
        if (this.spTree[1] && this.spTree[1].getPresetGeom() == "line") {
            let oTargetSp = this.spTree[1];
            let oLine = oTargetSp.spPr.ln;
            oLine.setTailEnd(new AscFormat.EndArrow());
            let nInnerType = getInnerLineEndType(nLineEndType);
            
            oLine.tailEnd.setType(nInnerType);
            oLine.tailEnd.setLen(AscFormat.LineEndSize.Mid);
        }
    };
    CAnnotationFreeText.prototype.SetCallout = function(aCallout) {
        let oDoc = this.GetDocument();
        oDoc.History.Add(new CChangesFreeTextCallout(this, this.GetCallout(), aCallout));
        
        this._callout = aCallout;
        this.recalcGeometry();
        this.SetWasChanged(true);
        this.AddToRedraw();
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

        for (let i = 1; i < this.spTree.length; i++) {
            let oLine = this.spTree[i].spPr.ln;
            oLine.setW((nWidthPt || 0.5) * g_dKoef_pt_to_mm * 36000.0);
        }
    };
    CAnnotationFreeText.prototype.SetStrokeColor = function(aColor) {
        this._strokeColor = aColor;

        let oRGB    = this.GetRGBColor(aColor);
        let oFill   = AscFormat.CreateSolidFillRGBA(oRGB.r, oRGB.g, oRGB.b, 255);

        let oTxBoxShape = this.GetTextBoxShape();
        let oLine = oTxBoxShape.spPr.ln;
        if (this.GetWidth() == 0) {
            oLine.setFill(AscFormat.CreateNoFillUniFill());
        }
        else {
            oLine.setFill(oFill);
        }
        

        for (let i = 1; i < this.spTree.length; i++) {
            let oLine = this.spTree[i].spPr.ln;
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

        oFreeText.lazyCopy = true;

        oFreeText._pagePos = {
            x: this._pagePos.x,
            y: this._pagePos.y,
            w: this._pagePos.w,
            h: this._pagePos.h
        }
        oFreeText._origRect = this._origRect.slice();

        // this.copy2(oFreeText);
        // oFreeText.recalculate();

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
        oFreeText._callout = this._callout ? this._callout.slice() : undefined;
        oFreeText._rectDiff = this._rectDiff ? this._rectDiff.slice() : undefined;
        oFreeText.SetWasChanged(oFreeText.IsChanged());
        oFreeText.recalcGeometry();
        
        return oFreeText;
    };
    CAnnotationFreeText.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        let aRect = this.GetRect();
        this.spPr.xfrm.setOffX(aRect[0] * g_dKoef_pix_to_mm);
        this.spPr.xfrm.setOffY(aRect[1] * g_dKoef_pix_to_mm);
        
        if (this.recalcInfo.recalculateGeometry)
            this.RefillGeometry();

        this.recalculate();
        this.recalculateTransform();
        this.updateTransformMatrix();
        this.spTree.forEach(function(sp, idx) {
            idx == 0 && sp.recalculateContent();
            sp.recalculateTransform();
            sp.updateTransformMatrix();
        });
        
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
        this.SetReplies([]);
    };
    CAnnotationFreeText.prototype.SetContents = function(contents) {
        if (this.GetContents() == contents)
            return;

        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.GetDocument();
        let sCurContents    = this.GetContents();
        
        this._contents = contents;
        
        if (oViewer.IsOpenAnnotsInProgress == false && contents != sCurContents) {
            if (oDoc.History.UndoRedoInProgress == false) {
                oDoc.History.Add(new CChangesPDFAnnotContents(this, sCurContents, contents));
            }
        }

        this.SetWasChanged(true);
    };
    CAnnotationFreeText.prototype.SetNeedUpdateRC = function(bUpdate) {
        this._needUpdateRC = bUpdate;
    };
    CAnnotationFreeText.prototype.IsNeedUpdateRC = function() {
        return this._needUpdateRC;
    };
    CAnnotationFreeText.prototype.SetRichContents = function(aRCInfo) {
        let oDoc            = this.GetDocument();
        let oContent        = this.GetDocContent();
        oContent.ClearContent();
        
        let oLastUsedPara   = oContent.GetElement(0);
        oLastUsedPara.RemoveFromContent(0, oLastUsedPara.GetElementsCount());

        this._richContents = aRCInfo;
        oDoc.History.Add(new CChangesPDFFreeTextRC(this, this.GetRichContents(), aRCInfo));

        if (!aRCInfo) {
            this.SetNeedRecalc(true);
            this.SetNeedUpdateRC(false);
            return;
        }

        for (let i = 0; i < aRCInfo.length; i++) {
            let oRCInfo = aRCInfo[i];

            let oRun = new ParaRun(oLastUsedPara, false);

            let oRGB    = AscPDF.CBaseField.prototype.GetRGBColor(oRCInfo["color"]);
            let oRFonts = new CRFonts();

            if (oRCInfo["actual"]) {
                oRFonts.SetAll(oRCInfo["actual"], -1);
            }
            else if (oRCInfo["name"]) {
                oRFonts.SetAll(AscFonts.getEmbeddedFontPrefix() + oRCInfo["name"], -1);
            }
            else {
                oRFonts.SetAll(AscPDF.DEFAULT_FIELD_FONT, -1);
            }

            oRun.Set_Unifill(AscFormat.CreateSolidFillRGB(oRGB.r, oRGB.g, oRGB.b));
            oRun.SetBold(Boolean(oRCInfo["bold"]));
            oRun.SetItalic(Boolean(oRCInfo["italic"]));
            oRun.SetStrikeout(Boolean(oRCInfo["strikethrough"]));
            oRun.SetUnderline(Boolean(oRCInfo["underlined"]));
            oRun.SetFontSize(oRCInfo["size"]);
            oRun.Set_RFonts2(oRFonts);

            let oIterator = oRCInfo["text"].replace('\r', '').getUnicodeIterator();

            oLastUsedPara.AddToContentToEnd(oRun);
            oLastUsedPara.Set_Align(AscPDF.getInternalAlignByPdfType(oRCInfo["alignment"]));

            if (oRCInfo["text"].indexOf('\r') != -1) {
                oLastUsedPara = new AscWord.Paragraph(oContent, true);
                oContent.Internal_Content_Add(oContent.GetElementsCount(), oLastUsedPara);
            }

            while (oIterator.check()) {
                let runElement = AscPDF.codePointToRunElement(oIterator.value());
                oRun.Add(runElement);
                oIterator.next();
            }            
        }

        let _t = this;
        if (oDoc.Viewer.IsOpenAnnotsInProgress) {
            new Promise(function(resolve) {
                AscFonts.FontPickerByCharacter.checkTextLight(aRCInfo.reduce(function(accumulator, rc) {
                    return accumulator + rc["text"];
                }, ""), _t, resolve);
            }).then(function() {
                _t.SetNeedRecalc(true);
                _t.SetNeedUpdateRC(false);
            })
        }
        else {
            _t.SetNeedRecalc(true);
            _t.SetNeedUpdateRC(false);
        }
    };
    CAnnotationFreeText.prototype.GetRichContents = function(bCalced) {
        if (!bCalced)
            return this._richContents;

        let oContent = this.GetDocContent();
        let aRCInfo = [];

        for (let i = 0, nCount = oContent.GetElementsCount(); i < nCount; i++) {
            let oPara = oContent.GetElement(i);

            for (let j = 0, nRunsCount = oPara.GetElementsCount(); j < nRunsCount; j++) {
                let oRun = oPara.GetElement(j);
                let sText = oRun.GetText();
                if (sText) {
                    let oUniColor   = oRun.Pr.Unifill;
                    let oRGBA       = oUniColor ? oUniColor.fill.color.color.RGBA : null;
                    let aPdfColor   = oRGBA ? [oRGBA.R / 255, oRGBA.G / 255, oRGBA.B / 255] : [0, 0, 0];

                    let sFont   = oRun.Get_RFonts().Ascii.Name;
                    let isEmbed = false;
                    let prefix  = AscFonts.getEmbeddedFontPrefix();

                    if (sFont.startsWith(prefix)) {
                        sFont = sFont.substr(prefix.length);
                        isEmbed = true;
                    }
                        
                    let oRCInfo = {
                        "alignment":        AscPDF.getPdfTypeAlignByInternal(oRun.Paragraph.GetParagraphAlign()),
                        "bold":             oRun.Get_Bold(),
                        "italic":           oRun.Get_Italic(),
                        "strikethrough":    oRun.Get_Strikeout(),
                        "underlined":       oRun.Get_Underline(),
                        "size":             oRun.Get_FontSize(),
                        "color":            aPdfColor,
                        "text":             sText
                    };

                    if (isEmbed) {
                        oRCInfo["name"] = sFont;
                    }
                    else {
                        oRCInfo["actual"] = sFont;
                    }

                    aRCInfo.push(oRCInfo);
                }
            }

            if (aRCInfo[aRCInfo.length - 1])
                aRCInfo[aRCInfo.length - 1]["text"] += '\r';
        }

        return aRCInfo;
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
    CAnnotationFreeText.prototype.GetAllFonts = function(fontMap) {
        let aRCInfo = this.GetRichContents();
        fontMap = fontMap || {};

        if (!aRCInfo) {
            return fontMap;
        }

        for (let i = 0; i < aRCInfo.length; i++) {
            let fontName = AscPDF.DEFAULT_FIELD_FONT;
            if (aRCInfo[i]["actual"]) {
                fontName = aRCInfo[i]["actual"];
            }
            else if (aRCInfo[i]["name"]) {
                fontName = AscFonts.getEmbeddedFontPrefix() + aRCInfo[i]["name"];
            }
            fontMap[fontName] = true;
        }

        return fontMap;
    };
    CAnnotationFreeText.prototype.hitInPath = function(x,y) {
        for (let i = 0; i < this.spTree.length; i++) {
            if (this.spTree[i].hitInPath(x,y))
                return true;
        }

        return false;
    };
    CAnnotationFreeText.prototype.hitToHandles = function(x,y) {
        for (let i = 0; i < this.spTree.length; i++) {
            let nHandeNum = this.spTree[i].hitToHandles(x,y);
            if (nHandeNum != -1)
                return nHandeNum;
        }

        return -1;
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
    CAnnotationFreeText.prototype.Draw = function(oGraphicsPDF, oGraphicsWord) {
        if (this.IsHidden() == true)
            return;

        this.Recalculate();
        if (this.IsInTextBox())
            this.GetDocContent().RecalculateCurPos();

        this.draw(oGraphicsWord);
    };
    CAnnotationFreeText.prototype.onMouseDown = function(x, y, e) {
        let oDoc                = this.GetDocument();
        let oController         = oDoc.GetController();
        this.selectStartPage    = this.GetPage();
        
        if (this.IsInTextBox() == false) {
            if (this.selectedObjects.length <= this.spTree.length - 1) {
                let _t = this;
                // селектим все фигуры в группе (кроме перпендикулярной линии) если до сих пор не заселекчены
                oController.selection.groupSelection = this;
                this.selectedObjects.length = 0;

                this.spTree.forEach(function(sp) {
                    if (!(sp instanceof AscFormat.CConnectionShape)) {
                        sp.selectStartPage = _t.selectStartPage;
                        _t.selectedObjects.push(sp);
                    }
                });
            }
        }
        else {
            if (e.shiftKey) {
                this.GetDocContent().StartSelectionFromCurPos();
                oDoc.SelectionSetEnd(x, y, e);
            }
            else {
                oDoc.SelectionSetStart(x, y, e);
            }
        }
    };
    CAnnotationFreeText.prototype.SelectionSetStart = function(X, Y, e) {
        this.selectStartPage = this.GetPage();

        let oTextBoxShape   = this.GetTextBoxShape();
        let oContent        = this.GetDocContent();
        
        let oTransform  = oTextBoxShape.invertTransformText;
        let xContent    = oTransform.TransformPointX(X, 0);
        let yContent    = oTransform.TransformPointY(0, Y);

        oContent.Selection_SetStart(xContent, yContent, 0, e);
        oContent.RecalculateCurPos();
    };
    CAnnotationFreeText.prototype.SelectionSetEnd = function(X, Y, e) {
        let oTextBoxShape   = this.GetTextBoxShape();
        let oContent        = this.GetDocContent();
        
        let oTransform  = oTextBoxShape.invertTransformText;
        let xContent    = oTransform.TransformPointX(X, 0);
        let yContent    = oTransform.TransformPointY(0, Y);

        oContent.Selection_SetEnd(xContent, yContent, 0, e);
    };
    CAnnotationFreeText.prototype.MoveCursorLeft = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent()
        oContent.MoveCursorLeft(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CAnnotationFreeText.prototype.MoveCursorRight = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent()
        oContent.MoveCursorRight(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CAnnotationFreeText.prototype.MoveCursorDown = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent()
        oContent.MoveCursorDown(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CAnnotationFreeText.prototype.MoveCursorUp = function(isShiftKey, isCtrlKey) {
        let oContent = this.GetDocContent()
        oContent.MoveCursorUp(isShiftKey, isCtrlKey);
        oContent.RecalculateCurPos();
    };
    CAnnotationFreeText.prototype.SetInTextBox = function(isIn) {
        let oDoc = this.GetDocument();
        if (isIn) {
            this.selection.textSelection = this.GetTextBoxShape();
            oDoc.SetLocalHistory();
        }
        else {
            oDoc.SetGlobalHistory();
        }

        if (false == this.IsChanged()) {
            this.SetDrawFromStream(!isIn);
            this.AddToRedraw();
        }
        
        this.isInTextBox = isIn;
    };
    CAnnotationFreeText.prototype.IsInTextBox = function() {
        return this.isInTextBox;
    };
    CAnnotationFreeText.prototype.GetDocContent = function() {
        if (this.spTree[0])
            return this.spTree[0].getDocContent();

        return null;
    };
    CAnnotationFreeText.prototype.EnterText = function(aChars) {
        let oDoc        = this.GetDocument();
        let oContent    = this.GetDocContent();
        let oParagraph  = oContent.GetCurrentParagraph();

        oDoc.CreateNewHistoryPoint({objects: [this]});

        // удаляем текст в селекте
        if (oContent.IsSelectionUse())
            oContent.Remove(-1);

		for (let index = 0; index < aChars.length; ++index) {
			let oRun = AscPDF.codePointToRunElement(aChars[index]);
			if (oRun)
				oParagraph.AddToParagraph(oRun, true);
		}

        this.FitTextBox();
        this.SetNeedRecalc(true);
        this.SetNeedUpdateRC(true);
        oContent.RecalculateCurPos();

        return true;
    };
    /**
	 * Removes char in current position by direction.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CAnnotationFreeText.prototype.Remove = function(nDirection, isCtrlKey) {
        let oDoc = this.GetDocument();
        oDoc.CreateNewHistoryPoint({objects: [this]});

        let oContent = this.GetDocContent();
        oContent.Remove(nDirection, true, false, false, isCtrlKey);
        oContent.RecalculateCurPos();
        this.SetNeedRecalc(true);

        if (AscCommon.History.Is_LastPointEmpty()) {
            AscCommon.History.Remove_LastPoint();
        }
        else {
            this.SetNeedRecalc(true);
            this.SetNeedUpdateRC(true);
        }
    };
    CAnnotationFreeText.prototype.SelectAllText = function() {
        this.GetDocContent().SelectAll();
    };
    /**
	 * Exit from this annot.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
	 */
    CAnnotationFreeText.prototype.Blur = function() {
        let oDoc        = this.GetDocument();
        let oContent    = this.GetDocContent();

        oContent.SetApplyToAll(true);
		let sText = oContent.GetSelectedText(false, {NewLineParagraph: true, ParaSeparator: '\r'}).replace(/\r\n$/, '');
		oContent.SetApplyToAll(false);

        this.SetInTextBox(false);

        if (this.GetContents() != sText || this.IsNeedUpdateRC()) {
            oDoc.CreateNewHistoryPoint();
            this.GetContents() != sText && this.SetContents(sText);
            
            if (this.IsNeedUpdateRC()) {
                oDoc.History.Add(new CChangesPDFFreeTextRC(this, this.GetRichContents(), this.GetRichContents(true)));
                this.SetNeedUpdateRC(false);
            }

            oDoc.TurnOffHistory();
        }
        
        this.resetSelection();
        oDoc.GetDrawingDocument().TargetEnd();
    };

    CAnnotationFreeText.prototype.FitTextBox = function() {
        let oDocContent     = this.GetDocContent();
        let oTextBoxShape   = this.GetTextBoxShape();
        let nPage           = this.GetPage();
        oTextBoxShape.recalculateContent();                

        let oContentBounds  = oDocContent.GetContentBounds(nPage);
        let nContentH       = oContentBounds.Bottom - oContentBounds.Top;

        if (nContentH > oTextBoxShape.extY) {
            let oViewer = Asc.editor.getDocumentRenderer();
            let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
            let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

            let aCurTextBoxRect = this.GetTextBoxRect();
            
            // Находим новый textbox rect 
            let xMin = aCurTextBoxRect[0];
            let yMin = aCurTextBoxRect[1];
            let xMax = aCurTextBoxRect[2];
            let yMax = aCurTextBoxRect[3] + (nContentH - oTextBoxShape.extY + 0.5) * g_dKoef_mm_to_pix / nScaleY;

            let aNewTextBoxRect = [xMin, yMin, xMax, yMax];

            // находит точку выхода callout для нового ректа textbox
            let nCalloutExitPos = this.GetCalloutExitPos(aNewTextBoxRect);

            // пересчитываем callout
            let aNewCallout = this.GetCallout() ? this.GetCallout().slice() : null;
            switch (nCalloutExitPos) {
                case AscPDF.CALLOUT_EXIT_POS.left: {
                    // точка выхода (x3, y3)
                    aNewCallout[2 * 2]      = xMin;
                    aNewCallout[2 * 2 + 1]  = (yMin + (yMax - yMin) / 2);

                    // точка начала стрелки
                    aNewCallout[2 * 1]      = xMin - this.defaultPerpLength;
                    aNewCallout[2 * 1 + 1]  = (yMin + (yMax - yMin) / 2);
                    break;
                }
                case AscPDF.CALLOUT_EXIT_POS.top: {
                    aNewCallout[2 * 2]      = (xMin + (xMax - xMin) / 2);
                    aNewCallout[2 * 2 + 1]  = yMin;

                    aNewCallout[2 * 1]      = (xMin + (xMax - xMin) / 2);
                    aNewCallout[2 * 1 + 1]  = yMin - this.defaultPerpLength;
                    break;
                }
                case AscPDF.CALLOUT_EXIT_POS.right: {
                    aNewCallout[2 * 2]      = xMax;
                    aNewCallout[2 * 2 + 1]  = (yMin + (yMax - yMin) / 2);

                    aNewCallout[2 * 1]      = xMax + this.defaultPerpLength;
                    aNewCallout[2 * 1 + 1]  = (yMin + (yMax - yMin) / 2);
                    break;
                }
                case AscPDF.CALLOUT_EXIT_POS.bottom: {
                    aNewCallout[2 * 2]      = (xMin + (xMax - xMin) / 2);
                    aNewCallout[2 * 2 + 1]  = yMax;

                    aNewCallout[2 * 1]      = (xMin + (xMax - xMin) / 2);
                    aNewCallout[2 * 1 + 1]  = yMax + this.defaultPerpLength;
                    break;
                }
            }

            function findBoundingRectangle(points) {
                if (!points) {
                    return null;
                }

                let minX = points[0];
                let minY = points[1];
                let maxX = points[0];
                let maxY = points[1];
            
                for (let i = 2; i < points.length; i += 2) {
                    minX = Math.min(minX, points[i]);
                    maxX = Math.max(maxX, points[i]);
                    minY = Math.min(minY, points[i + 1]);
                    maxY = Math.max(maxY, points[i + 1]);
                }
            
                return [minX, minY, maxX, maxY];
            }

            // находим рект стрелки, учитывая окончание линии
            let aArrowRect = aNewCallout ? this.GetArrowRect([aNewCallout[2], aNewCallout[3], aNewCallout[0], aNewCallout[1]]) : null;

            // находим результирующий rect аннотации
            let aNewRect = AscPDF.unionRectangles([aArrowRect, aNewTextBoxRect, findBoundingRectangle(aNewCallout)]).map(function(measure, idx) {
                return idx % 2 ? measure * nScaleY : measure * nScaleX;
            });
            
            // пересчитываем RD.
            let aNewRD = [
                xMin - aNewRect[0] / nScaleX,
                yMin - aNewRect[1] / nScaleY,
                aNewRect[2] / nScaleX - xMax,
                aNewRect[3] / nScaleY - yMax
            ];

            aNewCallout && this.SetCallout(aNewCallout);
            this.SetRectangleDiff(aNewRD);
            this.SetRect(aNewRect);
        }
    };

    CAnnotationFreeText.prototype.GetTextBoxShape = function() {
        return this.spTree[0];
    };

    CAnnotationFreeText.prototype.onMouseUp = function(e) {
        this.GetDocument().ShowComment([this.GetId()]);

        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.GetDocument();
        let oDrDoc          = oDoc.GetDrawingDocument();
        this.isInMove       = false;

        this.selectStartPage = this.GetPage();
        let oPos    = oDrDoc.ConvertCoordsFromCursor2(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
        let X       = oPos.X;
        let Y       = oPos.Y;
        
        let oTextBoxShape = this.GetTextBoxShape();
        if (oTextBoxShape.hitInTextRect(X, Y)) {
            let oContent = this.GetDocContent();

            if (global_mouseEvent.ClickCount == 2) {
                let oTransform  = oTextBoxShape.invertTransformText;
                let xContent    = oTransform.TransformPointX(X, 0);
                let yContent    = oTransform.TransformPointY(0, Y);

                if (this.IsInTextBox() == false) {
                    this.selectedObjects.length = 0;
                    oContent.Selection_SetStart(xContent, yContent, 0, e);
                    oContent.RemoveSelection();
                    oContent.RecalculateCurPos();

                    oDrDoc.UpdateTargetFromPaint = true;
                    oDrDoc.TargetStart();
                    oDrDoc.showTarget(true);
                    this.SetInTextBox(true);
                    this.FitTextBox();
                }
                else {
                    oContent.SelectAll();
                    if (oContent.IsSelectionEmpty() == false)
                        oViewer.Api.WordControl.m_oDrawingDocument.TargetEnd();
                    else
                        oContent.RemoveSelection();
                }
            }

            if (oContent.IsSelectionEmpty())
                oContent.RemoveSelection();

            oViewer.onUpdateOverlay();
        }
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
                    sp.selectStartPage = _t.selectStartPage;
                    _t.selectedObjects.push(sp);
                }
            });
        }
        // иначе move 
        else {
            // если попали в стрелку, тогда селектим группу, т.к. будем перемещать всю аннотацию целиком
            if (this.spTree.length == 1 || (this.spTree[1] && sShapeId == this.spTree[1].GetId() && this.spTree[1].getPresetGeom() == "line")) {
                this.selectedObjects.length = 0;
                oDrawingObjects.selection.groupSelection = null;
                oDrawingObjects.selectedObjects.length = 0;
                oDrawingObjects.selectedObjects.push(this);
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
            memory.WriteByte(nIntent);
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
    CAnnotationFreeText.prototype.Get_AbsolutePage = function() {
        return this.GetPage();
    };

    function fillShapeByPoints(arrOfArrPoints, aShapeRect, oParentAnnot) {
        let xMin = aShapeRect[0];
        let yMin = aShapeRect[1];

        let oRectShape = createInnerShape(arrOfArrPoints[0], oParentAnnot.spTree[0], oParentAnnot);
        oRectShape.spPr.xfrm.setFlipH(false);
        oRectShape.spPr.xfrm.setFlipV(false);

        if (oRectShape.getDocContent() == null) {
            oRectShape.createTextBody();
            oRectShape.txBody.bodyPr.setInsets(0.5,0.5,0.5,0.5);
            oRectShape.txBody.bodyPr.horzOverflow = AscFormat.nHOTClip;
            oRectShape.txBody.bodyPr.vertOverflow = AscFormat.nVOTClip;
        }

        oRectShape.setTxBox(true);
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

        let oTxBoxShape = createInnerShape([{x: 0, y: 0}], null, oParentFreeText);
        oTxBoxShape.spPr.xfrm.setFlipH(false);
        oTxBoxShape.spPr.xfrm.setFlipV(false);

        oTxBoxShape.createTextBody();
        oTxBoxShape.txBody.bodyPr.setInsets(0.5,0.5,0.5,0.5);
        oTxBoxShape.txBody.bodyPr.horzOverflow = AscFormat.nHOTClip;
        oTxBoxShape.txBody.bodyPr.vertOverflow = AscFormat.nVOTClip;

        oParentFreeText.addToSpTree(0, oTxBoxShape);
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

        let oShape = oExistShape ? oExistShape : new AscPDF.CPdfShape();
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
        oShape.setBDeleted(false);
        oShape.recalcInfo.recalculateGeometry = false;
        oShape.setGroup(oParentAnnot);
        oShape.spPr.setLn(new AscFormat.CLn());
        oShape.recalculateTransform();
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
        oShape.setBDeleted(false);
        oShape.recalcInfo.recalculateGeometry = false;
        oShape.setGroup(oParentAnnot);
        oShape.spPr.setLn(new AscFormat.CLn());
        oShape.recalculateTransform();
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

    window["AscPDF"].CAnnotationFreeText    = CAnnotationFreeText;
    window["AscPDF"].FREE_TEXT_INTENT_TYPE  = window["AscPDF"]["FREE_TEXT_INTENT_TYPE"] = FREE_TEXT_INTENT_TYPE;
    FREE_TEXT_INTENT_TYPE['FreeText']           = FREE_TEXT_INTENT_TYPE.FreeText;
    FREE_TEXT_INTENT_TYPE['FreeTextCallout']    = FREE_TEXT_INTENT_TYPE.FreeTextCallout;
    
    window["AscPDF"].CALLOUT_EXIT_POS       = CALLOUT_EXIT_POS;
})();

