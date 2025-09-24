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
    function CAnnotationLine(sName, aRect, oDoc)
    {
        AscPDF.CPdfShape.call(this);
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Line, aRect, oDoc);
        
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
        this._lineStart     = AscPDF.LINE_END_TYPE.None;
        this._lineEnd       = AscPDF.LINE_END_TYPE.None;
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
        let aPoints = this.GetLinePoints();
        let oGeometry = this.getGeometry();

        AscCommon.History.StartNoHistoryMode();

        if (!this.HasAdjustments()) {
            if (!oGeometry) {
                oGeometry = AscFormat.CreateGeometry("line");
                this.spPr.setGeometry(oGeometry);
            }

            let bFlipH = false;
            let bFlipV = false;

            if (aPoints[0] > aPoints[2]) {
                bFlipH = true;
            }

            if (aPoints[1] > aPoints[3]) {
                bFlipV = true;
            }

            let oXfrm = this.getXfrm();
            oXfrm.setFlipH(bFlipH);
            oXfrm.setFlipV(bFlipV);

            oGeometry.gdLstInfo = [];
            oGeometry.AddGuide("hCustom", AscFormat.FORMULA_TYPE_VALUE, "100");
            oGeometry.AddGuide("wCustom", AscFormat.FORMULA_TYPE_MOD, "w", "h", "0");
            oGeometry.AddGuide("hCustomD2", AscFormat.FORMULA_TYPE_MULT_DIV, "hCustom", "1", "2");
            oGeometry.AddGuide("wCustomD2", AscFormat.FORMULA_TYPE_MULT_DIV, "wCustom", "1", "2");

            oGeometry.AddGuide("lCustom", AscFormat.FORMULA_TYPE_PLUS_MINUS, "wd2", "0", "wCustomD2");
            oGeometry.AddGuide("tCustom", AscFormat.FORMULA_TYPE_PLUS_MINUS, "hd2", "0", "hCustomD2");
            oGeometry.AddGuide("rCustom", AscFormat.FORMULA_TYPE_PLUS_MINUS, "wd2", "wCustomD2", "0");
            oGeometry.AddGuide("bCustom", AscFormat.FORMULA_TYPE_PLUS_MINUS, "hd2", "hCustomD2", "0");

            oGeometry.rectS = {
                l: "lCustom",
                t: "tCustom",
                r: "rCustom",
                b: "bCustom"
            };

            oGeometry.setPreset("line");
        }
        else {
            let nLeaderExt = this.GetLeaderExtend();
            let nLeaderLength = this.GetLeaderLength();
            let nRad = Math.atan2(aPoints[3] - aPoints[1], aPoints[2] - aPoints[0]);
            let nHeight = Math.abs(nLeaderLength) + Math.abs(nLeaderExt);

            oGeometry = this.getGeometry();
            if (!oGeometry) {
                oGeometry = new AscFormat.Geometry();
            }
            else {
                nHeight = this.getXfrmExtY() * g_dKoef_mm_to_pt;
            }

            let nAdj1 = Math.abs(nLeaderLength) / nHeight * 100 * 1000;
            let nAdj2 = Math.abs(nLeaderExt) / nHeight * 100 * 1000;
            
            oGeometry.gdLstInfo = [];
            oGeometry.pathLst = [];
            oGeometry.AddRect("0", "txT", "w", "txB");

            // left leg
            oGeometry.AddPathCommand(0, undefined, "norm");
            oGeometry.AddPathCommand(1, "0", "0");
            oGeometry.AddPathCommand(2, "0", "sumH");
            oGeometry.AddPathCommand(2, "0", "sumH");
            oGeometry.AddPathCommand(6);
            
            // right leg
            oGeometry.AddPathCommand(0, undefined, "norm");
            oGeometry.AddPathCommand(1, "w", "0");
            oGeometry.AddPathCommand(2, "w", "sumH");
            oGeometry.AddPathCommand(2, "w", "sumH");
            oGeometry.AddPathCommand(6);
            
            // main line
            oGeometry.AddPathCommand(0, undefined, "norm");
            oGeometry.AddPathCommand(1, "0", "upperH");
            oGeometry.AddPathCommand(2, "w", "upperH");

            // formulas
            if (nLeaderLength < 0) {
                oGeometry.AddGuide("upperH", 0, "h", "adj1", "100000");
                oGeometry.AddGuide("lowerH", 0, "h", "adj2", "100000");
            }
            else {
                oGeometry.AddGuide("upperH", 0, "h", "adj2", "100000");
                oGeometry.AddGuide("lowerH", 0, "h", "adj1", "100000");
            }
            
            oGeometry.AddGuide("sumH", 1, "upperH", "lowerH", "0");
            oGeometry.AddGuide("txT", 1, "upperH", "0", "360000");
            oGeometry.AddGuide("txB", 1, "upperH", "360000", "0");
            
            if (nLeaderLength < 0) {
                oGeometry.AddHandleXY(undefined, undefined, undefined, "adj1", undefined, "100000", "0", "upperH");
                oGeometry.AddHandleXY(undefined, undefined, undefined, "adj2", undefined, "1000000", "w", "sumH");
            }
            else {
                oGeometry.AddHandleXY(undefined, undefined, undefined, "adj1", undefined, "1000000", "w", "sumH");
                oGeometry.AddHandleXY(undefined, undefined, undefined, "adj2", undefined, "100000", "0", "upperH");
            }
            

            oGeometry.AddAdj("adj1", 15, String(nAdj1));
            oGeometry.AddAdj("adj2", 15, String(nAdj2));

            this.spPr.setGeometry(oGeometry);
            this.getXfrm().setRot(nRad);
        }

        AscCommon.History.EndNoHistoryMode();
    };
    CAnnotationLine.prototype.HasAdjustments = function() {
        let nLeaderExt = this.GetLeaderExtend();
        if (nLeaderExt == undefined) {
            return false;
        }

        return true;
    };
    CAnnotationLine.prototype.canRotate = function() {
        return this.HasAdjustments();
    };
    CAnnotationLine.prototype.RecalcSizes = function() {
        let aPoints = this.GetLinePoints();
        
        if (!this.HasAdjustments()) {
            let extX = (Math.abs(aPoints[2] - aPoints[0])) * g_dKoef_pt_to_mm;
            let extY = (Math.abs(aPoints[3] - aPoints[1])) * g_dKoef_pt_to_mm;

            AscCommon.History.StartNoHistoryMode();
            this.spPr.xfrm.setOffX((Math.min(aPoints[0], aPoints[2])) * g_dKoef_pt_to_mm);
            this.spPr.xfrm.setOffY((Math.min(aPoints[1], aPoints[3])) * g_dKoef_pt_to_mm);

            this.spPr.xfrm.setExtX(extX);
            this.spPr.xfrm.setExtY(extY);
            if (!this.txBody) {
                AscCommon.History.EndNoHistoryMode();
                return;
            }

            let bFlipH = aPoints[0] > aPoints[2];
            let bFlipV = aPoints[1] > aPoints[3];

            let rot = Math.atan2(Math.abs(aPoints[3] - aPoints[1]), Math.abs(aPoints[2] - aPoints[0]));

            if (bFlipH) rot = Math.PI - rot;
            if (bFlipV) rot = -rot + Math.PI;
            let nRad = rot * AscFormat.cToDeg >> 0;

            let oBodyPr = this.txBody.bodyPr.createDuplicate();
            oBodyPr.setRot(nRad);
            this.txBody.setBodyPr(oBodyPr);

            AscCommon.History.EndNoHistoryMode();
            return;
        }
        
        let nLeaderExt = this.GetLeaderExtend();
        let nLeaderLength = this.GetLeaderLength();

        function computeRectOffset(points, width, height) {
            let x1 = points[0];
            let y1 = points[1];
            let x2 = points[2];
            let y2 = points[3];

            let dx = x2 - x1;
            let dy = y2 - y1;
            let theta = Math.atan2(dy, dx);

            let cx = (x1 + x2) / 2;
            let cy = (y1 + y2) / 2;

            let offsetX = 0;
            let offsetY = -height / 2;

            let rotatedOffsetX = offsetX * Math.cos(theta) - offsetY * Math.sin(theta);
            let rotatedOffsetY = offsetX * Math.sin(theta) + offsetY * Math.cos(theta);

            let centerX = cx - rotatedOffsetX;
            let centerY = cy - rotatedOffsetY;

            let topLeftX = centerX - width / 2;
            let topLeftY = centerY - height / 2;

            return { x: topLeftX, y: topLeftY };
        }

        function getOppositeSide(points, h) {
            let x1 = points[0];
            let y1 = points[1];
            let x2 = points[2];
            let y2 = points[3];

            let dx  = x2 - x1,
                dy  = y2 - y1,
                len = Math.sqrt(dx * dx + dy * dy) || 1;

            let nx = -dy / len,
                ny =  dx / len;

            nx *= -1;
            ny *= -1;

            return [
                x1 + nx * h,
                y1 + ny * h,
                x2 + nx * h,
                y2 + ny * h
            ];
        }

        var extX = Math.sqrt(Math.pow(aPoints[2] - aPoints[0], 2) + Math.pow(aPoints[3] - aPoints[1], 2));
        let extY = (Math.abs(nLeaderLength) + Math.abs(nLeaderExt));

        if (nLeaderLength > 0) {
            aPoints = getOppositeSide(aPoints, extY);
        }

        let pos = computeRectOffset(aPoints, extX, extY);
        
        AscCommon.History.StartNoHistoryMode();
        this.spPr.xfrm.setOffX(pos.x * g_dKoef_pt_to_mm);
        this.spPr.xfrm.setOffY(pos.y * g_dKoef_pt_to_mm);

        this.spPr.xfrm.setExtX(extX * g_dKoef_pt_to_mm);
        this.spPr.xfrm.setExtY(extY * g_dKoef_pt_to_mm);

        AscCommon.History.EndNoHistoryMode();
        
        this.SetNeedRecalcSizes(false);
    };
    CAnnotationLine.prototype.SetPosition = function(x, y) {
        let aCurRect = this.GetRect();

        let nOldX = aCurRect[0];
        let nOldY = aCurRect[1];

        let nDeltaX = x - nOldX;
        let nDeltaY = y - nOldY;

        if (0 == nDeltaX && 0 == nDeltaY) {
            return;
        }

        for (let i = 0; i < this._points.length; i+=2) {
            this._points[i] += nDeltaX;
            this._points[i+1] += nDeltaY;
        }

        let nWidth  = aCurRect[2] - aCurRect[0];
        let nHeight = aCurRect[3] - aCurRect[1];
        
        this._rect[0] = x;
        this._rect[1] = y;
        this._rect[2] = x + nWidth;
        this._rect[3] = y + nHeight;

        AscCommon.History.Add(new CChangesPDFAnnotPos(this, [aCurRect[0], aCurRect[1]], [x, y]));

        this.SetNeedRecalc(true);
        this.SetNeedRecalcSizes(true);
        this.SetWasChanged(true, false);
    };
    CAnnotationLine.prototype.SetLeaderLineOffset = function(nValue) {
        this._leaderLineOffset = nValue;
    };
    CAnnotationLine.prototype.GetLeaderLineOffset = function() {
        return this._leaderLineOffset;
    };
    CAnnotationLine.prototype.SetLeaderLength = function(nValue) {
        if (this._leaderLength == nValue) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFLineAnnotLeaderLength(this, this._leaderLength, nValue));

        this._leaderLength = nValue;
        this.SetWasChanged(true);
        this.recalcGeometry();
        this.SetNeedRecalc(true);
    };
    CAnnotationLine.prototype.GetLeaderLength = function() {
        return this._leaderLength;
    };
    CAnnotationLine.prototype.SetLeaderExtend = function(nValue) {
        if (this._leaderExtend == nValue) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFLineAnnotLeaderExt(this, this._leaderExtend, nValue));

        this._leaderExtend = nValue;
        this.SetWasChanged(true);
        this.recalcGeometry();
        this.SetNeedRecalc(true);
    };
    CAnnotationLine.prototype.GetLeaderExtend = function() {
        return this._leaderExtend;
    };
    CAnnotationLine.prototype.IsNeedRecalc = function() {
        return this._needRecalc;
    };
    CAnnotationLine.prototype.SetLinePoints = function(aPoints) {
        AscCommon.History.Add(new CChangesPDFLinePoints(this, this._points, aPoints));

        this._points = aPoints;

        this.SetWasChanged(true);
        this.recalcGeometry();
        this.SetNeedRecalc(true);
    };
    CAnnotationLine.prototype.GetLinePoints = function() {
        return this._points;
    };
    CAnnotationLine.prototype.Copy = function(isForMove) {
        let oCopy = AscPDF.CAnnotationBase.prototype.Copy.call(this, isForMove);

        let aRC = this.GetRichContents(true);
        let aPoints = this.GetLinePoints();
        
        oCopy.SetLinePoints(aPoints.slice());
        oCopy.SetLeaderExtend(this.GetLeaderExtend());
        oCopy.SetLeaderLength(this.GetLeaderLength());
        oCopy.SetLineEnd(this.GetLineEnd());
        oCopy.SetLineStart(this.GetLineStart());
        oCopy.SetDoCaption(this.IsDoCaption());
        aRC && oCopy.SetRichContents(aRC.slice());

        return oCopy;
    };
    CAnnotationLine.prototype.IsLine = function() {
        return true;
    };
    CAnnotationLine.prototype.private_CalcBoundingRect = function() {
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
    CAnnotationLine.prototype.SetStrokeColor = function(aColor) {
        AscCommon.History.Add(new CChangesPDFAnnotStroke(this, this.GetStrokeColor(), aColor));

        this._strokeColor = aColor;

        let oRGB    = this.GetRGBColor(aColor);
        let oFill   = AscFormat.CreateSolidFillRGBA(oRGB.r, oRGB.g, oRGB.b, 255);
        let oLine   = this.spPr.ln;
        oLine.setFill(oFill);
        this.handleUpdateLn();
    };
    CAnnotationLine.prototype.GetDrawing = function() {
        return this.content.GetAllDrawingObjects()[0];
    };
    CAnnotationLine.prototype.GetLinePoints = function() {
        return this._points;
    };
    CAnnotationLine.prototype.SetDoCaption = function(bCaption) {
        if (this._doCaption == bCaption) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFLineAnnotDoCaption(this, this._doCaption, bCaption));

        this._doCaption = bCaption;
        
        this.SetWasChanged(true);
        this.private_UpdateCaption();
        this.SetNeedRecalc(true);
    };
    CAnnotationLine.prototype.private_UpdateRect = function() {
        this.recalcBounds();
        this.recalcGeometry();
        this.SetNeedRecalcSizes(true);
        this.Recalculate(true);

        let nLineW = this.GetWidth() * g_dKoef_pt_to_mm;

        let oGrBounds = this.bounds;

        let aRect = [
            (oGrBounds.l - nLineW) * g_dKoef_mm_to_pt,
            (oGrBounds.t - nLineW) * g_dKoef_mm_to_pt,
            (oGrBounds.r + nLineW) * g_dKoef_mm_to_pt,
            (oGrBounds.b + nLineW) * g_dKoef_mm_to_pt
        ];

        this.SetRect(aRect);
    };
    CAnnotationLine.prototype.IsDoCaption = function() {
        return this._doCaption;
    };
    CAnnotationLine.prototype.private_UpdateCaption = function() {
        if (this.IsDoCaption()) {
            if (!this.txBody) {
                this.createTextBody();
                this.txBody.bodyPr.setInsets(2,2,2,2);
                this.txBody.bodyPr.horzOverflow = AscFormat.nHOTClip;
                this.txBody.bodyPr.vertOverflow = AscFormat.nVOTClip;
                this.txBody.bodyPr.setAnchor(AscFormat.VERTICAL_ANCHOR_TYPE_CENTER);
                // this.txBody.bodyPr.wrap = AscFormat.nTWTNone; 
                this.setTxBox(true);
            }

            let oContent = this.GetDocContent();
            oContent.ClearContent();
            
            let oLastUsedPara = oContent.GetElement(0);
            oLastUsedPara.RemoveFromContent(0, oLastUsedPara.GetElementsCount());
            oLastUsedPara.Set_Align(AscCommon.align_Center);
            
            let oShd = new CDocumentShd();
            oShd.Value = Asc.c_oAscShdClear;
            oShd.Color.Set(255, 255, 255, false);
            oShd.Fill = new CDocumentColor(255, 255, 255, false);

            let aRCInfo = this.GetRichContents();
            if (!aRCInfo) {
                let sContents = this.GetContents();
                if (!sContents) {
                    return;
                }

                aRCInfo = [{
                    "text": sContents,
                    "actual": "Arial",
                    "color": [0, 0, 0]
                }];
            }

            function setRunPr(oRun, oRCInfo) {
                let oRFonts = new CRFonts();
                let oRGB = AscPDF.CBaseField.prototype.GetRGBColor(oRCInfo["color"]);
                let nVertAlign = AscCommon.vertalign_Baseline;
                if (typeof(oRCInfo["vertical"]) == "number") {
                    if (oRCInfo["vertical"] >= 0) {
                        nVertAlign = AscCommon.vertalign_SuperScript;
                    }
                    else if (oRCInfo["vertical"] < 0) {
                        nVertAlign = AscCommon.vertalign_SubScript;
                    }
                }

                if (oRCInfo["actual"]) {
                    oRFonts.SetAll(oRCInfo["actual"], -1);
                }
                else if (oRCInfo["name"]) {
                    oRFonts.SetAll(oRCInfo["name"], -1);
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
                oRun.SetVertAlign(nVertAlign);
                oRun.Set_Shd(oShd);
            }
            
            for (let i = 0; i < aRCInfo.length; i++) {
                let oRCInfo = aRCInfo[i];

                let oRun = new ParaRun(oLastUsedPara, false);

                setRunPr(oRun, oRCInfo);
                oLastUsedPara.AddToContentToEnd(oRun);
                for (let nChar = 0; nChar < oRCInfo["text"].length; nChar++) {
                    let nCharCode = oRCInfo["text"][nChar].charCodeAt(0);
                    
                    if (nCharCode == 13) {
                        oLastUsedPara.Correct_Content();
                        oLastUsedPara.AddToParagraph(new AscWord.ParaTextPr(oRun.GetTextPr()));

                        oLastUsedPara = new AscWord.Paragraph(oContent, true);
                        oContent.Internal_Content_Add(oContent.GetElementsCount(), oLastUsedPara);

                        oRun = new ParaRun(oLastUsedPara, false);
                        setRunPr(oRun, oRCInfo);
                        oLastUsedPara.AddToContentToEnd(oRun);
                        oLastUsedPara.Set_Align(AscCommon.align_Center);
                    }
                    else {
                        oRun.AddToContentToEnd(AscPDF.codePointToRunElement(nCharCode));
                    }
                }
            }

            let _t = this;
            if (Asc.editor.getDocumentRenderer().IsOpenAnnotsInProgress) {
                new Promise(function(resolve) {
                    AscFonts.FontPickerByCharacter.checkTextLight(aRCInfo.reduce(function(accumulator, rc) {
                        return accumulator + rc["text"];
                    }, ""), _t, resolve);
                }).then(function() {
                    _t.SetNeedRecalc(true);
                })
            }
            else {
                _t.SetNeedRecalc(true);
            }
        }
        else {
            this.setTxBody(null);
            this.setTxBox(true);
        }
    };
    CAnnotationLine.prototype.SetRichContents = function(aRCInfo) {
        AscCommon.History.Add(new CChangesPDFAnnotRC(this, this.GetRichContents(), aRCInfo));
        this._richContents = aRCInfo;

        this.SetWasChanged(true);
        this.private_UpdateCaption();
        this.SetNeedRecalc(true);
    };
    CAnnotationLine.prototype.GetRichContents = function(bCalced) {
        if (!bCalced)
            return this._richContents;

        let oContent = this.GetDocContent();
        if (!oContent) {
            return null;
        }
        
        let aRCInfo = [];

        for (let i = 0, nCount = oContent.GetElementsCount(); i < nCount; i++) {
            let oPara = oContent.GetElement(i);

            for (let j = 0, nRunsCount = oPara.GetElementsCount(); j < nRunsCount; j++) {
                let oRun = oPara.GetElement(j);
                let sText = oRun.GetText();
                let oUniColor   = oRun.Pr.Unifill;
                let oRGBA       = oUniColor ? oUniColor.fill.color.color.RGBA : null;
                let aPdfColor   = oRGBA ? [oRGBA.R / 255, oRGBA.G / 255, oRGBA.B / 255] : [0, 0, 0];

                let sFont   = oRun.Get_RFonts().Ascii.Name;
                let isEmbed = false;
                let prefix  = AscFonts.getEmbeddedFontPrefix();

                let nVertAlign;
                switch (oRun.GetVertAlign()) {
                    case AscCommon.vertalign_SuperScript:
                        nVertAlign = 0;
                        break;
                    case AscCommon.vertalign_SubScript:
                        nVertAlign = -0.01;
                        break;
                }

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
                    "text":             sText,
                    "vertical":         nVertAlign
                };

                if (isEmbed) {
                    oRCInfo["name"] = sFont;
                }
                else {
                    oRCInfo["actual"] = sFont;
                }

                aRCInfo.push(oRCInfo);
            }

            if (aRCInfo[aRCInfo.length - 1])
                aRCInfo[aRCInfo.length - 1]["text"] += '\r';
        }

        return aRCInfo;
    };
    CAnnotationLine.prototype.GetContents = function() {
        if (!this.IsDoCaption()) {
            return AscPDF.CAnnotationBase.prototype.GetContents.call(this);
        }

        let oContent = this.GetDocContent();
        return oContent.GetText({ParaSeparator: ''}) || AscPDF.CAnnotationBase.prototype.GetContents.call(this);
    };
    CAnnotationLine.prototype.GetDocContent = function() {
        return this.getDocContent();
    };
	CAnnotationLine.prototype.EnterText = function(value) {
		let docContent = this.GetDocContent();
        let codePoints = typeof(value) === "string" ? value.codePointsArray() : value;

        function correctCodePoints(codePoints) {
            function correctCodePoint(codePoint) {
                if ([9, 10, 13].includes(codePoint)) {
                    return 32;
                }
                else {
                    return codePoint;
                }
            }

            if (Array.isArray(codePoints)) {
                return codePoints.map(correctCodePoint);
            }

            return correctCodePoint(codePoints);
        };

		let result = docContent.EnterText(correctCodePoints(codePoints));

        this.private_UpdateRect();
        this.SetWasChanged(true);
        docContent.RecalculateCurPos();

		return result;
	};
    CAnnotationLine.prototype.SetLineStart = function(nType) {
        if (this._lineStart == nType) {
            return;
        }
        
        AscCommon.History.Add(new CChangesPDFAnnotLineStart(this, this._lineStart, nType));

        this._lineStart = nType;

        this.SetWasChanged(true);
        this.private_UpdateLineStart();
    };
    CAnnotationLine.prototype.private_UpdateLineStart = function() {
        let oLine = this.spPr.ln;
        oLine.setHeadEnd(new AscFormat.EndArrow());
        let nLineEndType = AscPDF.getInnerLineEndType(this.GetLineStart());

        oLine.headEnd.setType(nLineEndType);
        oLine.headEnd.setLen(AscFormat.LineEndSize.Mid);
        this.handleUpdateLn();
    };
    CAnnotationLine.prototype.SetLineEnd = function(nType) {
        if (this._lineEnd == nType) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFAnnotLineEnd(this, this._lineEnd, nType));

        this._lineEnd = nType;

        this.SetWasChanged(true);
        this.private_UpdateLineEnd();
    };
    CAnnotationLine.prototype.private_UpdateLineEnd = function() {
        let oLine = this.spPr.ln;
        oLine.setTailEnd(new AscFormat.EndArrow());
        let nLineEndType = AscPDF.getInnerLineEndType(this.GetLineEnd());

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

    function getFigureSize(nType, nLineW) {
        let oSize = {width: 0, height: 0};

        switch (nType) {
            case AscPDF.LINE_END_TYPE.None:
                oSize.width = nLineW;
                oSize.height = nLineW;
            case AscPDF.LINE_END_TYPE.OpenArrow:
            case AscPDF.LINE_END_TYPE.ClosedArrow:
                oSize.width = 6 * nLineW;
                oSize.height = 6 * nLineW;
                break;
            case AscPDF.LINE_END_TYPE.Diamond:
            case AscPDF.LINE_END_TYPE.Square:
                oSize.width = 6 * nLineW;
                oSize.height = 6 * nLineW;
                break;
            case AscPDF.LINE_END_TYPE.Circle:
                oSize.width = 6 * nLineW;
                oSize.height = 6 * nLineW;
                break;
            case AscPDF.LINE_END_TYPE.RClosedArrow:
                oSize.width = 8 * nLineW;
                oSize.height = 8 * nLineW;
                break;
            case AscPDF.LINE_END_TYPE.ROpenArrow:
                oSize.width = 8 * nLineW;
                oSize.height = 8 * nLineW;
                break;
            case AscPDF.LINE_END_TYPE.Butt:
                oSize.width = 6 * nLineW;
                oSize.height = 6 * nLineW;
                break;
            case AscPDF.LINE_END_TYPE.Slash:
                oSize.width = 6 * nLineW;
                oSize.height = 6 * nLineW;
                break;
        }

        return oSize;
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

    window["AscPDF"].CAnnotationLine    = CAnnotationLine;
    window["AscPDF"].LINE_END_TYPE      = LINE_END_TYPE;
    window["AscPDF"].getInnerLineEndType= getInnerLineEndType;
})();

