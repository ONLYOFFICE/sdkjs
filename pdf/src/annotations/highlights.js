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
	 * Class representing a base highlight annotation.
	 * @constructor
    */
    function CAnnotationTextMarkup(sName, nType, aRect, oDoc)
    {
        AscPDF.CAnnotationBase.call(this, sName, nType, aRect, oDoc);

        this._quads         = [];
        this._richContents  = undefined;
        this._rotate        = undefined;
        this._width         = undefined;
    }
    CAnnotationTextMarkup.prototype = Object.create(AscPDF.CAnnotationBase.prototype);
	CAnnotationTextMarkup.prototype.constructor = CAnnotationTextMarkup;

    CAnnotationTextMarkup.prototype.select = AscFormat.CGraphicObjectBase.prototype.select;
    CAnnotationTextMarkup.prototype.deselect = AscFormat.CGraphicObjectBase.prototype.deselect;
    CAnnotationTextMarkup.prototype.canChangeAdjustments = function() {};
    CAnnotationTextMarkup.prototype.hitToHandles = function() {};
    CAnnotationTextMarkup.prototype.hitInBoundingRect = function() {};
    CAnnotationTextMarkup.prototype.getNoChangeAspect = function() {};
    CAnnotationTextMarkup.prototype.getMainGroup = function() {};
    CAnnotationTextMarkup.prototype.getObjectName = function() {};
    CAnnotationTextMarkup.prototype.isShape = function() {};
    CAnnotationTextMarkup.prototype.isImage = function() {};
    CAnnotationTextMarkup.prototype.createMoveTrack = function() {};
    CAnnotationTextMarkup.prototype.canMove = function() {
        return false;
    };
    CAnnotationTextMarkup.prototype.canResize = function() {
        return false;
    };
    CAnnotationTextMarkup.prototype.canRotate = function() {
        return false;
    };

    CAnnotationTextMarkup.prototype.IsTextMarkup = function() {
        return true;
    };

    CAnnotationTextMarkup.prototype.SetQuads = function(aFullQuads) {
        let oThis = this;
        aFullQuads.forEach(function(aQuads) {
            oThis.AddQuads(aQuads);
        });
    };
    CAnnotationTextMarkup.prototype.GetQuads = function() {
        return this._quads;
    };
    CAnnotationTextMarkup.prototype.AddQuads = function(aQuads) {
        AscCommon.History.Add(new CChangesPDFAnnotQuads(this, this._quads.length, aQuads, true));
        this._quads.push(aQuads);
        this.SetNeedRecalc(true);
    };
    
    CAnnotationTextMarkup.prototype.AddToRedraw = function() {
        let oViewer = editor.getDocumentRenderer();
        let nPage   = this.GetPage();
        
        function setRedrawPageOnRepaint() {
            if (oViewer.pagesInfo.pages[nPage]) {
                oViewer.pagesInfo.pages[nPage].needRedrawMarkups = true;
                oViewer.thumbnails && oViewer.thumbnails._repaintPage(nPage);
            }
        }

        oViewer.paint(setRedrawPageOnRepaint);
    };
    CAnnotationTextMarkup.prototype.IsInQuads = function(x, y) {
        return IsInQuads(this.GetQuads(), x, y);
    };
    CAnnotationTextMarkup.prototype.DrawSelected = function(overlay) {
        overlay.m_oContext.lineWidth    = 3;
        overlay.m_oContext.globalAlpha  = 1;
        overlay.m_oContext.strokeStyle  = "rgb(33, 117, 200)";
        overlay.m_oContext.beginPath();

        fillRegion(this.GetUnitedRegion(), overlay, this.GetPage());
        overlay.m_oContext.stroke();
    };
    CAnnotationTextMarkup.prototype.GetUnitedRegion = function() {
        if (this.unitedRegion)
            return this.unitedRegion;

        let aQuads  = this.GetQuads();
        let fUniter = AscGeometry.PolyBool.union;
        let resultRegion;

        //let time1 = performance.now();
        let aAllRegions = [];
        for (let i = 0; i < aQuads.length; i++) {
            let aPoints = aQuads[i];

            aAllRegions.push({
                inverted : false,
                regions : [
                    [
                        [aPoints[0], aPoints[1]],
                        [aPoints[2], aPoints[3]],
                        [aPoints[6], aPoints[7]],
                        [aPoints[4], aPoints[5]]
                    ]
                ]
            });
        }

        if (aAllRegions.length > 1) {
            resultRegion = fUniter(aAllRegions[0], aAllRegions[1]);
            for (let i = 2; i < aAllRegions.length; i++) {
                resultRegion = fUniter(resultRegion, aAllRegions[i]);
            }
        }
        else {
            resultRegion = aAllRegions[0];
        }
        //let time2 = performance.now();
        //console.log("union: " + (time2 - time1));

        this.unitedRegion = resultRegion;
        return this.unitedRegion;
    };
    CAnnotationTextMarkup.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        // quads
        let aQuads = this.GetQuads();
        let nLen = 0;
        for (let i = 0; i < aQuads.length; i++) {
            nLen += aQuads[i].length;
        }
        memory.WriteLong(nLen);  
        for (let i = 0; i < aQuads.length; i++) {
            for (let j = 0; j < aQuads[i].length; j++) {
                memory.WriteDouble(aQuads[i][j]);
            }
        }

        let nEndPos = memory.GetCurPosition();
        memory.Seek(memory.posForFlags);
        memory.WriteLong(memory.annotFlags);
        
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };
    /**
	 * Class representing a highlight annotation.
	 * @constructor
    */
    function CAnnotationHighlight(sName, nPage, aRect, oDoc)
    {
        CAnnotationTextMarkup.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Highlight, nPage, aRect, oDoc);
    }
    CAnnotationHighlight.prototype.constructor = CAnnotationHighlight;
    AscFormat.InitClass(CAnnotationHighlight, CAnnotationTextMarkup, AscDFH.historyitem_type_Pdf_Annot_Highlight);

    CAnnotationHighlight.prototype.IsHighlight = function() {
        return true;
    };

    CAnnotationHighlight.prototype.Draw = function(oGraphicsPDF) {
        if (this.IsHidden() == true)
            return;

        let oRGBFill = this.GetRGBColor(this.GetStrokeColor());

        let aQuads = this.GetQuads();
        for (let i = 0; i < aQuads.length; i++) {
            let aPoints = aQuads[i];

            let oPoint1 = {
                x: aPoints[0],
                y: aPoints[1]
            }
            let oPoint2 = {
                x: aPoints[2],
                y: aPoints[3]
            }
            let oPoint3 = {
                x: aPoints[4],
                y: aPoints[5]
            }
            let oPoint4 = {
                x: aPoints[6],
                y: aPoints[7]
            }

            let dx = oPoint2.x - oPoint1.x;
            let dy = oPoint2.y - oPoint1.y;
            let angle1          = Math.atan2(dy, dx);
            let rotationAngle   = angle1;

            oGraphicsPDF.SetGlobalAlpha(this.GetOpacity());
            AscPDF.startMultiplyMode(oGraphicsPDF.GetContext());

            oGraphicsPDF.BeginPath();
            oGraphicsPDF.SetFillStyle(oRGBFill.r, oRGBFill.g, oRGBFill.b);

            if (rotationAngle == 0 || rotationAngle == 3/2 * Math.PI) {
                let aMinRect = getMinRect(aPoints);

                oGraphicsPDF.SetIntegerGrid(true);
                oGraphicsPDF.Rect(aMinRect[0], aMinRect[1], aMinRect[2] - aMinRect[0], aMinRect[3] - aMinRect[1], true);
                oGraphicsPDF.SetIntegerGrid(false);
            }
            else {
                oGraphicsPDF.MoveTo(oPoint1.x, oPoint1.y);
                oGraphicsPDF.LineTo(oPoint2.x, oPoint2.y);
                oGraphicsPDF.LineTo(oPoint4.x, oPoint4.y);
                oGraphicsPDF.LineTo(oPoint3.x, oPoint3.y);
                oGraphicsPDF.ClosePath();
            }

            oGraphicsPDF.Fill();
            AscPDF.endMultiplyMode(oGraphicsPDF.GetContext());
        }

        let aUnitedRegion = this.GetUnitedRegion();
        oGraphicsPDF.DrawLockObjectRect(this.Lock.Get_Type(), aUnitedRegion.regions);
    };
        
    /**
	 * Class representing a highlight annotation.
	 * @constructor
    */
    function CAnnotationUnderline(sName, nPage, aRect, oDoc)
    {
        CAnnotationTextMarkup.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Underline, nPage, aRect, oDoc);
    }
    CAnnotationUnderline.prototype.constructor = CAnnotationUnderline;
    AscFormat.InitClass(CAnnotationUnderline, CAnnotationTextMarkup, AscDFH.historyitem_type_Pdf_Annot_Underline);

    CAnnotationUnderline.prototype.Draw = function(oGraphicsPDF) {
        if (this.IsHidden() == true)
            return;

        let aQuads      = this.GetQuads();
        let oRGBFill    = this.GetRGBColor(this.GetStrokeColor());
        
        for (let i = 0; i < aQuads.length; i++) {
            let aPoints     = aQuads[i];
            
            oGraphicsPDF.SetGlobalAlpha(this.GetOpacity());
            oGraphicsPDF.SetStrokeStyle(oRGBFill.r, oRGBFill.g, oRGBFill.b);
            oGraphicsPDF.BeginPath();

            let oPoint1 = {
                x: aPoints[0],
                y: aPoints[1]
            }
            let oPoint2 = {
                x: aPoints[2],
                y: aPoints[3]
            }
            let oPoint3 = {
                x: aPoints[4],
                y: aPoints[5]
            }
            let oPoint4 = {
                x: aPoints[6],
                y: aPoints[7]
            }

            let X1 = oPoint3.x
            let Y1 = oPoint3.y;
            let X2 = oPoint4.x;
            let Y2 = oPoint4.y;

            let dx1 = oPoint2.x - oPoint1.x;
            let dy1 = oPoint2.y - oPoint1.y;
            let dx2 = oPoint4.x - oPoint3.x;
            let dy2 = oPoint4.y - oPoint3.y;
            let angle1          = Math.atan2(dy1, dx1);
            let angle2          = Math.atan2(dy2, dx2);
            let rotationAngle   = angle1;

            let nSide;
            if (rotationAngle == 0 || rotationAngle == 3/2 * Math.PI) {
                nSide = Math.abs(oPoint3.y - oPoint1.y);
                oGraphicsPDF.SetLineWidth(Math.max(1, nSide * 0.1 >> 0));
            }
            else {
                nSide = findMaxSideWithRotation(oPoint1.x, oPoint1.y, oPoint2.x, oPoint2.y, oPoint3.x, oPoint3.y, oPoint4.x, oPoint4.y);
                oGraphicsPDF.SetLineWidth(Math.max(1, nSide * 0.1 >> 0));
            }

            let nLineW      = oGraphicsPDF.GetLineWidth();
            let nIndentX    = Math.sin(rotationAngle) * nLineW * 1.5;
            let nIndentY    = Math.cos(rotationAngle) * nLineW * 1.5;

            if (rotationAngle == 0 || rotationAngle == 3/2 * Math.PI) {
                oGraphicsPDF.HorLine(X1, X2, Y2 - nIndentY);
            }
            else {
                oGraphicsPDF.MoveTo(X1 + nIndentX, Y1 - nIndentY);
                oGraphicsPDF.LineTo(X2 + nIndentX, Y2 - nIndentY);
            }
            
            oGraphicsPDF.Stroke();
        }

        let aUnitedRegion = this.GetUnitedRegion();
        oGraphicsPDF.DrawLockObjectRect(this.Lock.Get_Type(), aUnitedRegion.regions);
    };
    
    /**
	 * Class representing a highlight annotation.
	 * @constructor
    */
    function CAnnotationStrikeout(sName, nPage, aRect, oDoc)
    {
        CAnnotationTextMarkup.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Strikeout, nPage, aRect, oDoc);
    }
    CAnnotationStrikeout.prototype.constructor = CAnnotationStrikeout;
    AscFormat.InitClass(CAnnotationStrikeout, CAnnotationTextMarkup, AscDFH.historyitem_type_Pdf_Annot_Strikeout);

    CAnnotationStrikeout.prototype.Draw = function(oGraphicsPDF) {
        if (this.IsHidden() == true)
            return;

        let aQuads      = this.GetQuads();
        let oRGBFill    = this.GetRGBColor(this.GetStrokeColor());

        for (let i = 0; i < aQuads.length; i++) {
            let aPoints = aQuads[i];

            oGraphicsPDF.BeginPath();

            oGraphicsPDF.SetGlobalAlpha(this.GetOpacity());
            oGraphicsPDF.SetStrokeStyle(oRGBFill.r, oRGBFill.g, oRGBFill.b);

            let oPoint1 = {
                x: aPoints[0],
                y: aPoints[1]
            }
            let oPoint2 = {
                x: aPoints[2],
                y: aPoints[3]
            }
            let oPoint3 = {
                x: aPoints[4],
                y: aPoints[5]
            }
            let oPoint4 = {
                x: aPoints[6],
                y: aPoints[7]
            }

            let dx1 = oPoint2.x - oPoint1.x;
            let dy1 = oPoint2.y - oPoint1.y;
            let dx2 = oPoint4.x - oPoint3.x;
            let dy2 = oPoint4.y - oPoint3.y;
            let angle1          = Math.atan2(dy1, dx1);
            let angle2          = Math.atan2(dy2, dx2);
            let rotationAngle   = angle1;

            let X1 = oPoint1.x + (oPoint3.x - oPoint1.x) / 2;
            let Y1 = oPoint1.y + (oPoint3.y - oPoint1.y) / 2;
            let X2 = oPoint2.x + (oPoint4.x - oPoint2.x) / 2;
            let Y2 = oPoint2.y + (oPoint4.y - oPoint2.y) / 2;

            let nSide;
            if (rotationAngle == 0 || rotationAngle == 3/2 * Math.PI) {
                nSide = Math.abs(oPoint3.y - oPoint1.y);
                oGraphicsPDF.SetLineWidth(Math.max(1, nSide * 0.1 >> 0));
                oGraphicsPDF.HorLine(X1, X2, Y2);
            }
            else {
                nSide = findMaxSideWithRotation(oPoint1.x, oPoint1.y, oPoint2.x, oPoint2.y, oPoint3.x, oPoint3.y, oPoint4.x, oPoint4.y);
                oGraphicsPDF.SetLineWidth(Math.max(1, nSide * 0.1 >> 0));

                oGraphicsPDF.MoveTo(X1, Y1);
                oGraphicsPDF.LineTo(X2, Y2);
            }
            
            oGraphicsPDF.Stroke();
        }

        let aUnitedRegion = this.GetUnitedRegion();
        oGraphicsPDF.DrawLockObjectRect(this.Lock.Get_Type(), aUnitedRegion.regions);
    };

    /**
	 * Class representing a squiggly annotation.
	 * @constructor
    */
    function CAnnotationSquiggly(sName, nPage, aRect, oDoc)
    {
        CAnnotationTextMarkup.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Squiggly, nPage, aRect, oDoc);
    }
    CAnnotationSquiggly.prototype.constructor = CAnnotationSquiggly;
    AscFormat.InitClass(CAnnotationSquiggly, CAnnotationTextMarkup, AscDFH.historyitem_type_Pdf_Annot_Squiggly);

    CAnnotationSquiggly.prototype.Draw = function(oGraphicsPDF) {
        if (this.IsHidden())
            return;
    
        let aQuads   = this.GetQuads();
        let oRGBFill = this.GetRGBColor(this.GetStrokeColor());
    
        for (let i = 0; i < aQuads.length; i++) {
            let aPoints = aQuads[i];
            
            oGraphicsPDF.SetGlobalAlpha(this.GetOpacity());
            oGraphicsPDF.SetStrokeStyle(oRGBFill.r, oRGBFill.g, oRGBFill.b);
            
            let oPoint1 = { x: aPoints[0], y: aPoints[1] };
            let oPoint2 = { x: aPoints[2], y: aPoints[3] };
            let oPoint3 = { x: aPoints[4], y: aPoints[5] };
            let oPoint4 = { x: aPoints[6], y: aPoints[7] };
    
            let X1 = oPoint3.x, Y1 = oPoint3.y;
            let X2 = oPoint4.x, Y2 = oPoint4.y;
    
            let dx1 = oPoint2.x - oPoint1.x;
            let dy1 = oPoint2.y - oPoint1.y;
            let angle1 = Math.atan2(dy1, dx1);
            
            let nSide = (angle1 == 0 || angle1 == 3/2 * Math.PI) 
                ? Math.abs(oPoint3.y - oPoint1.y) 
                : findMaxSideWithRotation(oPoint1.x, oPoint1.y, oPoint2.x, oPoint2.y, oPoint3.x, oPoint3.y, oPoint4.x, oPoint4.y);
    
            oGraphicsPDF.SetLineWidth(Math.max(1, nSide * 0.05 >> 0));
            let nLineW   = oGraphicsPDF.GetLineWidth();
            let nIndentX = Math.sin(angle1) * nLineW * 1.5;
            let nIndentY = Math.cos(angle1) * nLineW * 1.5;
    
            let startX = (angle1 == 0 || angle1 == 3/2 * Math.PI) ? X1 : X1 + nIndentX;
            let startY = (angle1 == 0 || angle1 == 3/2 * Math.PI) ? Y1 : Y1 - nIndentY;
            let endX   = (angle1 == 0 || angle1 == 3/2 * Math.PI) ? X2 : X2 + nIndentX;
            let endY   = (angle1 == 0 || angle1 == 3/2 * Math.PI) ? Y2 : Y2 - nIndentY;
    
            drawZigZagLine(oGraphicsPDF, startX, startY, endX, endY, nLineW);
        }
    
        let aUnitedRegion = this.GetUnitedRegion();
        oGraphicsPDF.DrawLockObjectRect(this.Lock.Get_Type(), aUnitedRegion.regions);
    };
    
    function drawZigZagLine(oGraphicsPDF, X1, Y1, X2, Y2, nLineW) {
        let length = Math.sqrt(Math.pow(X2 - X1, 2) + Math.pow(Y2 - Y1, 2));
        // Параметры волны
        let wavelength = 2;          // длина одного "зубчика"
        let amplitude = nLineW * 1;  // высота волны
        let dx = (X2 - X1) / length;
        let dy = (Y2 - Y1) / length;
        let nx = -dy;
        let ny = dx;
    
        // Сколько сегментов поместится на всей длине?
        // Один период (полный зубчик - вверх-вниз) занимает 2 сегмента по wavelength/2 каждый,
        // но для простоты возьмём wavelength как полный период.
        let segments = Math.floor(length / wavelength);
    
        oGraphicsPDF.BeginPath();
        oGraphicsPDF.MoveTo(X1, Y1);
    
        for (let i = 1; i <= segments; i++) {
            // Чередуем направление сдвига: вверх-амплитуда, вниз-амплитуда
            let isUp = (i % 2 === 1); 
            let dist = i * wavelength;
    
            let offset = isUp ? amplitude : -amplitude;
    
            oGraphicsPDF.LineTo(
                X1 + dx * dist + nx * offset,
                Y1 + dy * dist + ny * offset
            );
        }
    
        // Если длина не делится ровно на сегменты, дойдём до конца
        let remainder = length - segments * wavelength;
        if (remainder > 0) {
            let lastDist = length;
            // Последний сегмент: продолжаем паттерн
            let isUp = (segments % 2 === 0); 
            let offset = isUp ? amplitude : -amplitude;
            // Пропорционально оставшейся длине уменьшим offset, чтобы плавно закончить
            let ratio = remainder / wavelength;
            oGraphicsPDF.LineTo(
                X1 + dx * lastDist + nx * (offset * ratio),
                Y1 + dy * lastDist + ny * (offset * ratio)
            );
        }
    
        oGraphicsPDF.Stroke();
    }

    let CARET_SYMBOL = {
        None:       0,
        Paragraph:  1,
        Space:      2
    }

    /**
	 * Class representing a caret annotation.
	 * @constructor
    */
    function CAnnotationCaret(sName, nPage, aRect, oDoc)
    {
        CAnnotationTextMarkup.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Caret, nPage, aRect, oDoc);
        this._caretSymbol = CARET_SYMBOL.None;
    }
    CAnnotationCaret.prototype.constructor = CAnnotationCaret;
    AscFormat.InitClass(CAnnotationCaret, CAnnotationTextMarkup, AscDFH.historyitem_type_Pdf_Annot_Caret);

    CAnnotationCaret.prototype.Draw = function(oGraphicsPDF) {
        if (this.IsHidden() == true)
            return;

        let aQuads      = this.GetQuads();
        let oRGBFill    = this.GetRGBColor(this.GetStrokeColor());
        
        for (let i = 0; i < aQuads.length; i++) {
            let aPoints     = aQuads[i];
            
            oGraphicsPDF.SetGlobalAlpha(this.GetOpacity());
            oGraphicsPDF.SetStrokeStyle(oRGBFill.r, oRGBFill.g, oRGBFill.b);
            oGraphicsPDF.BeginPath();

            let oPoint1 = {
                x: aPoints[0],
                y: aPoints[1]
            }
            let oPoint2 = {
                x: aPoints[2],
                y: aPoints[3]
            }
            let oPoint3 = {
                x: aPoints[4],
                y: aPoints[5]
            }
            let oPoint4 = {
                x: aPoints[6],
                y: aPoints[7]
            }

            let X1 = oPoint3.x
            let Y1 = oPoint3.y;
            let X2 = oPoint4.x;
            let Y2 = oPoint4.y;

            let dx1 = oPoint2.x - oPoint1.x;
            let dy1 = oPoint2.y - oPoint1.y;
            let dx2 = oPoint4.x - oPoint3.x;
            let dy2 = oPoint4.y - oPoint3.y;
            let angle1          = Math.atan2(dy1, dx1);
            let angle2          = Math.atan2(dy2, dx2);
            let rotationAngle   = angle1;

            let nSide;
            if (rotationAngle == 0 || rotationAngle == 3/2 * Math.PI) {
                nSide = Math.abs(oPoint3.y - oPoint1.y);
                oGraphicsPDF.SetLineWidth(Math.max(1, nSide * 0.1 >> 0));
            }
            else {
                nSide = findMaxSideWithRotation(oPoint1.x, oPoint1.y, oPoint2.x, oPoint2.y, oPoint3.x, oPoint3.y, oPoint4.x, oPoint4.y);
                oGraphicsPDF.SetLineWidth(Math.max(1, nSide * 0.1 >> 0));
            }

            let nLineW      = oGraphicsPDF.GetLineWidth();
            let nIndentX    = Math.sin(rotationAngle) * nLineW * 1.5;
            let nIndentY    = Math.cos(rotationAngle) * nLineW * 1.5;

            if (rotationAngle == 0 || rotationAngle == 3/2 * Math.PI) {
                oGraphicsPDF.HorLine(X1, X2, Y2 - nIndentY);
            }
            else {
                oGraphicsPDF.MoveTo(X1 + nIndentX, Y1 - nIndentY);
                oGraphicsPDF.LineTo(X2 + nIndentX, Y2 - nIndentY);
            }
            
            oGraphicsPDF.Stroke();
        }

        let aUnitedRegion = this.GetUnitedRegion();
        oGraphicsPDF.DrawLockObjectRect(this.Lock.Get_Type(), aUnitedRegion.regions);
    };
    CAnnotationCaret.prototype.SetCaretSymbol = function(nType) {
        this._caretSymbol = nType;
    };
    CAnnotationCaret.prototype.GetCaretSymbol = function() {
        return this._caretSymbol;
    };
    CAnnotationCaret.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        // rectange diff
        let aRD = this.GetRectangleDiff();
        if (aRD) {
            memory.annotFlags |= (1 << 15);
            for (let i = 0; i < 4; i++) {
                memory.WriteDouble(aRD[i]);
            }
        }
        
        // caret symbol
        let nCaretSymbol = this.GetCaretSymbol();
        if (nCaretSymbol != null) {
            memory.annotFlags |= (1 << 16);
            memory.WriteByte(nCaretSymbol);
        }

        let nEndPos = memory.GetCurPosition();
        memory.Seek(memory.posForFlags);
        memory.WriteLong(memory.annotFlags);
        
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };

    function findMaxSideWithRotation(x1, y1, x2, y2, x3, y3, x4, y4) {
        // Найдите центр поворота
        const x_center = (x1 + x3) / 2;
        const y_center = (y1 + y3) / 2;
      
        // Найдите угол поворота
        const angle = Math.atan2(y3 - y1, x3 - x1);
      
        // Выполните поворот вершин прямоугольника на обратный угол
        const cosAngle = Math.cos(-angle);
        const sinAngle = Math.sin(-angle);
      
        const x1_rotated = cosAngle * (x1 - x_center) - sinAngle * (y1 - y_center) + x_center;
        const y1_rotated = sinAngle * (x1 - x_center) + cosAngle * (y1 - y_center) + y_center;
      
        const x2_rotated = cosAngle * (x2 - x_center) - sinAngle * (y2 - y_center) + x_center;
        const y2_rotated = sinAngle * (x2 - x_center) + cosAngle * (y2 - y_center) + y_center;
      
        const x3_rotated = cosAngle * (x3 - x_center) - sinAngle * (y3 - y_center) + x_center;
        const y3_rotated = sinAngle * (x3 - x_center) + cosAngle * (y3 - y_center) + y_center;
      
        const x4_rotated = cosAngle * (x4 - x_center) - sinAngle * (y4 - y_center) + x_center;
        const y4_rotated = sinAngle * (x4 - x_center) + cosAngle * (y4 - y_center) + y_center;
      
        // Найдите длины сторон
        const sideAB = Math.sqrt(Math.pow(x2_rotated - x1_rotated, 2) + Math.pow(y2_rotated - y1_rotated, 2));
        const sideBC = Math.sqrt(Math.pow(x3_rotated - x2_rotated, 2) + Math.pow(y3_rotated - y2_rotated, 2));
        const sideCD = Math.sqrt(Math.pow(x4_rotated - x3_rotated, 2) + Math.pow(y4_rotated - y3_rotated, 2));
        const sideDA = Math.sqrt(Math.pow(x1_rotated - x4_rotated, 2) + Math.pow(y1_rotated - y4_rotated, 2));
      
        // Найдите максимальную сторону
        const maxSide = Math.max(sideAB, sideBC, sideCD, sideDA);
      
        return maxSide;
    }

    function fillRegion(polygon, overlay, pageIndex)
    {
        let oCtx    = overlay.m_oContext;
        let oViewer = Asc.editor.getDocumentRenderer();
        let nScale  = oViewer.zoom * oViewer.getDrawingPageScale(pageIndex) * AscCommon.AscBrowser.retinaPixelRatio;

        let xCenter = oViewer.width >> 1;
        if (oViewer.documentWidth > oViewer.width)
		{
			xCenter = (oViewer.documentWidth >> 1) - (oViewer.scrollX) >> 0;
		}
		let yPos    = oViewer.scrollY >> 0;
        let page    = oViewer.drawingPages[pageIndex];
        let w       = (page.W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
        let h       = (page.H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
        let indLeft = ((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - (w >> 1);
        let indTop  = ((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

        if (true == oViewer.isLandscapePage(pageIndex))
            indLeft = indLeft + (w - h) / 2;

        // рисуем всегда в пиксельной сетке. при наклонных линиях - +- 1 пиксел - ничего страшного
        let pointOffset = (oCtx.lineWidth & 1) ? 0.5 : 0;
        
        for (let i = 0, countPolygons = polygon.regions.length; i < countPolygons; i++)
        {
            let region = polygon.regions[i];
            let countPoints = region.length;

            if (2 > countPoints)
                continue;

            let X = indLeft + region[0][0] * nScale;
            let Y = indTop + region[0][1] * nScale;

            oCtx.moveTo((X >> 0) + pointOffset, (Y >> 0) + pointOffset);

            overlay.CheckPoint1(X, Y);
            overlay.CheckPoint2(X, Y);

            for (let j = 1, countPoints = region.length; j < countPoints; j++)
            {
                X = indLeft + region[j][0] * nScale;
                Y = indTop + region[j][1] * nScale;;

                oCtx.lineTo((X >> 0) + pointOffset, (Y >> 0) + pointOffset);
                overlay.CheckPoint1(X, Y);
                overlay.CheckPoint2(X, Y);
            }

            oCtx.closePath();
        }
    }
    function getMinRect(aPoints) {
        let xMax = aPoints[0], yMax = aPoints[1], xMin = xMax, yMin = yMax;
        for(let i = 1; i < aPoints.length; i++) {
            if (i % 2 == 0) {
                if(aPoints[i] < xMin)
                {
                    xMin = aPoints[i];
                }
                if(aPoints[i] > xMax)
                {
                    xMax = aPoints[i];
                }
            }
            else {
                if(aPoints[i] < yMin)
                {
                    yMin = aPoints[i];
                }

                if(aPoints[i] > yMax)
                {
                    yMax = aPoints[i];
                }
            }
        }

        return [xMin, yMin, xMax, yMax];
    }

    function IsInQuads(aQuads, x, y) {
        let oCtx = Asc.editor.getDocumentRenderer().overlay.m_oContext;
        oCtx.save();
        oCtx.setTransform(1, 0, 0, 1, 0, 0);

        let isInQuads = false; 
        for (let i = 0; i < aQuads.length; i++) {
            let aPoints = aQuads[i];

            let oPoint1 = {
                x: aPoints[0],
                y: aPoints[1]
            }
            let oPoint2 = {
                x: aPoints[2],
                y: aPoints[3]
            }

            let oPoint3 = {
                x: aPoints[4],
                y: aPoints[5]
            }
            let oPoint4 = {
                x: aPoints[6],
                y: aPoints[7]
            }

            let X1 = oPoint1.x;
            let Y1 = oPoint1.y;
            let X2 = oPoint2.x;
            let Y2 = oPoint2.y;
            let X3 = oPoint3.x;
            let Y3 = oPoint3.y;
            let X4 = oPoint4.x;
            let Y4 = oPoint4.y;

            oCtx.beginPath();
            oCtx.moveTo(X1, Y1);
            oCtx.lineTo(X2, Y2);
            oCtx.lineTo(X4, Y4);
            oCtx.lineTo(X3, Y3);
            oCtx.closePath();

            if (oCtx.isPointInPath(x, y))
                isInQuads = true;
        }

        oCtx.restore();
        return isInQuads;
    }

    window["AscPDF"].CAnnotationTextMarkup  = CAnnotationTextMarkup;
    window["AscPDF"].CAnnotationHighlight   = CAnnotationHighlight;
    window["AscPDF"].CAnnotationUnderline   = CAnnotationUnderline;
    window["AscPDF"].CAnnotationStrikeout   = CAnnotationStrikeout;
    window["AscPDF"].CAnnotationSquiggly    = CAnnotationSquiggly;
    window["AscPDF"].CAnnotationCaret       = CAnnotationCaret;
    window["AscPDF"].fillRegion             = fillRegion;
    window["AscPDF"].IsInQuads              = IsInQuads;
    
})();

