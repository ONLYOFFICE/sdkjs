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

    let STAMP_TYPES = {
        // dinamic
        D_Approved:               'D_Approved',
        D_Revised:                'D_Revised',
        D_Reviewed:               'D_Reviewed',
        D_Received:               'D_Received',
        
        // standard
        SB_Approved:              'SB_Approved',
        SB_NotApproved:           'SB_NotApproved',
        SB_Revised:               'SB_Revised',
        SB_Confidential:          'SB_Confidential',
        SB_ForComment:            'SB_ForComment',
        SB_ForPublicRelease:      'SB_ForPublicRelease',
        SB_NotForPublicRelease:   'SB_NotForPublicRelease',
        SB_PreliminaryResults:    'SB_PreliminaryResults',
        SB_InformationOnly:       'SB_InformationOnly',
        SB_Draft:                 'SB_Draft',
        SB_Completed:             'SB_Completed',
        SB_Final:                 'SB_Final',
        SB_Void:                  'SB_Void',
        
        // sign
        SH_SignHere:              'SH_SignHere',
        SH_Witness:               'SH_Witness',
        SH_InitialHere:           'SH_InitialHere',
        
        // foxit
        Expired:                  'Expired',

        // onlyoffice
        Image:                    'Image'
    }

    /**
	 * Class representing a stamp annotation.
	 * @constructor
    */
    function CAnnotationStamp(sName, aRect, oDoc)
    {
        AscPDF.CPdfShape.call(this);
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Stamp, aRect, oDoc);
        
        this._rotate = 0;
        this._stampType = undefined;

        this.Init();
    }
    
	CAnnotationStamp.prototype.constructor = CAnnotationStamp;
    AscFormat.InitClass(CAnnotationStamp, AscPDF.CPdfShape, AscDFH.historyitem_type_Pdf_Annot_Stamp);
    Object.assign(CAnnotationStamp.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationStamp.prototype.Copy = function(isForMove) {
        let oCopy = AscPDF.CAnnotationBase.prototype.Copy.call(this, isForMove);

        let aInRect = this.GetInRect();

        oCopy.SetRenderStructure(this.GetRenderStructure());
        oCopy.SetInRect(aInRect.slice());
        oCopy.SetIconType(this.GetIconType());
        oCopy.SetRotate(this.GetRotate());

        return oCopy;
    };
    
    CAnnotationStamp.prototype.IsStamp = function() {
        return true;
    };
    CAnnotationStamp.prototype.RecalcSizes = function() {
        let aRect = this.GetRect();
        let aInRect = this.GetInRect();

        if (!aInRect) {
            AscPDF.CAnnotationBase.prototype.RecalcSizes.call(this);
            return;
        }

        function calcScale(rectA, angleRad, bboxB, eps) {
            eps = undefined !== eps ? eps : 1e-3;
            const cosL = Math.cos(angleRad);
            const sinL = Math.sin(angleRad);

            let minX =  Infinity, maxX = -Infinity;
            let minY =  Infinity, maxY = -Infinity;

            for (let i = 0; i < 8; i += 2) {
                const x = rectA[i],     y = rectA[i + 1];
                const uX =  cosL * x - sinL * y;
                const uY =  sinL * x + cosL * y;

                if (uX < minX) minX = uX;
                if (uX > maxX) maxX = uX;
                if (uY < minY) minY = uY;
                if (uY > maxY) maxY = uY;
            }

            const dUX = maxX - minX;
            const dUY = maxY - minY;
            if (dUX < eps || dUY < eps) return 1;

            const wB = Math.abs(bboxB[2] - bboxB[0]);
            const hB = Math.abs(bboxB[3] - bboxB[1]);

            const sX = wB / dUX;
            const sY = hB / dUY;

            return (sX + sY) * 0.5;
        }

        function getDistance(x1, y1, x2, y2) {
            return Math.sqrt(Math.pow(x2 - x1, 2) + Math.pow(y2 - y1, 2));
        }

        function rotateRect(pts, deg) {
            const rad = deg * Math.PI / 180;
            const cos = Math.cos(rad);
            const sin = Math.sin(rad);

            // центр прямоугольника
            let cx = 0, cy = 0;
            for (let i = 0; i < 8; i += 2) {
                cx += pts[i];
                cy += pts[i + 1];
            }
            cx /= 4;
            cy /= 4;

            // поворот каждой точки
            const res = new Array(8);
            for (let i = 0; i < 8; i += 2) {
                const dx = pts[i]     - cx;
                const dy = pts[i + 1] - cy;
                res[i]     = cx + dx * cos - dy * sin;
                res[i + 1] = cy + dx * sin + dy * cos;
            }
            return res;
        }

        function getRectRotation(pts) {
            let dx1 = pts[2] - pts[0];
            let dy1 = pts[3] - pts[1];
            let dx2 = pts[4] - pts[2];
            let dy2 = pts[5] - pts[3];

            let len1 = dx1 * dx1 + dy1 * dy1;
            let len2 = dx2 * dx2 + dy2 * dy2;
            let dx   = len1 >= len2 ? dx1 : dx2;
            let dy   = len1 >= len2 ? dy1 : dy2;

            let deg = Math.atan2(dy, dx) * 180 / Math.PI; // ±180°
            return deg < 0 ? deg + 360 : deg;             // 0‑360°
        }

        // If this annotation is from a file, then the rotation is not taken into account.
        let rad = this.GetRotate() * Math.PI / 180;
        
        let nInRectRot = getRectRotation(aInRect);
        if (nInRectRot != 0) {
            aInRect = rotateRect(aInRect, -nInRectRot);
        }

        let scale = calcScale(aInRect, rad, aRect);
        let nShapeW = getDistance(aInRect[0], aInRect[1], aInRect[6], aInRect[7]);
        let nShapeH = getDistance(aInRect[0], aInRect[1], aInRect[2], aInRect[3]);

        if (Math.abs(scale - 1) < 1e-3) scale = 1;

        let nExtX = nShapeW * scale * g_dKoef_pt_to_mm;
        let nExtY = nShapeH * scale * g_dKoef_pt_to_mm;
        let nOffX = aRect[0] - (nExtX / g_dKoef_pt_to_mm - (aRect[2] - aRect[0])) / 2;
        let nOffY = aRect[1] - (nExtY / g_dKoef_pt_to_mm - (aRect[3] - aRect[1])) / 2;

        AscCommon.History.StartNoHistoryMode();
        this.spPr.xfrm.setOffX(nOffX * g_dKoef_pt_to_mm);
        this.spPr.xfrm.setOffY(nOffY * g_dKoef_pt_to_mm);
        this.spPr.xfrm.setExtX(nExtX);
        this.spPr.xfrm.setExtY(nExtY);
        AscCommon.History.EndNoHistoryMode();

        this.SetNeedRecalcSizes(false);
    };
    CAnnotationStamp.prototype.Draw = function(oGraphicsPDF, oGraphicsWord) {
        let sType = this.GetIconType();
        if (sType == undefined) {
            return;
        }

        this.Recalculate();
        if (AscPDF.STAMP_TYPES.Image == sType) {
            this.draw(oGraphicsWord, oGraphicsPDF);
        }
        else {
            this._draw(oGraphicsWord, oGraphicsPDF);
        }
    };
    CAnnotationStamp.prototype._draw = function(oGraphicsWord, oGraphicsPDF) {
        let oStructure = this.GetRenderStructure();
        if (!oStructure) {
            return;
        }

        let nScale = this.GetOriginViewScale();
        let oTr = new AscCommon.CMatrix();
        
        // draw without rotate and scale for saving
        if (oGraphicsWord.isPdf()) {
            let hc = this.extX * 0.5;
            let vc = this.extY * 0.5;
            AscCommon.global_MatrixTransformer.TranslateAppend(oTr, -hc, -vc);
            AscCommon.global_MatrixTransformer.TranslateAppend(oTr, this.x + hc, this.y + vc);

            let aInRect = this.GetInRect();
            let nSourceX = aInRect[0] * g_dKoef_pt_to_mm;
            let nSourceY = aInRect[3] * g_dKoef_pt_to_mm;
            
            oTr.tx = nSourceX;
            oTr.ty = nSourceY;
        }
        else {
            oTr.Scale(nScale, nScale);
            let oOwnTr = this.getTransformMatrix();
            AscCommon.global_MatrixTransformer.MultiplyAppend(oTr, oOwnTr);
        }

        oStructure.draw(oGraphicsWord, oTr);
    };
    CAnnotationStamp.prototype.SetRenderStructure = function(oStructure) {
        AscCommon.History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_Pdf_Stamp_RenderStructure, this.renderStructure, oStructure));
        this.renderStructure = oStructure;
    };
    CAnnotationStamp.prototype.GetRenderStructure = function() {
        return this.renderStructure;
    };
    CAnnotationStamp.prototype.SetInRect = function(aInRect) {
        AscCommon.History.Add(new CChangesPDFAnnotStampInRect(this, this._inRect, aInRect));

        this._inRect = aInRect;
        
        function getDistance(x1, y1, x2, y2) {
            return Math.sqrt(Math.pow(x2 - x1, 2) + Math.pow(y2 - y1, 2));
        }

        if (!aInRect) {
            return;
        }
        
        let nShapeW = getDistance(aInRect[0], aInRect[1], aInRect[6], aInRect[7]);
        let nShapeH = getDistance(aInRect[0], aInRect[1], aInRect[2], aInRect[3]);

        if (nShapeH == 0 || nShapeW == 0) {
            return;
        }

        AscCommon.History.StartNoHistoryMode();
        this.spPr.xfrm.setExtX(nShapeW * g_dKoef_pt_to_mm);
        this.spPr.xfrm.setExtY(nShapeH * g_dKoef_pt_to_mm);

        let aRect = this.GetRect();
        
        let nAnnotW = aRect[2] - aRect[0];
        let nAnnotH = aRect[3] - aRect[1];
        
        let nOffX = (aRect[0] - (nShapeW - nAnnotW) / 2);
        let nOffY = (aRect[1] - (nShapeH - nAnnotH) / 2);
        
        this.spPr.xfrm.setOffX(nOffX * g_dKoef_pt_to_mm);
        this.spPr.xfrm.setOffY(nOffY * g_dKoef_pt_to_mm);

        AscCommon.History.EndNoHistoryMode();
    };
    CAnnotationStamp.prototype.GetInRect = function() {
        return this._inRect;
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
        let nCurrentH = this.getXfrmExtY();

        return nCurrentH / nSourceH;
    };
    CAnnotationStamp.prototype.DrawFromStream = function(oGraphicsPDF) {
        if (this.IsHidden() == true)
            return;
            
        let nScale      = this.GetOriginViewScale();
        let originView  = this.GetOriginView(oGraphicsPDF.GetDrawingPageW() * nScale, oGraphicsPDF.GetDrawingPageH() * nScale);
        let nRot        = -this.GetRotate() * Math.PI / 180;

        if (originView) {
            let oXfrm = this.getXfrm();
            
            let X = oXfrm.offX / g_dKoef_pt_to_mm;
            let Y = oXfrm.offY / g_dKoef_pt_to_mm;

            if (this.IsHighlight())
                AscPDF.startMultiplyMode(oGraphicsPDF.GetContext());
            
            oGraphicsPDF.SetGlobalAlpha(1);
            oGraphicsPDF.DrawImageXY(originView, X, Y, nRot, true);
            AscPDF.endMultiplyMode(oGraphicsPDF.GetContext());
        }

        // if (oGraphicsPDF) {
        //     oGraphicsPDF.SetLineWidth(1);
        //     let aOringRect  = this.GetRect();
        //     let X       = aOringRect[0];
        //     let Y       = aOringRect[1];
        //     let nWidth  = aOringRect[2] - aOringRect[0];
        //     let nHeight = aOringRect[3] - aOringRect[1];

        //     Y += 1 / 2;
        //     X += 1 / 2;
        //     nWidth  -= 1;
        //     nHeight -= 1;

        //     oGraphicsPDF.SetStrokeStyle(0, 255, 255);
        //     oGraphicsPDF.SetLineDash([]);
        //     oGraphicsPDF.BeginPath();
        //     oGraphicsPDF.Rect(X, Y, nWidth, nHeight);
        //     oGraphicsPDF.Stroke();
        // }
    };
    CAnnotationStamp.prototype.ClearCache = function() {
        this._originView.normal = null;
        this.APInfo = null;
    };
    CAnnotationStamp.prototype.GetOriginViewInfo = function(nPageW, nPageH) {
        let oViewer     = editor.getDocumentRenderer();
        let oFile       = oViewer.file;
        let nPage       = this.GetOriginPage();
        let nApIdx      = this.GetApIdx();

        if (this.APInfo == null || this.APInfo.size.w != nPageW || this.APInfo.size.h != nPageH) {
            this.APInfo = {
                info: oFile.nativeFile["getAnnotationsAP"](nPage, nPageW, nPageH, undefined, nApIdx),
                size: {
                    w: nPageW,
                    h: nPageH
                }
            }
        }
        
        for (let i = 0; i < this.APInfo.info.length; i++) {
            if (this.APInfo.info[i]["i"] == nApIdx)
                return this.APInfo.info[i];
        }

        return null;
    };
    CAnnotationStamp.prototype.SetPosition = function(x, y) {
        let aCurRect = this.GetRect();

        let nOldX = aCurRect[0];
        let nOldY = aCurRect[1];

        let nDeltaX = x - nOldX;
        let nDeltaY = y - nOldY;

        if (0 == nDeltaX && 0 == nDeltaY) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFAnnotPos(this, [aCurRect[0], aCurRect[1]], [x, y]));

        let nWidth  = aCurRect[2] - aCurRect[0];
        let nHeight = aCurRect[3] - aCurRect[1];

        this._rect[0] = x;
        this._rect[1] = y;
        this._rect[2] = x + nWidth;
        this._rect[3] = y + nHeight;
        
        let oXfrm = this.getXfrm();
        let nOffX = x - (this.getXfrmExtX() / g_dKoef_pt_to_mm - nWidth) / 2;
        let nOffY = y - (this.getXfrmExtY() / g_dKoef_pt_to_mm - nHeight) / 2;

        AscCommon.History.StartNoHistoryMode();
        oXfrm.setOffX(nOffX * g_dKoef_pt_to_mm);
        oXfrm.setOffY(nOffY * g_dKoef_pt_to_mm);
        AscCommon.History.EndNoHistoryMode();

        this.SetNeedRecalc(true);
        this.SetWasChanged(true);
    };
    CAnnotationStamp.prototype.canRotate = function() {
        return true;
    };
    CAnnotationStamp.prototype.recalculateBounds = function() {
        let sType = this.GetIconType();
        if (sType == undefined) {
            return;
        }

        let boundsChecker = new AscFormat.CSlideBoundsChecker();
        
        boundsChecker.DO_NOT_DRAW_ANIM_LABEL = true;
        if (AscPDF.STAMP_TYPES.Image == sType || !this.GetRenderStructure()) {
            this.draw(boundsChecker);
        } else {
            this._draw(boundsChecker);
        }

        this.bounds.x = boundsChecker.Bounds.min_x;
        this.bounds.y = boundsChecker.Bounds.min_y;
        this.bounds.l = boundsChecker.Bounds.min_x;
        this.bounds.t = boundsChecker.Bounds.min_y;
        this.bounds.r = boundsChecker.Bounds.max_x;
        this.bounds.b = boundsChecker.Bounds.max_y;
        this.bounds.w = boundsChecker.Bounds.max_x - boundsChecker.Bounds.min_x;
        this.bounds.h = boundsChecker.Bounds.max_y - boundsChecker.Bounds.min_y;
    };
    CAnnotationStamp.prototype.RefillGeometry = function(oGeometry, aBounds) {
        // skip on recalc
        if (!oGeometry) {
            return;
        }

        oGeometry.Recalculate(aBounds[2] - aBounds[0], aBounds[3] - aBounds[1]);
        return oGeometry;
    };
    CAnnotationStamp.prototype.IsSelected = function() {
        let oViewer         = editor.getDocumentRenderer();
        let oDrawingObjects = oViewer.DrawingObjects;
        return oDrawingObjects.selectedObjects.includes(this);
    };
    
    CAnnotationStamp.prototype.SetIconType = function(sType) {
        if (typeof(sType) == "string") {
            let aSplitted = sType.split('#');
            sType = aSplitted[aSplitted.length - 1];
        }

        if (sType == this._stampType) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFAnnotStampType(this, this._stampType, sType));
        this._stampType = sType;
        this.SetWasChanged(true);
    };
    CAnnotationStamp.prototype.GetIconType = function() {
        return this._stampType;
    };
    CAnnotationStamp.prototype.SetRotate = function(nAngle) {
        if (this._rotate == nAngle) {
            return;
        }
        
        AscCommon.History.Add(new CChangesPDFAnnotRotate(this, this._rotate, nAngle));

		let oXfrm = this.getXfrm();
        oXfrm.rot = -nAngle * (Math.PI / 180);
		
        this._rotate = nAngle;

        this.SetWasChanged(true);
        this.SetNeedRecalc(true);
    };
    CAnnotationStamp.prototype.GetRotate = function() {
        return this._rotate;
    };
    CAnnotationStamp.prototype.handleUpdateRot = function(){
        AscFormat.CShape.prototype.handleUpdateRot.call(this);
        let oXfrm = this.getXfrm();

        this.SetRotate(-oXfrm.rot * (180 / Math.PI));
        this.recalcBounds();
        this.recalcGeometry();
        this.Recalculate(true);
        
        let aNewRect = [];
        let oGrBounds = this.bounds;
        aNewRect[0] = Math.round(oGrBounds.l) * g_dKoef_mm_to_pt;
        aNewRect[1] = Math.round(oGrBounds.t) * g_dKoef_mm_to_pt;
        aNewRect[2] = Math.round(oGrBounds.r) * g_dKoef_mm_to_pt;
        aNewRect[3] = Math.round(oGrBounds.b) * g_dKoef_mm_to_pt;

        this.SetRect(aNewRect, true);
    };
    CAnnotationStamp.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        memory.WriteString(this.GetIconType());
        memory.WriteDouble(this.GetRotate());
        
        let aInRect = this.GetInRect();
        let nBorderW = this.GetWidth();
        let nScale = this.GetOriginViewScale();

        // original rect (save)
        if (memory.docRenderer) {
            memory.WriteDouble(aInRect[0] - (nBorderW / 2) * nScale); // x1
            memory.WriteDouble(aInRect[3] - (nBorderW / 2) * nScale); // y1
            memory.WriteDouble(aInRect[4] + (nBorderW / 2) * nScale); // x2
            memory.WriteDouble(aInRect[1] + (nBorderW / 2) * nScale); // y2
        }
        else { // copying
            aInRect.forEach(function(measure) {
                memory.WriteDouble(measure);
            });
        }

        let nEndPos = memory.GetCurPosition();
        memory.Seek(memory.posForFlags);
        memory.WriteLong(memory.annotFlags);
        
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };
    CAnnotationStamp.prototype.SetDrawFromStream = function(bDraw, bForce) {
        let oViewer = editor.getDocumentRenderer();
        if (oViewer.IsOpenAnnotsInProgress || bForce) {
            AscCommon.History.Add(new CChangesPDFAnnotChangedView(this, this._bDrawFromStream, bDraw));
            this._bDrawFromStream = bDraw;
        }
    };
    CAnnotationStamp.prototype.GetAllFonts = function(fontMap) {
        if (this.IsNeedDrawFromStream() || this.fontsChecked) {
            return fontMap;
        }

        let oContent = this.getDocContent();
        if (!oContent) {
            return fontMap;
        }
        
        oContent.SetApplyToAll(true);
        AscFonts.FontPickerByCharacter.getFontsByString(oContent.GetSelectedText());
        oContent.SetApplyToAll(false);
        this.fontsChecked = true;
        this.renderStructure = null;
        this.AddToRedraw()

        return fontMap;
    };
    CAnnotationStamp.prototype.WriteRenderToBinary = function(memory) {
        // пока только для основанных на фигурах
        if (this.IsNeedDrawFromStream() || !memory.docRenderer || (memory.isForSplit || memory.isCopyPaste)) {
            return;
        }

        // тут будет длина комманд
        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.Draw(undefined, memory.docRenderer); // для каждой страницы инициализируется свой renderer

        // запись длины комманд
        let nEndPos = memory.GetCurPosition();
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };
    
    CAnnotationStamp.prototype.getNoChangeAspect = function() {
        return true;
    };
    CAnnotationStamp.prototype.hitInPath = function(x, y) {
        return this.hitInInnerArea(x, y);
    };
    CAnnotationStamp.prototype.Init = function() {
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
    };
    
    window["AscPDF"].CAnnotationStamp = CAnnotationStamp;
    window["AscPDF"].STAMP_TYPES = window["AscPDF"]["STAMP_TYPES"] = STAMP_TYPES;
})();

