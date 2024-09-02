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
	 * Class representing a Ink annotation.
	 * @constructor
    */
    function CAnnotationSquare(sName, nPage, aRect, oDoc)
    {
        AscPDF.CPdfShape.call(this);
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Square, nPage, aRect, oDoc);
        
        AscPDF.initShape(this);
        this.spPr.setGeometry(AscFormat.CreateGeometry("rect"));

        this._point         = undefined;
        this._popupOpen     = false;
        this._popupRect     = undefined;
        this._richContents  = undefined;
        this._rotate        = undefined;
        this._state         = undefined;
        this._stateModel    = undefined;
        this._width         = undefined;
        this._rectDiff      = [0, 0, 0, 0];
    }
    CAnnotationSquare.prototype.constructor = CAnnotationSquare;
    AscFormat.InitClass(CAnnotationSquare, AscPDF.CPdfShape, AscDFH.historyitem_type_Pdf_Annot_Square);
    Object.assign(CAnnotationSquare.prototype, AscPDF.CAnnotationBase.prototype);

    CAnnotationSquare.prototype.LazyCopy = function() {
        let oDoc = this.GetDocument();
        oDoc.StartNoHistoryMode();

        let oSquare = new CAnnotationSquare(AscCommon.CreateGUID(), this.GetPage(), this.GetOrigRect().slice(), oDoc);

        oSquare.lazyCopy = true;

        oSquare._pagePos = {
            x: this._pagePos.x,
            y: this._pagePos.y,
            w: this._pagePos.w,
            h: this._pagePos.h
        }

        oSquare._origRect = this._origRect.slice();
        oSquare._rect = this._rect.slice();

        this.fillObject(oSquare);

        let aStrokeColor = this.GetStrokeColor();
        let aFillColor = this.GetFillColor();

        oSquare._apIdx = this._apIdx;
        oSquare._originView = this._originView;
        oSquare.SetOriginPage(this.GetOriginPage());
        oSquare.SetAuthor(this.GetAuthor());
        oSquare.SetModDate(this.GetModDate());
        oSquare.SetCreationDate(this.GetCreationDate());
        oSquare.SetStrokeColor(aStrokeColor.slice());
        oSquare.SetFillColor(aFillColor.slice());
        oSquare.SetWidth(this.GetWidth());
        oSquare.SetOpacity(this.GetOpacity());
        oSquare.recalcGeometry()
        oSquare.SetRectangleDiff(this.GetRectangleDiff().slice(), true);
        oSquare.Recalculate(true);

        oDoc.EndNoHistoryMode();

        return oSquare;
    };
    CAnnotationSquare.prototype.RefillGeometry = function(oGeometry, aShapeRectInMM) {
        if (this.GetBorderEffectStyle() !== AscPDF.BORDER_EFFECT_STYLES.Cloud)
            return;

        let oViewer     = editor.getDocumentRenderer();
        let nPage       = this.GetPage();
        let oDoc        = this.GetDocument();

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom * g_dKoef_pix_to_mm;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom * g_dKoef_pix_to_mm;

        let aOrigRect   = this.GetOrigRect();
        let aRD         = this.GetRectangleDiff() || [0, 0, 0, 0];

        let aPoints;
        if (!oGeometry)
            oGeometry = this.spPr.geometry;
        if (!aShapeRectInMM) {
            aShapeRectInMM = [
                (aOrigRect[0] + aRD[0]) * nScaleX, (aOrigRect[1] + aRD[1]) * nScaleY,
                (aOrigRect[2] - aRD[2]) * nScaleX, (aOrigRect[3] - aRD[3]) * nScaleY
            ];

            aPoints = [
                {x: (aOrigRect[0] + aRD[0]) * nScaleX, y: (aOrigRect[1] + aRD[1]) * nScaleY},
                {x: (aOrigRect[2] - aRD[2]) * nScaleX, y: (aOrigRect[1] + aRD[1]) * nScaleY},
                {x: (aOrigRect[2] - aRD[2]) * nScaleX, y: (aOrigRect[3] - aRD[3]) * nScaleY},
                {x: (aOrigRect[0] + aRD[0]) * nScaleX, y: (aOrigRect[3] - aRD[3]) * nScaleY}
            ]
        }
        else {
            aPoints = [
                {x: aShapeRectInMM[0], y: aShapeRectInMM[1]},
                {x: aShapeRectInMM[2], y: aShapeRectInMM[1]},
                {x: aShapeRectInMM[2], y: aShapeRectInMM[3]},
                {x: aShapeRectInMM[0], y: aShapeRectInMM[3]}
            ]
        }

        oDoc.StartNoHistoryMode();
        AscPDF.generateCloudyGeometry(aPoints, aShapeRectInMM, oGeometry, this.GetBorderEffectIntensity());
        oDoc.EndNoHistoryMode();

        oGeometry.preset = undefined;
    };
    CAnnotationSquare.prototype.SetRect = function(aRect) {
        let oViewer     = editor.getDocumentRenderer();
        let oDoc        = oViewer.getPDFDoc();
        let nPage       = this.GetPage();
        let aCurRect    = this.GetRect();

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        this._rect = aRect;
        this._pagePos = {
            x: aRect[0],
            y: aRect[1],
            w: (aRect[2] - aRect[0]),
            h: (aRect[3] - aRect[1])
        };

        AscCommon.History.StartNoHistoryMode();
        let oXfrm = this.getXfrm();
        oXfrm.setOffX(aRect[0] * g_dKoef_pix_to_mm);
        oXfrm.setOffY(aRect[1] * g_dKoef_pix_to_mm);
        oXfrm.setExtX((aRect[2] - aRect[0]) * g_dKoef_pix_to_mm);
        oXfrm.setExtY((aRect[3] - aRect[1]) * g_dKoef_pix_to_mm);

        this._origRect[0] = this._rect[0] / nScaleX;
        this._origRect[1] = this._rect[1] / nScaleY;
        this._origRect[2] = this._rect[2] / nScaleX;
        this._origRect[3] = this._rect[3] / nScaleY;

        if (false == AscCommon.History.UndoRedoInProgress) {
            let aCurRD = this._rectDiff;
            this.recalcBounds();
            this.recalcGeometry();
            this.Recalculate(true);
            this.recalcInfo.recalculateGeometry = false;
            AscCommon.History.EndNoHistoryMode();

            let oGrBounds = this.bounds;
            let oShapeBounds = this.getRectBounds();

            this._origRect[0] = Math.round(oGrBounds.l) * g_dKoef_mm_to_pix / nScaleX;
            this._origRect[1] = Math.round(oGrBounds.t) * g_dKoef_mm_to_pix / nScaleY;
            this._origRect[2] = Math.round(oGrBounds.r) * g_dKoef_mm_to_pix / nScaleX;
            this._origRect[3] = Math.round(oGrBounds.b) * g_dKoef_mm_to_pix / nScaleY;

            this._rectDiff = aCurRD;
            this.SetRectangleDiff([
                Math.round(oShapeBounds.l - oGrBounds.l) * g_dKoef_mm_to_pix / nScaleX,
                Math.round(oShapeBounds.t - oGrBounds.t) * g_dKoef_mm_to_pix / nScaleY,
                Math.round(oGrBounds.r - oShapeBounds.r) * g_dKoef_mm_to_pix / nScaleX,
                Math.round(oGrBounds.b - oShapeBounds.b) * g_dKoef_mm_to_pix / nScaleY
            ], true);

            oDoc.History.Add(new CChangesPDFAnnotRect(this, aCurRect, aRect));
        }

        this.SetWasChanged(true);
    };
    CAnnotationSquare.prototype.SetRectangleDiff = function(aDiff, bOnResize) {
        let oDoc = this.GetDocument();
        oDoc.History.Add(new CChangesPDFAnnotRD(this, this.GetRectangleDiff(), aDiff));

        this._rectDiff = aDiff;

        if (true != bOnResize) {
            let oViewer     = editor.getDocumentRenderer();
            let nPage       = this.GetPage();

            let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom * g_dKoef_pix_to_mm;
            let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom * g_dKoef_pix_to_mm;

            let aOrigRect = this.GetOrigRect();

            let extX = ((aOrigRect[2] - aOrigRect[0]) - aDiff[0] - aDiff[2]) * nScaleX;
            let extY = ((aOrigRect[3] - aOrigRect[1]) - aDiff[1] - aDiff[3]) * nScaleY;

            this.spPr.xfrm.setOffX(aDiff[0] * nScaleX + this.spPr.xfrm.offX);
            this.spPr.xfrm.setOffY(aDiff[1] * nScaleY + this.spPr.xfrm.offY);

            this.spPr.xfrm.setExtX(extX);
            this.spPr.xfrm.setExtY(extY);
        }
    };
    CAnnotationSquare.prototype.IsSquare = function() {
        return true;
    };
    CAnnotationSquare.prototype.Recalculate = function(bForce) {
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
    
    CAnnotationSquare.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(AscCommon.CommandType.ctAnnotField);

        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        this.WriteToBinaryBase(memory);
        this.WriteToBinaryBase2(memory);
        
        // rectangle diff
        let aRD = this.GetRectangleDiff();
        if (aRD) {
            memory.annotFlags |= (1 << 15);
            for (let i = 0; i < 4; i++) {
                memory.WriteDouble(aRD[i]);
            }
        }
        
        // fill
        let aFill = this.GetFillColor();
        if (aFill != null) {
            memory.annotFlags |= (1 << 16);
            memory.WriteLong(aFill.length);
            for (let i = 0; i < aFill.length; i++)
                memory.WriteDouble(aFill[i]);
        }

        let nEndPos = memory.GetCurPosition();
        memory.Seek(memory.posForFlags);
        memory.WriteLong(memory.annotFlags);
        
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };

    window["AscPDF"].CAnnotationSquare = CAnnotationSquare;
})();

