/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

(function(){

    let POLYGON_INTENT_TYPE = {
        PolygonCloud:       0,
        PolyLineDimension:  1,
        PolygonDimension:   2
    }

    /**
	 * Class representing a Ink annotation.
	 * @constructor
    */
    function CAnnotationPolygon(sName, aRect, oDoc)
    {
        AscPDF.CPdfShape.call(this);
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Polygon, aRect, oDoc);
        
        AscPDF.initShape(this);

        this._point         = undefined;
        this._popupOpen     = false;
        this._popupRect     = undefined;
        this._richContents  = undefined;
        this._rotate        = undefined;
        this._state         = undefined;
        this._stateModel    = undefined;
        this._vertices      = undefined;
        this._intent        = undefined;
    }
    CAnnotationPolygon.prototype.constructor = CAnnotationPolygon;
    AscFormat.InitClass(CAnnotationPolygon, AscPDF.CPdfShape, AscDFH.historyitem_type_Pdf_Annot_Polygon);
    Object.assign(CAnnotationPolygon.prototype, AscPDF.CAnnotationBase.prototype);
    
	CAnnotationPolygon.prototype.private_UpdateRect = function(rect) {
		AscCommon.History.StartNoHistoryMode();
		let aCurRect = this.GetRect();
		let aCurRD = this.GetRectangleDiff().slice();
		let nLineW = this.GetBorderWidth() * g_dKoef_pt_to_mm;
		rect && this.SetRect(rect);
		this.SetRectangleDiff([0, 0, 0, 0]);
		this.recalcBounds();
		this.recalcGeometry();
		this.Recalculate(true);
		
		AscCommon.History.EndNoHistoryMode();
		
		let oGrBounds = this.bounds;
		let oShapeBounds = this.getRectBounds();

		if (!rect) {
			rect = [];
		}

		rect[0] = (oGrBounds.l - nLineW) * g_dKoef_mm_to_pt;
		rect[1] = (oGrBounds.t - nLineW) * g_dKoef_mm_to_pt;
		rect[2] = (oGrBounds.r + nLineW) * g_dKoef_mm_to_pt;
		rect[3] = (oGrBounds.b + nLineW) * g_dKoef_mm_to_pt;

		this._rect = aCurRect;
		this._rectDiff = aCurRD;

		this.SetRect(rect);
		this.SetRectangleDiff([
			(oShapeBounds.l - oGrBounds.l + nLineW) * g_dKoef_mm_to_pt,
			(oShapeBounds.t - oGrBounds.t + nLineW) * g_dKoef_mm_to_pt,
			(oGrBounds.r - oShapeBounds.r + nLineW) * g_dKoef_mm_to_pt,
			(oGrBounds.b - oShapeBounds.b + nLineW) * g_dKoef_mm_to_pt
		]);
	};

    CAnnotationPolygon.prototype.Copy = function(isForMove) {
        let oCopy = AscPDF.CAnnotationBase.prototype.Copy.call(this, isForMove);

        let aVertices = this.GetVertices();

        oCopy.SetVertices(aVertices.slice());

        return oCopy;
    };
    CAnnotationPolygon.prototype.SetVertices = function(aVertices) {
        AscCommon.History.Add(new CChangesPDFAnnotVertices(this, this.GetVertices(), aVertices));

        this._vertices = aVertices;
		this.recalcGeometry();
		this.SetWasChanged(true);
		this.SetNeedRecalc(true);
    };
    CAnnotationPolygon.prototype.GetVertices = function() {
        return this._vertices;
    };

    CAnnotationPolygon.prototype.RefillGeometry = function() {
        let aPoints = this.GetVertices();

        let aPolygonPoints = [];
        for (let i = 0; i < aPoints.length - 1; i += 2) {
            aPolygonPoints.push({
                x: aPoints[i] * g_dKoef_pt_to_mm,
                y: (aPoints[i + 1])* g_dKoef_pt_to_mm
            });
        }
        
        let aShapeRectInMM = this.GetRect().map(function(measure) {
            return measure * g_dKoef_pt_to_mm;
        });

        AscCommon.History.StartNoHistoryMode();

        let geometry;
        if (this.GetBorderEffectStyle() === AscPDF.BORDER_EFFECT_STYLES.cloud) {
            geometry = AscPDF.generateCloudyGeometry(aPolygonPoints, aShapeRectInMM, this.spPr.geometry, this.GetBorderEffectIntensity());
        }
        else {
            geometry = generateGeometry(aPolygonPoints, aShapeRectInMM, this.spPr.geometry);
        }

        if (this.spPr.geometry == null)
            this.spPr.setGeometry(geometry);

        AscCommon.History.EndNoHistoryMode();
    };
    CAnnotationPolygon.prototype.onMouseDown = function(x, y, e) {
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
    CAnnotationPolygon.prototype.GetGeometryEdit = function() {
        if (this.GetBorderEffectStyle() !== AscPDF.BORDER_EFFECT_STYLES.cloud)
            return this.spPr.geometry;
        
        let aPoints = this.GetVertices();
        let aPolygonPoints = [];
        for (let i = 0; i < aPoints.length - 1; i += 2) {
            aPolygonPoints.push({
                x: aPoints[i] * g_dKoef_pt_to_mm,
                y: (aPoints[i + 1])* g_dKoef_pt_to_mm
            });
        }
        
        let aShapeRectInMM = this.GetRect().map(function(measure) {
            return measure * g_dKoef_pt_to_mm;
        });

        AscCommon.History.StartNoHistoryMode();
        this._internalGeomForEdit = generateGeometry(aPolygonPoints, aShapeRectInMM, this._internalGeomForEdit);
        this._internalGeomForEdit.Recalculate(aShapeRectInMM[2] - aShapeRectInMM[0], aShapeRectInMM[3] - aShapeRectInMM[1]);
        AscCommon.History.EndNoHistoryMode();

        return this._internalGeomForEdit;
    };
    CAnnotationPolygon.prototype.IsPolygon = function() {
        return true;
    };
    CAnnotationPolygon.prototype.WriteToBinary = function(memory) {
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
        let nPtCount = aPoints.length;
        // if the last point matches the first one, we don't count it
        if (aPoints[0].x == aPoints[aPoints.length - 1].x && aPoints[0].y == aPoints[aPoints.length - 1].y)
            nPtCount = aPoints.length - 1;

        for(nPt = 1; nPt < nPtCount; nPt++) {
            oPt = aPoints[nPt];

            geometry.AddPathCommand(2,
                (((oPt.x - xMin) * kw) >> 0) + "", (((oPt.y - yMin) * kh) >> 0) + ""
            );
        }
        
        geometry.AddPathCommand(6);

        geometry.preset = null;
        geometry.rectS = null;
        return geometry;
    }
    
    window["AscPDF"].CAnnotationPolygon = CAnnotationPolygon;
})();

