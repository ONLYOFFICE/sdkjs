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

    /**
	 * Class representing a Ink annotation.
	 * @constructor
    */
    function CAnnotationSquare(sName, aRect, oDoc)
    {
        AscPDF.CPdfShape.call(this);
        AscPDF.CAnnotationBase.call(this, sName, AscPDF.ANNOTATIONS_TYPES.Square, aRect, oDoc);
        
        AscPDF.initShape(this);
        this.spPr.setGeometry(AscFormat.CreateGeometry("rect"));

        this._point         = undefined;
        this._popupOpen     = false;
        this._popupRect     = undefined;
        this._richContents  = undefined;
        this._rotate        = undefined;
        this._state         = undefined;
        this._stateModel    = undefined;
        this._rectDiff      = [0, 0, 0, 0];
    }
    CAnnotationSquare.prototype.constructor = CAnnotationSquare;
    AscFormat.InitClass(CAnnotationSquare, AscPDF.CPdfShape, AscDFH.historyitem_type_Pdf_Annot_Square);
    Object.assign(CAnnotationSquare.prototype, AscPDF.CAnnotationBase.prototype);

	CAnnotationSquare.prototype.private_UpdateRect = function(rect) {
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

    CAnnotationSquare.prototype.RefillGeometry = function(oGeometry, aShapeRectInMM) {
        if (this.GetBorderEffectStyle() !== AscPDF.BORDER_EFFECT_STYLES.cloud)
            return;

        let aOrigRect   = this.GetRect();
        let aRD         = this.GetRectangleDiff() || [0, 0, 0, 0];

        let aPoints;
        if (!oGeometry)
            oGeometry = this.spPr.geometry;
        if (!aShapeRectInMM) {
            aShapeRectInMM = [
                (aOrigRect[0] + aRD[0]) * g_dKoef_pt_to_mm, (aOrigRect[1] + aRD[1]) * g_dKoef_pt_to_mm,
                (aOrigRect[2] - aRD[2]) * g_dKoef_pt_to_mm, (aOrigRect[3] - aRD[3]) * g_dKoef_pt_to_mm
            ];

            aPoints = [
                {x: (aOrigRect[0] + aRD[0]) * g_dKoef_pt_to_mm, y: (aOrigRect[1] + aRD[1]) * g_dKoef_pt_to_mm},
                {x: (aOrigRect[2] - aRD[2]) * g_dKoef_pt_to_mm, y: (aOrigRect[1] + aRD[1]) * g_dKoef_pt_to_mm},
                {x: (aOrigRect[2] - aRD[2]) * g_dKoef_pt_to_mm, y: (aOrigRect[3] - aRD[3]) * g_dKoef_pt_to_mm},
                {x: (aOrigRect[0] + aRD[0]) * g_dKoef_pt_to_mm, y: (aOrigRect[3] - aRD[3]) * g_dKoef_pt_to_mm}
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

        AscCommon.History.StartNoHistoryMode();
        AscPDF.generateCloudyGeometry(aPoints, aShapeRectInMM, oGeometry, this.GetBorderEffectIntensity());
        AscCommon.History.EndNoHistoryMode();

        oGeometry.preset = undefined;
    };
    CAnnotationSquare.prototype.SetDefaultGeometry = function() {
        AscCommon.History.StartNoHistoryMode();
        this.spPr.setGeometry(AscFormat.CreateGeometry("rect"));
        AscCommon.History.EndNoHistoryMode();
    };
    CAnnotationSquare.prototype.IsSquare = function() {
        return true;
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

