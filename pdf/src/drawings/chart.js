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
	 * Class representing a pdf text shape.
	 * @constructor
    */
    function CPdfChartSpace() {
        AscFormat.CChartSpace.call(this);
        AscPDF.CPdfDrawingPrototype.call(this);
    }
    
    CPdfChartSpace.prototype.constructor = CPdfChartSpace;
    CPdfChartSpace.prototype = Object.create(AscFormat.CChartSpace.prototype);
    Object.assign(CPdfChartSpace.prototype, AscPDF.CPdfDrawingPrototype.prototype);

    CPdfChartSpace.prototype.IsChart = function() {
        return true;
    };
    CPdfChartSpace.prototype.canRotate = function() {
        return true;
    };
    CPdfChartSpace.prototype.createRotateTrack = function () {
        return new AscFormat.RotateTrackShapeImage(this);
    };
    CPdfChartSpace.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        this.recalculateTransform();
        this.recalcGeometry();
        this.recalculate();
        this.updateTransformMatrixPDF();
        this.SetNeedRecalc(false);
    };


	CPdfChartSpace.prototype.updateTransformMatrixPDF = function() {
		this.posX = 0;
		this.posY = 0;
		this.updateTransformMatrix();
		let posX = this.localTransform.tx + this.posX;
		let posY = this.localTransform.ty + this.posY;
		let updateMatrix = new AscCommon.CMatrix();
		AscCommon.global_MatrixTransformer.TranslateAppend(updateMatrix, -posX, -posY);
		AscCommon.global_MatrixTransformer.MultiplyAppend(updateMatrix, this.localTransform);
		this.checkShapeChildTransform(updateMatrix);
	};
    CPdfChartSpace.prototype.onMouseDown = function(x, y, e) {
        let oViewer             = Asc.editor.getDocumentRenderer();
        let oDoc                = this.GetDocument();
        let oDrawingObjects     = oDoc.Viewer.DrawingObjects;
        this.selectStartPage    = this.GetPage();

        let pageObject = oViewer.getPageByCoords2(x, y);
        if (!pageObject)
            return false;

        let X = pageObject.x;
        let Y = pageObject.y;

        oDrawingObjects.OnMouseDown(e, X, Y, pageObject.index);
    };
    CPdfChartSpace.prototype.onMouseUp = function(x, y, e) {
        let oViewer         = Asc.editor.getDocumentRenderer();
        
        this.selectStartPage    = this.GetPage();
        let oContent            = this.GetDocContent();

        if (global_mouseEvent.ClickCount == 2) {
            oContent.SelectAll();
            if (oContent.IsSelectionEmpty() == false)
                oViewer.Api.WordControl.m_oDrawingDocument.TargetEnd();
            else
                oContent.RemoveSelection();
        }
                
        if (oContent.IsSelectionEmpty())
            oContent.RemoveSelection();
    };
    CPdfChartSpace.prototype.GetDocContent = function() {
        let oTextSelection = this.selection.textSelection;
        if (oTextSelection) {
            return oTextSelection.getDocContent();
        }

        return null;
    };

    /////////////////////////////
    /// saving
    ////////////////////////////

    CPdfChartSpace.prototype.WriteToBinary = function(memory) {
        this.toXml(memory, '');
    };

    /**
     * Get calculated text settings (fully filled)
     * @returns {CTextPr}
     */
    CPdfChartSpace.prototype.GetCalculatedTextPr = function() {
        return this.GetDocContent().GetCalculatedTextPr();
    };
    CPdfChartSpace.prototype.GetCalculatedParaPr = function() {
        return this.GetDocContent().GetCalculatedParaPr();
    };

    //////////////////////////////////////////////////////////////////////////////
    ///// Overrides
    /////////////////////////////////////////////////////////////////////////////
    
    CPdfChartSpace.prototype.getLogicDocument = function() {
        return this.GetDocument();
    };
    CPdfChartSpace.prototype.IsThisElementCurrent = function() {
        return true;
    };
    CPdfChartSpace.prototype.copy = function (oPr) {
		let drawingDocument = oPr && oPr.drawingDocument;
		let copy = new CPdfChartSpace();
		if (this.chart) {
			copy.setChart(this.chart.createDuplicate(drawingDocument));
		}
		if (this.clrMapOvr) {
			copy.setClrMapOvr(this.clrMapOvr.createDuplicate());
		}
		copy.setDate1904(this.date1904);
		if (this.externalData) {
			const oCopyExternalData = {};
			if (this.externalData.m_autoUpdate)
			{
				oCopyExternalData.m_autoUpdate = {m_val: this.externalData.m_autoUpdate.m_val};
			}
			copy.setExternalData(oCopyExternalData);
		}
		if (this.XLSX) {
			copy.setXLSX(this.XLSX.slice());
		}
		if (this.externalReference) {
			copy.setExternalReference(this.externalReference.createDuplicate());
		}
		copy.setLang(this.lang);
		if (this.pivotSource) {
			copy.setPivotSource(this.pivotSource.createDuplicate());
		}
		if (this.printSettings) {
			copy.setPrintSettings(this.printSettings.createDuplicate());
		}
		if (this.protection) {
			copy.setProtection(this.protection.createDuplicate());
		}
		copy.setRoundedCorners(this.roundedCorners);
		if (this.spPr) {
			copy.setSpPr(this.spPr.createDuplicate());
			copy.spPr.setParent(copy);
		}
		copy.setStyle(this.style);
		if (this.txPr) {
			copy.setTxPr(this.txPr.createDuplicate(oPr))
		}
		for (var i = 0; i < this.userShapes.length; ++i) {
			copy.addUserShape(undefined, this.userShapes[i].copy(oPr));
		}
		copy.setThemeOverride(this.themeOverride);
		copy.setBDeleted(this.bDeleted);
		copy.setLocks(this.locks);
		if (this.chartStyle && this.chartColors) {
			copy.setChartStyle(this.chartStyle.createDuplicate());
			copy.setChartColors(this.chartColors.createDuplicate());
		}
		if (this.macro !== null) {
			copy.setMacro(this.macro);
		}
		if (this.textLink !== null) {
			copy.setTextLink(this.textLink);
		}
		if(this.chartData) {
			copy.setChartData(this.chartData.createDuplicate());
		}
		if (!oPr || false !== oPr.cacheImage) {
			copy.cachedImage = this.getBase64Img();
			copy.cachedPixH = this.cachedPixH;
			copy.cachedPixW = this.cachedPixW;
		}

		if ((!oPr || !oPr.bSkipRedactsIds) && this.GetRedactIds) {
            this.GetRedactIds().forEach(function(id) {
                copy.AddRedactId(id);
            });
        }
		
		return copy;
	};
		CPdfChartSpace.prototype.applySpecialPasteProps = function() {
			this.setExternalReference(null);
			this.setXLSX(new Uint8Array(0));
		};

    window["AscPDF"].CPdfChartSpace = CPdfChartSpace;
})();

