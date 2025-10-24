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
     * Получаем рассчитанные настройки текста (полностью заполненные)
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

    window["AscPDF"].CPdfChartSpace = CPdfChartSpace;
})();

