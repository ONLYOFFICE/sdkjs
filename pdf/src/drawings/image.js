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

(function() {

    /**
	 * Class representing a pdf text shape.
	 * @constructor
    */
    function CPdfImage()
    {
        AscFormat.CImageShape.call(this);
        AscPDF.CPdfDrawingPrototype.call(this);
    }
    
    CPdfImage.prototype.constructor = CPdfImage;
    CPdfImage.prototype = Object.create(AscFormat.CImageShape.prototype);
    Object.assign(CPdfImage.prototype, AscPDF.CPdfDrawingPrototype.prototype);

    CPdfImage.prototype.IsImage = function() {
        return true;
    };
    CPdfImage.prototype.copy = function (oPr) {
        let copy = this.convertToPdf();

        if ((!oPr || !oPr.bSkipRedactsIds) && this.GetRedactIds) {
            this.GetRedactIds().forEach(function(id) {
                copy.AddRedactId(id);
            });
        }

        return copy;
    };
    CPdfImage.prototype.IsInTextBox = function() {
        return false;
    };
    CPdfImage.prototype.GetDocContent = function() {
        return null;
    };
    CPdfImage.prototype.Recalculate = function() {
        if (this.IsNeedRecalc() == false)
            return;

        this.recalcGeometry();
        this.recalculate();
        this.recalculateTransform();
        this.updateTransformMatrix();
        this.recalculateShdw();
        this.SetNeedRecalc(false);
    };
    CPdfImage.prototype.onMouseDown = function(x, y, e) {
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
    
    //////////////////////////////////////////////////////////////////////////////
    ///// Overrides
    /////////////////////////////////////////////////////////////////////////////
    

    window["AscPDF"].CPdfImage = CPdfImage;
})();

