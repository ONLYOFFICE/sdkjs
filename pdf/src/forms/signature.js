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
	 * Class representing a signature field.
	 * @constructor
     * @extends {CBaseField}
	 */
    function CSignatureField(sName, aRect, oDoc)
    {
        AscPDF.CBaseField.call(this, sName, AscPDF.FIELD_TYPES.signature, aRect, oDoc);
        this._filled = false;
    };

    CSignatureField.prototype.constructor = CSignatureField;
    AscFormat.InitClass(CSignatureField, AscPDF.CBaseField, AscDFH.historyitem_type_Pdf_Signature_Field);
    
    CSignatureField.prototype.SetValue = function() {
        return;
    };
    CSignatureField.prototype.private_SetValue = CSignatureField.prototype.SetValue
    CSignatureField.prototype.Draw = function(oGraphicsPDF, oGraphicsWord) {
        return;
    };
    CSignatureField.prototype.DrawPressed = function() {
        return;
    };
    CSignatureField.prototype.DrawUnpressed = function() {
        return;
    };
    CSignatureField.prototype.Recalculate = function() {
    };

    CSignatureField.prototype.SetPressed = function(bValue) {
        this._pressed = bValue;
    };
    CSignatureField.prototype.IsPressed = function() {
        return this._pressed;
    };
    CSignatureField.prototype.IsHovered = function() {
        return this._hovered;
    };
    CSignatureField.prototype.SetHovered = function(bValue) {
        this._hovered = bValue;
    };

    CSignatureField.prototype.onMouseDown = function(x, y, e) {
    };
    CSignatureField.prototype.onMouseUp = function() {
    };

    CSignatureField.prototype.SetFilled = function(bValue) {
        this._filled = bValue;
        this.SetDrawHighlight(!bValue);
    };
    CSignatureField.prototype.IsFilled = function() {
        return this._filled;
    };
    CSignatureField.prototype.SetDrawHighlight = function(bDraw) {
        if (this.IsFilled()) {
            this._needDrawHighlight = false;
        }
        else {
            this._needDrawHighlight = bDraw;
        }
    };
    /**
     * Synchronizes this field with fields with the same name.
     * @memberof CSignatureField
     * @typeofeditors ["PDF"]
     */
    CSignatureField.prototype.SyncValue = function() {
    };
    /**
     * Applies value of this field to all field with the same name.
     * @memberof CSignatureField
     * @typeofeditors ["PDF"]
     */
    CSignatureField.prototype.Commit = function() {
    };

    CSignatureField.prototype.Reset = function() {
    };
	
    CSignatureField.prototype.WriteToBinary = function(memory) {};
    function MakeColorMoreGray(rgbColor, nPower) {
        // Get color component values
        const r = rgbColor.r;
        const g = rgbColor.g;
        const b = rgbColor.b;
      
        // Calculate new component values with darkening (reducing intensity)
        const grayR = Math.max(0, r - nPower);
        const grayG = Math.max(0, g - nPower);
        const grayB = Math.max(0, b - nPower);
      
        // Return new gray color
        return {
            r: grayR,
            g: grayG,
            b: grayB
        };
    }

    window["AscPDF"].CSignatureField = CSignatureField;
})();

