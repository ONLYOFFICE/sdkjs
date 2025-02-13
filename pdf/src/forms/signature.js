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
	 * Class representing a signature field.
	 * @constructor
     * @extends {CBaseField}
	 */
    function CSignatureField(sName, nPage, aRect, oDoc)
    {
        AscPDF.CBaseField.call(this, sName, AscPDF.FIELD_TYPES.signature, nPage, aRect, oDoc);
        this._filled = false;
    };

    CSignatureField.prototype = Object.create(AscPDF.CBaseField.prototype);
    CSignatureField.prototype.constructor = CSignatureField;
    
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
    /**
     * Synchronizes this field with fields with the same name.
     * @memberof CSignatureField
     * @typeofeditors ["PDF"]
     */
    CSignatureField.prototype.SyncField = function() {
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
	
    function MakeColorMoreGray(rgbColor, nPower) {
        // Получаем значения компонентов цвета
        const r = rgbColor.r;
        const g = rgbColor.g;
        const b = rgbColor.b;
      
        // Вычисляем новые значения компонентов с учетом затемнения (уменьшения интенсивности)
        const grayR = Math.max(0, r - nPower);
        const grayG = Math.max(0, g - nPower);
        const grayB = Math.max(0, b - nPower);
      
        // Возвращаем новый серый цвет
        return {
            r: grayR,
            g: grayG,
            b: grayB
        };
    }

    window["AscPDF"].CSignatureField = CSignatureField;
})();

