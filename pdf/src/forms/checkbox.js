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
	 * Class representing a checkbox field.
	 * @constructor
     * @extends {CBaseCheckBoxField}
	 */
    function CCheckBoxField(sName, aRect, oDoc)
    {
        AscPDF.CBaseCheckBoxField.call(this, sName, AscPDF.FIELD_TYPES.checkbox, aRect, oDoc);

        this._noToggleToOff = false;
        this._chStyle       = AscPDF.CHECKBOX_STYLES.check;
        this._caption       = undefined;
    }
    CCheckBoxField.prototype.constructor = CCheckBoxField;
    AscFormat.InitClass(CCheckBoxField, AscPDF.CBaseCheckBoxField, AscDFH.historyitem_type_Pdf_Checkbox_Field);
    
    /**
	 * Applies value of this field to all field with the same name
     * Note: Uses after mouseUp action.
	 * @memberof CCheckBoxField
	 * @typeofeditors ["PDF"]
	 */
    CCheckBoxField.prototype.Commit2 = function() {
        let aFields = this.GetAllWidgets();
        let oThis = this;

        aFields.forEach(function(field) {
            if (field == oThis)
                return;

            if (field.GetExportValue() != oThis.GetExportValue() && field.IsChecked() == true) {
                field.SetChecked(false);
                field.SetNeedRecalc(true);
            }
            else if (field.GetExportValue() == oThis.GetExportValue() && oThis.IsChecked() == false) {
                field.SetChecked(false);
                field.SetNeedRecalc(true);
            }
            else if (field.GetExportValue() == oThis.GetExportValue() && field.IsChecked() == false) {
                field.SetChecked(true);
                field.SetNeedRecalc(true);
            }
        });
    };

    if (!window["AscPDF"])
	    window["AscPDF"] = {};
        
	window["AscPDF"].CCheckBoxField = CCheckBoxField;
})();

