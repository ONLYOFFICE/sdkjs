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
	 * Class representing a radiobutton field.
	 * @constructor
     * @extends {CBaseCheckBoxField}
	 */
    function CRadioButtonField(sName, aRect, oDoc)
    {
        AscPDF.CBaseCheckBoxField.call(this, sName, AscPDF.FIELD_TYPES.radiobutton, aRect, oDoc);
        
        this._exportValue   = "Choice1";
        this._radiosInUnison = false;
        this._noToggleToOff = true;

        this._chStyle       = AscPDF.CHECKBOX_STYLES.circle;
    }
	CRadioButtonField.prototype.constructor = CRadioButtonField;
    AscFormat.InitClass(CRadioButtonField, AscPDF.CBaseCheckBoxField, AscDFH.historyitem_type_Pdf_Radiobutton_Field);
    
    /**
	 * Updates all field with this field name.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.UpdateAll = function() {
        let aOptions    = this.GetOptions();
        let aFields     = this.GetAllWidgets();
        let value       = this.GetLogicValue();

        let bFromOpt = false;
        let sExportValue;
        
        // if there's opt array, then value is 99% of the time an index in opt array (export values)
        if (aOptions && aOptions[value]) {
            sExportValue = aOptions[value];
            bFromOpt = true;
        }
        else {
            sExportValue = value;
        }

        if (this.IsRadiosInUnison()) {
            aFields.forEach(function(field) {
                if (field.GetExportValue() == sExportValue) {
                    field.SetChecked(true);
                }
                else {
                    field.SetChecked(false);
                }
            });
        }
        else {
            let oFirstField = null;
            aFields.forEach(function(field, index) {
                // set by index from Opt
                if (bFromOpt) {
                    if (index == value) {
                        field.SetChecked(true);
                    }
                    else {
                        field.SetChecked(false);
                    }
                }
                else {
                    // keep only first one checked if not by index from Opt
                    if (oFirstField == null && field.GetExportValue() == sExportValue) {
                        field.SetChecked(true);
                        oFirstField = field;
                    }
                    else {
                        field.SetChecked(false);
                    }
                }
                
            })
        }
    };
    
    /**
	 * Applies value of this field to all field with the same name
     * Note: Uses after mouseUp action.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.Commit2 = function() {
        let aFields = this.GetAllWidgets();
        let oThis = this;

        if (false == this.IsRadiosInUnison()) {
            aFields.forEach(function(field) {
                if (field == oThis)
                    return;

                if (field.IsChecked() == true) {
                    field.SetChecked(false);
                    field.SetNeedRecalc(true);
                }
            }); 
        }
        else {
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

            if (AscCommon.History.Is_LastPointEmpty())
                AscCommon.History.Remove_LastPoint();
        }
    };
    
    CRadioButtonField.prototype.SetRadiosInUnison = function(bValue) {
        let oParent = this.GetParent(true);
        if (oParent)
            return oParent.SetRadiosInUnison(bValue);

        if (this._radiosInUnison === bValue) {
            return true;
        }
        
        AscCommon.History.Add(new CChangesPDFRadiobuttonIsUnison(this, this._radiosInUnison, bValue));

        this._radiosInUnison = bValue;
        this.SetWasChanged(true);

        return true;
    };
    CRadioButtonField.prototype.IsRadiosInUnison = function(bInherit) {
        let oParent = this.GetParent(true);
        if (bInherit !== false && oParent)
            return oParent.IsRadiosInUnison();

        return this._radiosInUnison;
    };

	
    if (!window["AscPDF"])
	    window["AscPDF"] = {};
        
	window["AscPDF"].CRadioButtonField = CRadioButtonField;
})();

