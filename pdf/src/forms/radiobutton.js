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
	 * Class representing a radiobutton field.
	 * @constructor
     * @extends {CBaseCheckBoxField}
	 */
    function CRadioButtonField(sName, nPage, aRect, oDoc)
    {
        AscPDF.CBaseCheckBoxField.call(this, sName, AscPDF.FIELD_TYPES.radiobutton, nPage, aRect, oDoc);
        
        this._radiosInUnison = false;
        this._noToggleToOff = true;

        this._chStyle       = AscPDF.CHECKBOX_STYLES.circle;
    }
    CRadioButtonField.prototype = Object.create(AscPDF.CBaseCheckBoxField.prototype);
	CRadioButtonField.prototype.constructor = CRadioButtonField;
    
    /**
	 * Synchronizes this field with fields with the same name.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.SyncField = function() {
        // to do
    };
    
    /**
	 * Updates all field with this field name.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.UpdateAll = function() {
        let oParent     = this.GetParent();
        let aParentOpt  = oParent ? oParent.GetOptions() : undefined;
        let aFields     = this.GetDocument().GetAllWidgets(this.GetFullName());
        let value       = this.GetParentValue();

        let bFromOpt = false;
        let sExportValue;
        
        // если есть массив опт, то value в 99% случаев - это индекс в массиве opt (export values)
        if (aParentOpt && aParentOpt[value]) {
            sExportValue = aParentOpt[value];
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
                // выставляем по индексу из Opt
                if (bFromOpt) {
                    if (index == value) {
                        field.SetChecked(true);
                    }
                    else {
                        field.SetChecked(false);
                    }
                }
                else {
                    // оставляем отмеченным только первую если не по индексу из Opt
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
        let aFields = this.GetDocument().GetAllWidgets(this.GetFullName());
        let oThis = this;

        if (false == this.IsRadiosInUnison()) {
            aFields.forEach(function(field) {
                if (field == oThis)
                    return;

                if (field.IsChecked() == true && oThis.IsChecked()) {
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
        this._radiosInUnison = bValue;
        this.SetWasChanged(true);
    };
    CRadioButtonField.prototype.IsRadiosInUnison = function() {
        return this._radiosInUnison;
    };

	
    if (!window["AscPDF"])
	    window["AscPDF"] = {};
        
	window["AscPDF"].CRadioButtonField = CRadioButtonField;
})();

