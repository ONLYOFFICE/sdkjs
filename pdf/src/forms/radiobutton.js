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
        let aFields = this._doc.GetFields(this.GetFullName());
        
        if (this._radiosInUnison) {
            // отмечаем все radiobuttons с тем же экспортом, что и отмеченные
            let sExportValue;
            for (let i = 0; i < aFields.length; i++) {
                if (!sExportValue && aFields[i]._value != "Off") {
                    sExportValue = aFields[i]._exportValue;
                    break;
                }
            }
            if (!sExportValue) {
                aFields.forEach(function(field) {
                    field.SetChecked(false);
                });
            }
            else {
                aFields.forEach(function(field) {
                    if (field._exportValue != sExportValue) {
                        field.SetChecked(false);
                    }
                    else {
                        field.SetChecked(true);
                    }
                });
            }
        }
        else {
            let oCheckedFld = null;
            // оставляем активной первую отмеченную radiobutton
            for (let i = 0; i < aFields.length; i++) {
                if (!oCheckedFld && aFields[i]._value != "Off") {
                    oCheckedFld = aFields[i];
                    continue;
                }
                if (oCheckedFld) {
                    aFields[i].SetChecked(false);
                }
            }
        }
    };
    
    /**
	 * The value application logic for all fields with the same name has been changed for this field type.
     * The method was left for compatibility.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.Commit = function() {
        this.SetNeedCommit(false);
    };

    /**
	 * Applies value of this field to all field with the same name
     * Note: Uses after mouseUp action.
	 * @memberof CRadioButtonField
	 * @typeofeditors ["PDF"]
	 */
    CRadioButtonField.prototype.Commit2 = function() {
        let aFields = this.GetDocument().GetFields(this.GetFullName());
        let oThis = this;

        if (false == this._radiosInUnison) {
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

                if (field._exportValue != oThis._exportValue && field.IsChecked() == true) {
                    field.SetChecked(false);
                    field.SetNeedRecalc(true);
                }
                else if (field._exportValue == oThis._exportValue && oThis.IsChecked() == false) {
                    field.SetChecked(false);
                    field.SetNeedRecalc(true);
                }
                else if (field._exportValue == oThis._exportValue && field.IsChecked() == false) {
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
    };
	
    if (!window["AscPDF"])
	    window["AscPDF"] = {};
        
	window["AscPDF"].CRadioButtonField = CRadioButtonField;
})();

