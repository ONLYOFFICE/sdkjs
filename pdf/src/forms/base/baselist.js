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
	 * A string that sets the trigger for the action. Values are:
	 * @typedef {"MouseUp" | "MouseDown" | "MouseEnter" | "MouseExit" | "OnFocus" | "OnBlur" | "Keystroke" | "Validate" | "Calculate" | "Format"} cTrigger
	 * For a list box, use the Keystroke trigger for the Selection Change event.
     */
    
    /**
	 * Class representing a base listbox/combobox field.
	 * @constructor
     * @extends {CBaseField}
	 */
    function CBaseListField(sName, sType, nPage, aRect, oDoc)
    {
        AscPDF.CBaseField.call(this, sName, sType, nPage, aRect, oDoc);

        this._commitOnSelChange     = false;
        this._currentValueIndexes   = [];
        this._textFont              = AscPDF.DEFAULT_FIELD_FONT;
        this._options               = [];
    }
    CBaseListField.prototype = Object.create(AscPDF.CBaseField.prototype);
	CBaseListField.prototype.constructor = CBaseListField;

    CBaseListField.prototype.AddKid = function(oField) {
        oField.SetOptions([]);
        AscPDF.CBaseField.prototype.AddKid.call(this, oField);
    };
    CBaseListField.prototype.SetLogicCurIdxs = function(aIdxs) {
        let oParent = this.GetParent(true);
        if (oParent && this.IsWidget())
            oParent.SetLogicCurIdxs(aIdxs);
        else {
            AscCommon.History.Add(new CChangesPDFListFormParentCurIdxs(this, this.GetLogicCurIdxs(), aIdxs));
            this._currentValueIndexes = aIdxs;
        }
    };
    CBaseListField.prototype.GetLogicCurIdxs = function(bInherit) {
        let oParent = this.GetParent(true);
        if (oParent == null)
            return this._currentValueIndexes;
        else if (bInherit === false || (this.GetPartialName() != null)) {
            return this._currentValueIndexes;
        }
        
        if (oParent)
            return oParent.GetLogicCurIdxs();
    };
    CBaseListField.prototype.SetTopIndex = function() {};
    CBaseListField.prototype.SetCommitOnSelChange = function(bValue) {
        let oParent = this.GetParent(true);
        if (oParent) {
            return oParent.SetCommitOnSelChange(bValue);
        }

        AscCommon.History.Add(new CChangesPDFListCommitOnSelChange(this, this._commitOnSelChange, bValue));

        this._commitOnSelChange = bValue;
        this.SetWasChanged(true);

        return true;
    };
    CBaseListField.prototype.IsCommitOnSelChange = function(bInherit) {
        let oParent = this.GetParent(true);
        if (bInherit !== false && oParent)
            return oParent.IsCommitOnSelChange();

        return this._commitOnSelChange;
    };

    CBaseListField.prototype.GetOptions = function(bInherit) {
        let oParent = this.GetParent(true);
        if (oParent == null)
            return this._options;
        else if (bInherit === false || (this.GetPartialName() != null)) {
            return this._options;
        }
        
        if (oParent)
            return oParent.GetOptions();
    };
	// export
	CBaseListField.prototype["getOptions"] = function() {
		return this.GetOptions();
	};

    if (!window["AscPDF"])
	    window["AscPDF"] = {};
        
	window["AscPDF"].CBaseListField = CBaseListField;
})();

