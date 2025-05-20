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
        this._currentValueIndices   = [];
        this._textFont              = AscPDF.DEFAULT_FIELD_FONT;
        this._options               = [];
    }
    CBaseListField.prototype = Object.create(AscPDF.CBaseField.prototype);
	CBaseListField.prototype.constructor = CBaseListField;

    CBaseListField.prototype.SetParentCurIdxs = function(aIdxs) {
        let oParent = this.GetParent();
        if (oParent && this.IsWidget() && oParent.IsAllKidsWidgets())
            oParent.SetParentCurIdxs(aIdxs);
        else {
            let oDoc = this.GetDocument();
            oDoc.History.Add(new CChangesPDFListFormParentCurIdxs(this, this.GetParentCurIdxs(), aIdxs));
            this._currentValueIndices = aIdxs;
        }
    };
    CBaseListField.prototype.GetParentCurIdxs = function(bInherit) {
        let oParent = this.GetParent();
        if (oParent == null)
            return this._currentValueIndices;
        else if (bInherit === false || (this.GetPartialName() != null)) {
            return this._currentValueIndices;
        }
        
        if (oParent)
            return oParent.GetParentCurIdxs();
    };
    CBaseListField.prototype.SetTopIndex = function() {};
    CBaseListField.prototype.SetCommitOnSelChange = function(bValue) {
        let oParent = this.GetParent();
        if (oParent && oParent.IsAllKidsWidgets()) {
            oParent.SetCommitOnSelChange(bValue);
            return;
        }

        let oDoc = this.GetDocument();
        oDoc.History.Add(new CChangesPDFListCommitOnSelChange(this, this._commitOnSelChange, bValue));

        this._commitOnSelChange = bValue;
        this.SetWasChanged(true);
    };
    CBaseListField.prototype.IsCommitOnSelChange = function(bInherit) {
        let oParent = this.GetParent();
        if (bInherit !== false && oParent && oParent.IsAllKidsWidgets())
            return oParent.IsCommitOnSelChange();

        return this._commitOnSelChange;
    };

    CBaseListField.prototype.GetOptions = function(bInherit) {
        let oParent = this.GetParent();
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

