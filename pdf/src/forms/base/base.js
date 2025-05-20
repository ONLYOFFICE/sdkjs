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

    //------------------------------------------------------------------------------------------------------------------
	//
	// Internal
	//
	//------------------------------------------------------------------------------------------------------------------

    let FIELDS_HIGHLIGHT = {
        r: 221,
        g: 228,
        b: 255
    };
    let FIELDS_HIGHLIGHT_REQ = {
        r: 176,
        g: 176,
        b: 255
    };

    //------------------------------------------------------------------------------------------------------------------
	//
	// pdf api types
	//
	//------------------------------------------------------------------------------------------------------------------
    
    let ALIGN_TYPE = {
        left:   0,
        center: 1,
        right:  2
    };

    // For Span attributes (start)
    let FONT_STRETCH = ["ultra-condensed", "extra-condensed", "condensed", "semi-condensed", "normal",
        "semi-expanded", "expanded", "extra-expanded", "ultra-expanded"];

    let FONT_STYLE = {
        italic: "italic",
        normal: "normal"
    }

    let FONT_WEIGHT = [100, 200, 300, 400, 500, 600, 700, 800, 900];
    let VALID_ROTATIONS = [0, 90, 180, 270];

    const MAX_TEXT_SIZE = 32767;
	
	const DEFAULT_FIELD_FONT = "Arial";
	
	// freeze objects
    Object.freeze(FIELDS_HIGHLIGHT);
    Object.freeze(ALIGN_TYPE);
    Object.freeze(FONT_STRETCH);
    Object.freeze(FONT_STYLE);
    Object.freeze(FONT_WEIGHT);
    Object.freeze(VALID_ROTATIONS);

    
    /**
	 * Class representing a base field class.
	 * @constructor
    */
    function CBaseField(sName, nType, aRect)
    {
        this.Id = AscCommon.g_oIdCounter.Get_NewId();
        if ((AscCommon.g_oIdCounter.m_bLoad || AscCommon.History.CanAddChanges())) {
            AscCommon.g_oTableId.Add(this, this.Id);
        }

        this.type           = nType;
        this._partialName   = undefined;
        this._kids          = [];
        this._apIdx         = undefined; // индекс формы на странице в исходном файле (в массиве метода getInteractiveForms), используется для получения appearance
        this._borderStyle   = undefined;
        this._lineWidth     = undefined;        // In older versions of this specification, this property was borderWidth
        this._borderWidth   = undefined;       
        this._strokeColor   = undefined;     // In older versions of this specification, this property was borderColor. The use of borderColor is now discouraged, although it is still valid for backward compatibility.
        this._borderColor   = undefined;
        this._fillColor     = undefined;
        this._bgColor       = undefined;          // prop for old versions (fillColor)
        this._fgColor       = undefined;
        this._textColor     = [0,0,0];
        this._textFont      = "";           // исходный
        this._textFontActual= undefined;    // фактический используемый
        this._textSize      = 10;
        this._fontStyle     = 0;    // информация о стиле шрифта (bold, italic)

        this._display       = AscPDF.Api.Types.display["visible"];
        this._hidden        = false;             // This property has been superseded by the display property and its use is discouraged.
        this._print         = true;        // This property has been superseded by the display property and its use is discouraged.
        this._readOnly      = false;
        this._required      = false;       // for all except button
        
        this._delay         = false;
        this._noExport      = false;
        this._defaultValue  = undefined;
        this._parent        = null;
        this._rotation      = 0;
        this._submitName    = "";
        this._userName      = "";   // It is intended to be used as tooltip text whenever the cursor enters a field. 
        //It can also be used as a user-friendly name, instead of the field name, when generating error messages.

        // internal
        this._triggers = new AscPDF.CFormTriggers();
        this._curShiftView = { // смещение, когда мы скролим, т.е. активное смещение
            x: 0,
            y: 0
        }
        this._oldShiftView = { // смещение, когда значение формы применено (т.е. форма не активна)
            x: 0,
            y: 0
        }
        this.needCommit             = false;
        this._needDrawHighlight     = true;
        this._needDrawHoverBorder   = false;
        this._needRecalc            = true;
        this._wasChanged            = false; // была ли изменена форма
        this._bDrawFromStream       = false; // нужно ли рисовать из стрима
        this._hasOriginView         = false; // имеет ли внешний вид из файла
        this._originView = {
            normal:     null,
            mouseDown:  null,
            rollover:   null
        }
        this.api = this.GetFormApi();
        this["api"] = this.api;
		this.compositeInput = null;
		this.compositeReplaceCount = 0;
        this.Lock = new AscCommon.CLock();
        this.meta = {};
        this.SetPartialName(sName);
        this.SetRect(aRect);

        this.kidsContentChanges = new AscCommon.CContentChanges();
    }

    CBaseField.prototype.constructor = CBaseField;
    CBaseField.prototype = Object.create(AscFormat.CBaseNoIdObject.prototype);

    CBaseField.prototype.IsAnnot = function() {
        return false;
    };
    CBaseField.prototype.IsForm = function() {
        return true;
    };
    CBaseField.prototype.IsDrawing = function() {
        return false;
    };
    CBaseField.prototype.IsPdfObject = function() {
        return true;
    };
    CBaseField.prototype.SetApIdx = function(nIdx) {
        this._apIdx = nIdx;
    };
    CBaseField.prototype.GetApIdx = function() {
        if (undefined == this._apIdx) {
            if (undefined == this.GetId()) {
                return -1;
            }
            else {
                let nApIdx = Number(this.GetId().replace("_", ""));
                if (!isNaN(nApIdx)) {
                    return nApIdx;
                }
            }
        }

        return this._apIdx;
    };

    CBaseField.prototype.SetMeta = function(oMeta) {
        AscCommon.History.Add(new CChangesPDFFormMeta(this, this.meta, oMeta))

        this.meta = oMeta;
    };
    CBaseField.prototype.GetMeta = function() {
        return this.meta;
    };
    CBaseField.prototype.AddToChildsMap = function(nIdx) {
        this.GetDocument().AddFieldToChildsMap(this, nIdx);
    };
    /**
	 * Can or not enter text into form.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 */
    CBaseField.prototype.IsCanEditText = function() {
        return false;
    };
	/**
	 * Check if fonts are available, if not then we download them
	 * @returns {boolean}
	 */
	CBaseField.prototype.checkFonts = function() {
		return this._doc && 1 === this._doc.defaultFontsLoaded;
	};
	CBaseField.prototype.getAllRasterImages = function(images) {};
    /**
	 * Invokes only on open forms.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 */
    CBaseField.prototype.SetOriginPage = function(nPage) {
        this._origPage = nPage;
    };
    CBaseField.prototype.GetOriginPage = function() {
        return this._origPage;
    };

    /**
	 * Gets the child field by the specified partial name.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 * @returns {?CBaseField}
	 */
    CBaseField.prototype.GetField = function(sName) {
        for (let i = 0; i < this._kids.length; i++) {
            if (this._kids[i]._partialName == sName)
                return this._kids[i];
        }

        return null;
    };
    CBaseField.prototype.SetNeedCommit = function(bValue) {
        if (editor.getDocumentRenderer().IsOpenFormsInProgress)
            return;

        this.needCommit = bValue;
    };
    CBaseField.prototype.IsNeedCommit = function() {
        return this.needCommit;
    };
    CBaseField.prototype.SetPage = function(nPage) {
        if (this.GetPage() == nPage) {
            return;
        }
        
        let oDoc        = this.GetDocument();
        let oCurPage    = this.GetParentPage();
        let oNewPage    = oDoc.GetPageInfo(nPage);
        
        if (oNewPage) {
            let sId = this.GetId();
            oCurPage.RemoveField(sId);
            oNewPage.AddField(this);
        }
    };
    CBaseField.prototype.GetPage = function() {
        let oParentPage = this.GetParentPage();
        if (!oParentPage || !(oParentPage instanceof AscPDF.CPageInfo)) {
            return -1;
        }

        return oParentPage.GetIndex();
    };

    /**
	 * Gets the child field by the specified partial name.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 * @returns {?CBaseField}
	 */
    CBaseField.prototype.AddKid = function(oField) {
        AscCommon.History.Add(new CChangesPDFFormKidsContent(this, this._kids.length, [oField], true))

        this._kids.push(oField);
        oField._parent = this;
    };
    CBaseField.prototype.GetKids = function() {
        return this._kids;
    };
    CBaseField.prototype.GetKid = function(nPos, bStrict) {
        if (this.IsWidget() && bStrict != true)
            return this;

        return this._kids[nPos];
    };
    /**
	 * Removes field from kids.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
     * @param {CBaseField} oField - the field to remove.
	 * @returns {boolean} - returns false if field isn't in the field kids.
	 */
    CBaseField.prototype.RemoveKid = function(oField) {
        let nIndex = this._kids.indexOf(oField);
        if (nIndex != -1) {
            this._kids.splice(nIndex, 1);
            AscCommon.History.Add(new CChangesPDFFormKidsContent(this, nIndex, [oField], false))
            oField._parent = null;
            return true;
        }

        return false;
    };
    
    /**
	 * Gets all widgets fields of this parent field.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 * @returns {Array}
	 */
    CBaseField.prototype.GetAllWidgets = function() {
        let oParent = this.GetParent();
        if (oParent && oParent.IsAllKidsWidgets()) {
            return oParent.GetAllWidgets()
        }

        if (this.IsWidget())
            return [this];

        let aWidgets    = [];
        let aKids       = this.GetKids();
        for (let i = 0; i < aKids.length; i++) {
            if (aKids[i].IsWidget()) {
                aWidgets.push(aKids[i]);
            }
            else
                aWidgets = aWidgets.concat(aKids[i].GetAllWidgets());
        }

        return aWidgets;
    };

    CBaseField.prototype.Recalculate = function() {};
    
    CBaseField.prototype.GetDocContent = function(bFormatContent) {
        return bFormatContent ? this.contentFormat : this.content;
    };

    CBaseField.prototype.getFormRelRect = function() {
        return this.contentClipRect;
    };
    CBaseField.prototype.getFormRect = function() {
        let aOrigRect = this.GetOrigRect();

        return {
            X : aOrigRect[0] * g_dKoef_pt_to_mm,
            Y : aOrigRect[1] * g_dKoef_pt_to_mm,
            W : (aOrigRect[2] - aOrigRect[0]) * g_dKoef_pt_to_mm,
            H : (aOrigRect[3] - aOrigRect[1]) * g_dKoef_pt_to_mm
        }
    };
    
    CBaseField.prototype.GetFullName = function() {
        if (this._parent)
        {
            if (this._partialName != undefined)
                return this._parent.GetFullName() + "." + this._partialName;
            else
                return this._parent.GetFullName();
        }

        return this._partialName ? this._partialName : "";
    };
    CBaseField.prototype.SetPartialName = function(sName) {
        this._partialName = sName;
        AscCommon.History.Add(new CChangesPDFFormPartialName(this, this._partialName, sName));

        this.SetWasChanged(true);

        if (this.IsEditMode()) {
            AscCommon.History.StartNoHistoryMode();
            let oShape = this.GetEditShape();
            let oContent = oShape.GetDocContent();
            let oPara = oContent.GetElement(0);
            let oRun = oPara.GetElement(0);
            oRun.Remove_FromContent(0, oRun.Content.length);
            oRun.AddText(this.GetFullName());
            oShape.recalcContent();
            oShape.recalculate();
            this.AddToRedraw();
            AscCommon.History.EndNoHistoryMode();
        }
    };
    CBaseField.prototype.GetPartialName = function() {
        return this._partialName;
    };
    CBaseField.prototype.SyncValue = function() {};
    /**
	 * Sets the action of the field for a given trigger.
     * Note: This method will overwrite any action already defined for the chosen trigger.
	 * @memberof CBaseField
     * @param {number} nTriggerType - A string that sets the trigger for the action. (FORMS_TRIGGERS_TYPES)
	 * @param {Array} aActionsInfo - array with actions info for specified trigger. (info from openForms method)
     * @typeofeditors ["PDF"]
	 */
    CBaseField.prototype.SetActions = function(nTriggerType, aActionsInfo) {
        switch (nTriggerType) {
            case AscPDF.FORMS_TRIGGERS_TYPES.Keystroke:
            case AscPDF.FORMS_TRIGGERS_TYPES.Validate:
            case AscPDF.FORMS_TRIGGERS_TYPES.Calculate:
            case AscPDF.FORMS_TRIGGERS_TYPES.Format: {
                let oParent = this.GetParent();
                if (oParent && oParent.IsAllKidsWidgets())
                    return oParent.SetActions(nTriggerType, aActionsInfo);
            }
        }

        let aActions = [];
        if (aActionsInfo) {
            for (let i = 0; i < aActionsInfo.length; i++) {
                let oAction;
                switch (aActionsInfo[i]["S"]) {
                    case AscPDF.ACTIONS_TYPES.JavaScript:
                        oAction = new AscPDF.CActionRunScript(aActionsInfo[i]["JS"]);
                        aActions.push(oAction);
                        break;
                    case AscPDF.ACTIONS_TYPES.ResetForm:
                        oAction = new AscPDF.CActionReset(aActionsInfo[i]["Fields"], Boolean(aActionsInfo[i]["Flags"]));
                        aActions.push(oAction);
                        break;
                    case AscPDF.ACTIONS_TYPES.URI:
                        oAction = new AscPDF.CActionURI(aActionsInfo[i]["URI"]);
                        aActions.push(oAction);
                        break;
                    case AscPDF.ACTIONS_TYPES.HideShow:
                        oAction = new AscPDF.CActionHideShow(Boolean(aActionsInfo[i]["H"]), aActionsInfo[i]["T"]);
                        aActions.push(oAction);
                        break;
                    case AscPDF.ACTIONS_TYPES.GoTo:
                        let oRect = {
                            top:    aActionsInfo[i]["top"],
                            right:  aActionsInfo[i]["right"],
                            bottom: aActionsInfo[i]["bottom"],
                            left:   aActionsInfo[i]["left"]
                        }
                        if (aActionsInfo[i]["bottom"] != null && aActionsInfo[i]["top"] != null) {
                            oRect.top = aActionsInfo[i]["bottom"];
                            oRect.bottom = aActionsInfo[i]["top"];
                        }
    
                        oAction = new AscPDF.CActionGoTo(aActionsInfo[i]["page"], aActionsInfo[i]["kind"], aActionsInfo[i]["zoom"], oRect);
                        aActions.push(oAction);
                        break;
                    case AscPDF.ACTIONS_TYPES.Named:
                        oAction = new AscPDF.CActionNamed(AscPDF.CActionNamed.GetInternalType(aActionsInfo[i]["N"]));
                        aActions.push(oAction);
                        break;
                }
            }
        }
        
        const oNewTrigger = aActions.length != 0 ? new AscPDF.CFormTrigger(nTriggerType, aActions) : null;
        if (oNewTrigger) {
            oNewTrigger.SetParentField(this);
        }

        const aCurActionsInfo = this.GetActions();
        AscCommon.History.Add(new CChangesPDFFormActions(this, aCurActionsInfo, aActionsInfo, nTriggerType));

        switch (nTriggerType) {
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseUp:
                this._triggers.MouseUp = oNewTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseDown:
                this._triggers.MouseDown = oNewTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseEnter:
                this._triggers.MouseEnter = oNewTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseExit:
                this._triggers.MouseExit = oNewTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.OnFocus:
                this._triggers.OnFocus = oNewTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.OnBlur:
                this._triggers.OnBlur = oNewTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Keystroke:
                this._triggers.Keystroke = oNewTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Validate:
                this._triggers.Validate = oNewTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Calculate:
                this._triggers.Calculate = oNewTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Format:
                this._triggers.Format = oNewTrigger;
                break;
        }

        return aActions;
    };
    CBaseField.prototype.GetActions = function(nTriggerType) {
        switch (nTriggerType) {
            case AscPDF.FORMS_TRIGGERS_TYPES.Keystroke:
            case AscPDF.FORMS_TRIGGERS_TYPES.Validate:
            case AscPDF.FORMS_TRIGGERS_TYPES.Calculate:
            case AscPDF.FORMS_TRIGGERS_TYPES.Format: {
                let oParent = this.GetParent();
                if (oParent && oParent.IsAllKidsWidgets())
                    return oParent.GetActions(nTriggerType);
            }
        }

        // Get the trigger by type
        let oTrigger = this.GetTrigger(nTriggerType);
        if (!oTrigger || !oTrigger.Actions) {
            return [];
        }
        
        let aActionsInfo = [];
        // Iterate through all actions associated with the trigger
        for (let i = 0; i < oTrigger.Actions.length; i++) {
            let oAction = oTrigger.Actions[i];
            let actionInfo = {};
            
            // Determine the action type and populate the object with information
            switch (oAction.GetType()) {
                case AscPDF.ACTIONS_TYPES.JavaScript:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.JavaScript;
                    actionInfo["JS"] = oAction.GetScript();
                    break;
                case AscPDF.ACTIONS_TYPES.ResetForm:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.ResetForm;
                    actionInfo["Fields"] = oAction.GetNames();
                    actionInfo["Flags"] = Number(oAction.GetNeedAllExcept());
                    break;
                case AscPDF.ACTIONS_TYPES.URI:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.URI;
                    actionInfo["URI"] = oAction.GetURI();
                    break;
                case AscPDF.ACTIONS_TYPES.HideShow:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.HideShow;
                    actionInfo["H"] = oAction.GetHidden();
                    actionInfo["T"] = oAction.GetNames();
                    break;
                case AscPDF.ACTIONS_TYPES.GoTo:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.GoTo;
                    actionInfo["page"] = oAction.GetPage();
                    actionInfo["kind"] = oAction.GetKind();
                    actionInfo["zoom"] = oAction.GetZoom();
                    let oRect = oAction.GetRect();
                    actionInfo["top"] = oRect.top;
                    actionInfo["right"] = oRect.right;
                    actionInfo["bottom"] = oRect.bottom;
                    actionInfo["left"] = oRect.left;
                    break;
                case AscPDF.ACTIONS_TYPES.Named:
                    actionInfo["S"] = AscPDF.ACTIONS_TYPES.Named;
                    actionInfo["N"] = oAction.GetName();
                    break;
                default:
                    // If the type is not recognized, add handling or skip
                    break;
            }
            
            aActionsInfo.push(actionInfo);
        }
        
        return aActionsInfo;
    };


    /**
	 * Sets the JavaScript action of the field for a given trigger.
     * Note: This method will overwrite any action already defined for the chosen trigger.
	 * @memberof CBaseField
     * @param {number} nTriggerType - A string that sets the trigger for the action.
     * @param {string} sScript - The JavaScript code to be executed when the trigger is activated.
	 * @typeofeditors ["PDF"]
	 */
    CBaseField.prototype.SetAction = function(nTriggerType, sScript) {
        let oDoc        = this.GetDocument();
        let oCalcInfo   = oDoc.GetCalculateInfo();
        let oAction     = new AscPDF.CActionRunScript(sScript);

        let oTrigger = new AscPDF.CFormTrigger(nTriggerType, [oAction]);
        oTrigger.SetParentField(this);

        switch (nTriggerType) {
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseUp:
                this._triggers.MouseUp = oTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseDown:
                this._triggers.MouseDown = oTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseEnter:
                this._triggers.MouseEnter = oTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseExit:
                this._triggers.MouseExit = oTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.OnFocus:
                this._triggers.OnFocus = oTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.OnBlur:
                this._triggers.OnBlur = oTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Keystroke:
                this._triggers.Keystroke = oTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Validate:
                this._triggers.Validate = oTrigger;
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Calculate:
                this._triggers.Calculate = oTrigger;
                oCalcInfo.RemoveFieldFromOrder(this.GetFullName());
                oCalcInfo.AddFieldToOrder(oDoc.GetField(this.GetFullName()).GetApIdx());
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Format:
                this._triggers.Format = oTrigger;
                break;
        }
    };

    /**
	 * Sets a flag that we have entered the field.
     * This is not the same as an doc.activeField.
	 * @memberof CBaseField
     * @param {boolean} bInField
	 */
    CBaseField.prototype.SetInForm = function(bInField) {
        this.isInForm = bInField;
    };
    CBaseField.prototype.IsInForm = function() {
        return !!this.isInForm;
    };
    
    /**
	 * Gets the JavaScript action of the field for a given trigger.
	 * @memberof CBaseField
     * @param {number} nType - A string that sets the trigger for the action. (FORMS_TRIGGERS_TYPES)
	 * @typeofeditors ["PDF"]
     * @returns {CFormTrigger}
	 */
    CBaseField.prototype.GetTrigger = function(nType, bInherit) {
        switch (nType) {
            case AscPDF.FORMS_TRIGGERS_TYPES.Keystroke:
            case AscPDF.FORMS_TRIGGERS_TYPES.Validate:
            case AscPDF.FORMS_TRIGGERS_TYPES.Calculate:
            case AscPDF.FORMS_TRIGGERS_TYPES.Format: {
                let oParent = this.GetParent();
                if (bInherit !== false && oParent && oParent.IsAllKidsWidgets())
                    return oParent.GetTrigger(nType);
            }
        }

        switch (nType) {
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseUp:
                return this._triggers.MouseUp;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseDown:
                return this._triggers.MouseDown;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseEnter:
                return this._triggers.MouseEnter;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseExit:
                return this._triggers.MouseExit;
            case AscPDF.FORMS_TRIGGERS_TYPES.OnFocus:
                return this._triggers.OnFocus;
            case AscPDF.FORMS_TRIGGERS_TYPES.OnBlur:
                return this._triggers.OnBlur;
            case AscPDF.FORMS_TRIGGERS_TYPES.Keystroke:
                return this._triggers.Keystroke;
            case AscPDF.FORMS_TRIGGERS_TYPES.Validate:
                return this._triggers.Validate;
            case AscPDF.FORMS_TRIGGERS_TYPES.Calculate:
                return this._triggers.Calculate;
            case AscPDF.FORMS_TRIGGERS_TYPES.Format:
                return this._triggers.Format;
        }

        return null;
    };
    CBaseField.prototype.GetListActions = function(bInherit) {
        let aActions = [];

        let oAction = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.MouseUp);
        if (oAction) {
            aActions.push(oAction);
        }
        
        oAction = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.MouseDown);
        if (oAction) {
            aActions.push(oAction);
        }

        oAction = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.MouseEnter);
        if (oAction) {
            aActions.push(oAction);
        }

        oAction = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.MouseExit);
        if (oAction) {
            aActions.push(oAction);
        }

        oAction = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.OnFocus);
        if (oAction) {
            aActions.push(oAction);
        }

        oAction = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.OnBlur);
        if (oAction) {
            aActions.push(oAction);
        }

        oAction = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Keystroke, bInherit);
        if (oAction) {
            aActions.push(oAction);
        }

        oAction = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Validate, bInherit);
        if (oAction) {
            aActions.push(oAction);
        }

        oAction = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Calculate, bInherit);
        if (oAction) {
            aActions.push(oAction);
        }

        oAction = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Format, bInherit);
        if (oAction) {
            aActions.push(oAction);
        }
        
        return aActions;
    };
    CBaseField.prototype.SetDocument = function(oDoc) {
        if (this._doc == oDoc) {
            return;
        }

        AscCommon.History.Add(new CChangesPDFObjectSetDocument(this, this._doc, oDoc));
        this._doc = oDoc;
    };
    CBaseField.prototype.GetDocument = function() {
        return this._doc;
    };

    CBaseField.prototype.GetParent = function() {
        return this._parent;
    };

    /**
	 * Gets form value (which was commited).
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 */
    CBaseField.prototype.GetParentValue = function(bInherit) {
        let oParent = this.GetParent();
        if (oParent == null && this._value == null)
            return undefined;
        else if (bInherit === false || (this._value != null && this.GetPartialName() != null)) {
            return this._value;
        }
        
        if (oParent)
            return oParent.GetParentValue();
    };
    /**
	 * Sets api value of form.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 */
    CBaseField.prototype.SetParentValue = function(value) {
        let oParent = this.GetParent();
        if (oParent && this.IsWidget() && oParent.IsAllKidsWidgets())
            oParent.SetParentValue(value);
        else {
            if (this._value === value) {
                return true;
            }
            
            this.SetWasChanged(true);
            let oDoc = this.GetDocument();
            oDoc.History.Add(new CChangesPDFFormParentValue(this, this._value, value));
            this._value = value;
        }

        return true;
    };

    /**
	 * Checks if all kids are widgets (and have same names).
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 */
    CBaseField.prototype.IsAllKidsWidgets = function() {
        let aKids = this.GetKids();

        if (aKids.length > 0) {
            if (aKids[0].IsWidget() == false)
                return false;

            let sFullName = aKids[0].GetFullName();
            for (let i = 1; i < aKids.length; i++) {
                if (sFullName != aKids[i].GetFullName() || aKids[i].IsWidget() == false)
                    return false;
            }

            return true;
        }

        return false;
    };
    /**
     * Does the actions setted for specifed trigger type.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
     * @returns {canvas}
	 */
    CBaseField.prototype.AddActionsToQueue = function() {
        let oThis           = this;
        let oDoc            = this.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        Object.values(arguments).forEach(function(type) {
            let oTrigger = oThis.GetTrigger(type);
        
            if (oTrigger && oTrigger.Actions.length > 0 && false == AscCommon.History.UndoRedoInProgress) {
                oActionsQueue.AddActions(oTrigger.Actions);
            }
        })
        
        if (oActionsQueue.actions.length !== 0) {
            oActionsQueue.Start();
        }
    };
    CBaseField.prototype.IsUseInDocument = function() {
        let oPage = this.GetParentPage();
        if (oPage && oPage.fields.includes(this)) {
            return true;
        }

        return false;
    };
    CBaseField.prototype.DrawHighlight = function(oCtx) {
        if (this.IsHidden() && !this.IsEditMode())
            return;

        let oViewer     = Asc.editor.getDocumentRenderer();
        let nPage       = this.GetPage();
        let nScale      = AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom * oViewer.getDrawingPageScale(nPage);
        let aBgColor    = this.GetBackgroundColor();

        let xCenter = oViewer.width >> 1;
        if (oViewer.documentWidth > oViewer.width) {
            xCenter = (oViewer.documentWidth >> 1) - (oViewer.scrollX) >> 0;
        }
        let yPos    = oViewer.scrollY >> 0;
        let page    = oViewer.drawingPages[nPage];
        let w       = (page.W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
        let h       = (page.H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
        let indLeft = ((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - (w >> 1);
        let indTop  = ((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
        
        let isLandscape = oViewer.isLandscapePage(nPage);
        if (isLandscape) {
            indLeft = indLeft + (w - h) / 2;
        }

        let aOrigRect = this.GetOrigRect();
        let X = aOrigRect[0] * nScale + indLeft;
        let Y = aOrigRect[1] * nScale + indTop;
        let W = (aOrigRect[2] - aOrigRect[0]) * nScale;
        let H = (aOrigRect[3] - aOrigRect[1]) * nScale;

        if (null == aBgColor)
            oCtx.globalAlpha = 0.8;
        
        oCtx.globalCompositeOperation = "destination-over";
        
        if (this.IsNeedDrawFromStream())
            AscPDF.startMultiplyMode(oCtx);
        
        if (this.GetType() == AscPDF.FIELD_TYPES.radiobutton && this._chStyle == AscPDF.CHECKBOX_STYLES.circle) {
            oCtx.beginPath();
            if (this.IsRequired())
                oCtx.fillStyle = 'rgb(' + FIELDS_HIGHLIGHT_REQ.r + ', ' + FIELDS_HIGHLIGHT_REQ.g + ', ' + FIELDS_HIGHLIGHT_REQ.b + ')';
            else 
                oCtx.fillStyle = 'rgb(' + FIELDS_HIGHLIGHT.r + ', ' + FIELDS_HIGHLIGHT.g + ', ' + FIELDS_HIGHLIGHT.b + ')';
            // выставляем в центр круга
            let centerX = X + W / 2;
            let centerY = Y + H / 2;
            let nRadius = Math.min(W / 2, H / 2);
            oCtx.arc(centerX, centerY, nRadius, 0, 2 * Math.PI, false);
            oCtx.fill();
            oCtx.closePath();
        }
        else if (this.GetType() != AscPDF.FIELD_TYPES.button){
            oCtx.beginPath();
            if (this.IsRequired())
                oCtx.fillStyle = 'rgb(' + FIELDS_HIGHLIGHT_REQ.r + ', ' + FIELDS_HIGHLIGHT_REQ.g + ', ' + FIELDS_HIGHLIGHT_REQ.b + ')';
            else 
                oCtx.fillStyle = 'rgb(' + FIELDS_HIGHLIGHT.r + ', ' + FIELDS_HIGHLIGHT.g + ', ' + FIELDS_HIGHLIGHT.b + ')';
            oCtx.fillRect(X, Y, W, H);
            oCtx.closePath();
        }
        
        AscPDF.endMultiplyMode(oCtx);
    };
    CBaseField.prototype.DrawBorders = function(oGraphicsPDF) {
        let aOringRect  = this.GetOrigRect();
        let aBgColor;
        if (this.GetType() == AscPDF.FIELD_TYPES.button)
            aBgColor = this.GetBackgroundColor() || [1];
        else
            aBgColor = this.IsNeedDrawHighlight() == false ? (this.GetBackgroundColor() || [1]) : [1];

        let oBgRGBColor = this.GetRGBColor(aBgColor);

        if (aBgColor && aBgColor.length != 0)
            oBgRGBColor = AscPDF.MakeColorMoreGray(oBgRGBColor, 50);
        
        let nLineWidth = this.GetBorderWidth() != undefined ? this.GetBorderWidth() : 1;

        if (nLineWidth == 0) {
            return;
        }

        oGraphicsPDF.SetLineWidth(nLineWidth);

        let X       = aOringRect[0];
        let Y       = aOringRect[1];
        let nWidth  = aOringRect[2] - aOringRect[0];
        let nHeight = aOringRect[3] - aOringRect[1];

        let color;
        if (this._strokeColor != null) {
            color = this.GetRGBColor(this._strokeColor);
            oGraphicsPDF.SetStrokeStyle(color.r, color.g, color.b);
        }

        // корректировка координат по бордеру
        Y += nLineWidth / 2;
        X += nLineWidth / 2;
        nWidth  -= nLineWidth;
        nHeight -= nLineWidth;

        // по умолчанию рисуется solid
        let nBorderStyle = this.GetBorderStyle() != undefined ? this.GetBorderStyle() : AscPDF.BORDER_TYPES.solid;

        if (this.GetType() == AscPDF.FIELD_TYPES.radiobutton && this._chStyle == AscPDF.CHECKBOX_STYLES.circle) {
            // выставляем в центр круга
            let centerX = X + nWidth / 2;
            let centerY = Y + nHeight / 2;
            let nRadius = Math.min(nWidth / 2, nHeight / 2);

            // отрисовка
            switch (nBorderStyle) {
                case AscPDF.BORDER_TYPES.solid:
                case AscPDF.BORDER_TYPES.underline:
                    if (color == null)
                        break;
                    oGraphicsPDF.SetLineDash([]);
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.Arc(centerX, centerY, nRadius, 0, 2 * Math.PI, false);
                    oGraphicsPDF.Stroke();
                    break;
                case AscPDF.BORDER_TYPES.beveled:
                    if (color) {
                        oGraphicsPDF.SetLineDash([]);
                        oGraphicsPDF.BeginPath();
                        oGraphicsPDF.Arc(centerX, centerY, nRadius, 0, 2 * Math.PI, false);
                        oGraphicsPDF.Stroke();
                    }

                    // left semicircle
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.Arc(centerX, centerY, nRadius - nLineWidth, 3 * Math.PI / 4, - Math.PI / 4, false);
                    if (this.IsPressed()) {
                        oGraphicsPDF.SetStrokeStyle(oBgRGBColor.r, oBgRGBColor.g, oBgRGBColor.b);
                    }
                    else {
                        oGraphicsPDF.SetStrokeStyle(255, 255, 255);
                    }
                    oGraphicsPDF.Stroke();

                    // right semicircle
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.Arc(centerX, centerY, nRadius - nLineWidth, - Math.PI / 4, 3 * Math.PI / 4, false);
                    if (this.IsPressed()) {
                        oGraphicsPDF.SetStrokeStyle(255, 255, 255);
                    }
                    else {
                        oGraphicsPDF.SetStrokeStyle(oBgRGBColor.r, oBgRGBColor.g, oBgRGBColor.b);
                    }
                    oGraphicsPDF.Stroke();

                    break;
                case AscPDF.BORDER_TYPES.dashed:
                    if (color == null)
                        break;

                    oGraphicsPDF.SetLineDash([3]);
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.Arc(centerX, centerY, nRadius, 0, 2 * Math.PI, false);
                    oGraphicsPDF.Stroke();
                    break;
                case AscPDF.BORDER_TYPES.inset:
                    if (color) {
                        oGraphicsPDF.SetLineDash([]);
                        oGraphicsPDF.BeginPath();
                        oGraphicsPDF.Arc(centerX, centerY, nRadius, 0, 2 * Math.PI, false);
                        oGraphicsPDF.Stroke();
                    }
                    
                    // left semicircle
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.Arc(centerX, centerY, nRadius - nLineWidth, 3 * Math.PI / 4, - Math.PI / 4, false);
                    if (this.IsPressed())
                        oGraphicsPDF.SetStrokeStyle(0, 0, 0);
                    else
                        oGraphicsPDF.SetStrokeStyle(128, 128, 128);
                    oGraphicsPDF.Stroke();

                    // right semicircle
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.Arc(centerX, centerY, nRadius - nLineWidth, - Math.PI / 4, 3 * Math.PI / 4, false);
                    if (this.IsPressed())
                        oGraphicsPDF.SetStrokeStyle(255, 255, 255);
                    else
                        oGraphicsPDF.SetStrokeStyle(191, 191, 191);
                    oGraphicsPDF.Stroke();

                    break;
            }

            return;
        }
        else {
            // отрисовка
            switch (nBorderStyle) {
                case AscPDF.BORDER_TYPES.solid:
                    if (color == null)
                        break;
                    
                    oGraphicsPDF.SetLineDash([]);
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.Rect(X, Y, nWidth, nHeight);
                    oGraphicsPDF.Stroke();

                    break;
                case AscPDF.BORDER_TYPES.beveled:
                    if (color) {
                        oGraphicsPDF.SetLineDash([]);
                        oGraphicsPDF.BeginPath();
                        oGraphicsPDF.Rect(X, Y, nWidth, nHeight);
                        oGraphicsPDF.Stroke();
                    }
                    
                    // left part
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.MoveTo(X + nLineWidth + nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nLineWidth + nLineWidth / 2, Y + nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                    
                    oGraphicsPDF.MoveTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nLineWidth + nLineWidth / 2, Y + nLineWidth / 2);

                    if (this.GetType() == AscPDF.FIELD_TYPES.button && this.IsPressed() && this.GetHighlight() == AscPDF.BUTTON_HIGHLIGHT_TYPES.push) {
                        oGraphicsPDF.SetFillStyle(oBgRGBColor.r, oBgRGBColor.g, oBgRGBColor.b);
                    }
                    else {
                        if ((this.GetType() == AscPDF.FIELD_TYPES.radiobutton || this.GetType() == AscPDF.FIELD_TYPES.checkbox) && this.IsPressed())
                            oGraphicsPDF.SetFillStyle(oBgRGBColor.r, oBgRGBColor.g, oBgRGBColor.b);
                        else
                            oGraphicsPDF.SetFillStyle(255, 255, 255);
                    }

                    oGraphicsPDF.ClosePath();
                    oGraphicsPDF.Fill();

                    // top part
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.MoveTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                    
                    oGraphicsPDF.MoveTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);

                    oGraphicsPDF.ClosePath();
                    oGraphicsPDF.Fill();

                    // bottom part
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.MoveTo(X + nLineWidth + nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                    
                    oGraphicsPDF.MoveTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);

                    if (this.GetType() == AscPDF.FIELD_TYPES.button && this.IsPressed() && this.GetHighlight() == AscPDF.BUTTON_HIGHLIGHT_TYPES.push) {
                        oGraphicsPDF.SetFillStyle(255, 255, 255);
                    }
                    else {
                        if ((this.GetType() == AscPDF.FIELD_TYPES.radiobutton || this.GetType() == AscPDF.FIELD_TYPES.checkbox) && this.IsPressed())
                            oGraphicsPDF.SetFillStyle(255, 255, 255);
                        else
                            oGraphicsPDF.SetFillStyle(oBgRGBColor.r, oBgRGBColor.g, oBgRGBColor.b);
                    }

                    oGraphicsPDF.ClosePath();
                    oGraphicsPDF.Fill();

                    // right part
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.MoveTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nHeight - nLineWidth);
                    oGraphicsPDF.LineTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                    
                    oGraphicsPDF.MoveTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth);
                    oGraphicsPDF.LineTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nHeight - nLineWidth);

                    oGraphicsPDF.ClosePath();
                    oGraphicsPDF.Fill();

                    break;
                case AscPDF.BORDER_TYPES.dashed:
                    if (color == null)
                        break;

                    oGraphicsPDF.SetLineDash([3]);
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.Rect(X, Y, nWidth, nHeight);
                    oGraphicsPDF.Stroke();
                    break;
                case AscPDF.BORDER_TYPES.inset:
                    if (color) {
                        oGraphicsPDF.SetLineDash([]);
                        oGraphicsPDF.BeginPath();
                        oGraphicsPDF.Rect(X, Y, nWidth, nHeight);
                        oGraphicsPDF.Stroke();
                    }

                    // left part
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.MoveTo(X + nLineWidth + nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nLineWidth + nLineWidth / 2, Y + nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                    
                    oGraphicsPDF.MoveTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nLineWidth + nLineWidth / 2, Y + nLineWidth / 2);

                    if (this.GetType() == AscPDF.FIELD_TYPES.button && this.IsPressed() && this.GetHighlight() == AscPDF.BUTTON_HIGHLIGHT_TYPES.push) {
                        oGraphicsPDF.SetFillStyle(0, 0, 0);
                    }
                    else {
                        if ((this.GetType() == AscPDF.FIELD_TYPES.radiobutton || this.GetType() == AscPDF.FIELD_TYPES.checkbox) && this.IsPressed())
                            oGraphicsPDF.SetFillStyle(0, 0, 0);
                        else
                            oGraphicsPDF.SetFillStyle(128, 128, 128);
                    }

                    oGraphicsPDF.ClosePath();
                    oGraphicsPDF.Fill();

                    // top part
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.MoveTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                    
                    oGraphicsPDF.MoveTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);

                    oGraphicsPDF.ClosePath();
                    oGraphicsPDF.Fill();

                    // bottom part
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.MoveTo(X + nLineWidth + nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                    
                    oGraphicsPDF.MoveTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);

                    if (this.GetType() == AscPDF.FIELD_TYPES.button && this.IsPressed() && this.GetHighlight() == AscPDF.BUTTON_HIGHLIGHT_TYPES.push) {
                        oGraphicsPDF.SetFillStyle(255, 255, 255);
                    }
                    else {
                        if ((this.GetType() == AscPDF.FIELD_TYPES.radiobutton || this.GetType() == AscPDF.FIELD_TYPES.checkbox) && this.IsPressed())
                            oGraphicsPDF.SetFillStyle(255, 255, 255);
                        else
                            oGraphicsPDF.SetFillStyle(191, 191, 191);
                    }
                    
                    oGraphicsPDF.ClosePath();
                    oGraphicsPDF.Fill();

                    // right part
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.MoveTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nHeight - nLineWidth);
                    oGraphicsPDF.LineTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                    
                    oGraphicsPDF.MoveTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                    oGraphicsPDF.LineTo(X + nWidth - nLineWidth / 2, Y + nHeight - nLineWidth);
                    oGraphicsPDF.LineTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nHeight - nLineWidth);

                    oGraphicsPDF.ClosePath();
                    oGraphicsPDF.Fill();

                    break;
                case AscPDF.BORDER_TYPES.underline:
                    if (color == null)
                        break;
                    
                    oGraphicsPDF.SetLineDash([]);
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.MoveTo(X - nLineWidth / 2, Y + nHeight);
                    oGraphicsPDF.LineTo(X + nWidth + nLineWidth / 2, Y + nHeight);
                    oGraphicsPDF.Stroke();
                    break;
            }
        }        

        // pressed border
        if (this.GetType() == AscPDF.FIELD_TYPES.button && this.IsPressed() && this.GetHighlight() == AscPDF.BUTTON_HIGHLIGHT_TYPES.push && this._imgData.mouseDown == undefined) {
            switch (nBorderStyle) {
                case AscPDF.BORDER_TYPES.solid:
                case AscPDF.BORDER_TYPES.dashed:
                case AscPDF.BORDER_TYPES.underline: {
                    
                        // left part
                        oGraphicsPDF.SetFillStyle(oBgRGBColor.r, oBgRGBColor.g, oBgRGBColor.b);

                        oGraphicsPDF.BeginPath();
                        oGraphicsPDF.MoveTo(X + nLineWidth + nLineWidth / 2, Y + nHeight - nLineWidth - nLineWidth / 2);
                        oGraphicsPDF.LineTo(X + nLineWidth + nLineWidth / 2, Y + nLineWidth / 2);
                        oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                        oGraphicsPDF.MoveTo(X + nLineWidth / 2, Y + nHeight - nLineWidth / 2);
                        oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nLineWidth / 2);
                        oGraphicsPDF.LineTo(X + nLineWidth + nLineWidth / 2, Y + nLineWidth / 2);
                        oGraphicsPDF.ClosePath();
                        oGraphicsPDF.Fill();

                        // top part
                        oGraphicsPDF.BeginPath();
                        oGraphicsPDF.MoveTo(X + nWidth - nLineWidth - nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                        oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                        oGraphicsPDF.LineTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                        oGraphicsPDF.MoveTo(X + nWidth - nLineWidth / 2, Y + nLineWidth / 2);
                        oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nLineWidth / 2);
                        oGraphicsPDF.LineTo(X + nLineWidth / 2, Y + nLineWidth + nLineWidth / 2);
                        oGraphicsPDF.ClosePath();
                        oGraphicsPDF.Fill();
                    }
                    break;
                }
        }

        // draw comb cells
        if ((this.GetType() == AscPDF.FIELD_TYPES.text && this.IsComb() == true) && this._borderColor != null && (nBorderStyle == AscPDF.BORDER_TYPES.solid || nBorderStyle == AscPDF.BORDER_TYPES.dashed)) {
            let nCombWidth = (nWidth / this.GetCharLimit());
            let nIndentX = nCombWidth;
            
            for (let i = 0; i < this.GetCharLimit() - 1; i++) {
                oGraphicsPDF.MoveTo(X + nIndentX, Y);
                oGraphicsPDF.LineTo(X + nIndentX, Y + nHeight);
                oGraphicsPDF.Stroke();
                nIndentX += nCombWidth;
            }
        }
    };
    /**
	 * Gets rgb color object from internal color array.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
     * @returns {object}
	 */
    CBaseField.prototype.GetRGBColor = function(aInternalColor) {
        let oColor = {};

        if (aInternalColor.length == 1) {
            oColor = {
                r: Math.round(aInternalColor[0] * 255),
                g: Math.round(aInternalColor[0] * 255),
                b: Math.round(aInternalColor[0] * 255)
            }
        }
        else if (aInternalColor.length == 3) {
            oColor = {
                r: Math.round(aInternalColor[0] * 255),
                g: Math.round(aInternalColor[1] * 255),
                b: Math.round(aInternalColor[2] * 255)
            }
        }
        else if (aInternalColor.length == 4) {
            function cmykToRgb(c, m, y, k) {
                return {
                    r: Math.round(255 * (1 - c) * (1 - k)),
                    g: Math.round(255 * (1 - m) * (1 - k)),
                    b: Math.round(255 * (1 - y) * (1 - k))
                }
            }

            oColor = cmykToRgb(aInternalColor[0], aInternalColor[1], aInternalColor[2], aInternalColor[3]);
        }

        return oColor;
    };
    
    CBaseField.prototype.DrawSelected = function() {
        return;
    };
    CBaseField.prototype.SetParentPage = function(oParent) {
        this.parentPage = oParent;
    };
    CBaseField.prototype.GetParentPage = function() {
        return this.parentPage;
    };
    CBaseField.prototype.Get_Id = function() {
        return this.Id;
    };
    CBaseField.prototype.GetId = function() {
        return this.Id;
    };
    CBaseField.prototype.SetNeedRecalc = function(bRecalc, bSkipAddToRedraw) {
        if (bRecalc == false) {
            this._needRecalc = false;
        }
        else {
            if ([AscPDF.FIELD_TYPES.text, AscPDF.FIELD_TYPES.combobox].includes(this.GetType())) {
                let oDoc = this.GetDocument();
                if (oDoc) {
                    oDoc.SetNeedUpdateTarget(true);
                }
            }
            this._needRecalc = true;
            if (bSkipAddToRedraw != true)
                this.AddToRedraw();
        }
    };
    CBaseField.prototype.DrainLogicFrom = function(oFieldToInherit, bClearFrom) {
        this.SetParentValue(oFieldToInherit.GetParentValue());
        this.SetDefaultValue(oFieldToInherit.GetDefaultValue());
        this.SetReadOnly(oFieldToInherit.IsReadOnly());
        this.SetNoExport(oFieldToInherit.IsNoExport());
        this.SetRequired(oFieldToInherit.IsRequired());
        
        if (bClearFrom !== false) {
            oFieldToInherit.SetDefaultValue(undefined);
            oFieldToInherit.SetParentValue(undefined);
            oFieldToInherit.SetReadOnly(false);
            oFieldToInherit.SetNoExport(false);
            oFieldToInherit.SetRequired(false);
        }

        let _t = this;

        let aLogicActions = [
            AscPDF.FORMS_TRIGGERS_TYPES.Keystroke,
            AscPDF.FORMS_TRIGGERS_TYPES.Validate,
            AscPDF.FORMS_TRIGGERS_TYPES.Calculate,
            AscPDF.FORMS_TRIGGERS_TYPES.Format
        ]

        Object.values(aLogicActions).forEach(function(type) {
            _t.SetActions(type, oFieldToInherit.GetActions(type));

            if (bClearFrom !== false) {
                oFieldToInherit.SetActions(type, []);
            }
        });
    };

    CBaseField.prototype.SetName = function(sName) {
        let oDoc        = this.GetDocument();
        let nFieldType  = this.GetType();

        while (sName.indexOf('..') != -1)
            sName = sName.replace(new RegExp("\.\.", "g"), ".");

        let oExistsWidget = oDoc.GetField(sName);
        // если есть виджет-поле с таким именем то не добавляем 
        if (oExistsWidget && oExistsWidget.GetType() != nFieldType)
            return false; // to do выдавать ошибку смены имени поля

        // получаем partial names
        let aPartNames = sName.split('.').filter(function(item) {
            if (item != "")
                return item;
        })

        // по формату не больше 20 вложенностей
        if (aPartNames.length > 20)
            return false;

        // удаляем поле из родителя
        let oParent = this.GetParent();
        if (oParent) {
            oParent.RemoveKid(this);
        }

        let aWidgetForms = oDoc.GetAllWidgets(sName);

        if (aPartNames.length > 0) {
            let oExistsField = oDoc.GetField(sName);
            let oParentField;

            if (oExistsField) {
                if (!oExistsField.IsWidget()) {
                    if (!oExistsField.IsAllKidsWidgets()) {
                        return false;
                    }
                    else {
                        oParentField = oExistsField;
                    }
                }
            }

            let bParentInherit = !oParentField;
            if (!oParentField) {
                if (oExistsField) {
                    for (let i = 0; i < aPartNames.length; i++) {
                        let oNewParent = oDoc.CreateField(aPartNames[i], nFieldType, []);
    
                        if (oParentField) {
                            oParentField.AddKid(oNewParent);
                        }
                        
                        oParentField = oNewParent;
                    }
                }
                // will be field-widget
                else {
                    if (oParent) {
                        this.DrainLogicFrom(oParent, false);
                    }
                    
                    this.SetPartialName(sName);
                    return true;
                }
            }

            if (bParentInherit) {
                oParentField.DrainLogicFrom(oExistsField);
            }

            aWidgetForms.forEach(function(widget) {
                oParentField.AddKid(widget);
                widget.SetPartialName(undefined);
            });

            oParentField.AddKid(this);
            this.SetPartialName(undefined);
        }

        this.SyncValue();
        return true;
    };
    CBaseField.prototype.IsNeedRecalc = function() {
        return this._needRecalc;
    };
    CBaseField.prototype.Refresh_RecalcData = function(){};
    CBaseField.prototype.SetWasChanged = function(isChanged) {
        let oViewer = Asc.editor.getDocumentRenderer();
        let oDoc = Asc.editor.getPDFDoc();
        let canChange = !oViewer.IsOpenFormsInProgress && !oDoc.History.UndoRedoInProgress;
        if (this._wasChanged == isChanged || !canChange) {
            return;
        }
        
        oDoc.History.Add(new CChangesPDFFormChanged(this, this._wasChanged, isChanged));

        this._wasChanged = isChanged;
        this.IsWidget() && this.SetDrawFromStream(!isChanged);
    };

    CBaseField.prototype.UndoNotAppliedChanges = function() {
        let isChanged = this.IsChanged();
        this.SetValue(this.GetParentValue());
        this.SetNeedRecalc(true);
        this.SetNeedCommit(false);

        if (!isChanged)
            this.SetWasChanged(false);
    };

    CBaseField.prototype.ClearCache = function() {
        this._originView.normal     = null;
        this._originView.mouseDown  = null;
        this._originView.rollover   = null;
    };
    CBaseField.prototype.IsChanged = function() {
        return this._wasChanged;  
    };
    CBaseField.prototype.SetHasOriginView = function(bHas) {
        this._hasOriginView = bHas;
        this.SetDrawFromStream(bHas);
    };
    CBaseField.prototype.HasOriginView = function() {
        return this._hasOriginView;
    };
    CBaseField.prototype.IsNeedDrawFromStream = function() {
        return this._bDrawFromStream;
    };
    CBaseField.prototype.SetDrawFromStream = function(bFromStream) {
        if (bFromStream && this.HasOriginView())
            this._bDrawFromStream = true;
        else
            this._bDrawFromStream = false;
    };
    CBaseField.prototype.SetDrawHighlight = function(bDraw) {
        this._needDrawHighlight = bDraw;
    };
    CBaseField.prototype.IsNeedDrawHighlight = function() {
        return false == this.IsReadOnly() && this._needDrawHighlight;
    };

    CBaseField.prototype.DrawEdit = function(oGraphicsWord) {
        if (this.IsEditMode() && !oGraphicsWord.isSkipEditShapes) {
            this.editShape.Draw(oGraphicsWord);
        }
    };
    CBaseField.prototype.AddToRedraw = function() {
        let oViewer = editor.getDocumentRenderer();
        let nPage   = this.GetPage();
        
        if (false == this.IsUseInDocument()) {
            return;
        }
        
        function setRedrawPageOnRepaint() {
            if (oViewer.pagesInfo.pages[nPage]) {
                oViewer.pagesInfo.pages[nPage].needRedrawForms = true;
                // oViewer.thumbnails && oViewer.thumbnails._repaintPage(nPage);
            }
        }

        oViewer.paint(setRedrawPageOnRepaint);
    };

    CBaseField.prototype.GetType = function() {
        return this.type;
    };
    CBaseField.prototype.GetStrType = function() {
        switch (this.GetType()) {
            case AscPDF.FIELD_TYPES.button:
                return "button";
            case AscPDF.FIELD_TYPES.checkbox:
                return "checkbox";
            case AscPDF.FIELD_TYPES.combobox:
                return "combobox";
            case AscPDF.FIELD_TYPES.listbox:
                return "listbox";
            case AscPDF.FIELD_TYPES.radiobutton:
                return "radiobutton";
            case AscPDF.FIELD_TYPES.signature:
                return "signature";
            case AscPDF.FIELD_TYPES.text:
                return "text";
            case AscPDF.FIELD_TYPES.unknown: 
                return "unknown";
        }
    };

    CBaseField.prototype.SetReadOnly = function(bReadOnly) {
        let oParent = this.GetParent();
        if (oParent && oParent.IsAllKidsWidgets())
            return oParent.SetReadOnly(bReadOnly);

        if (this._readOnly === bReadOnly) {
            return true;
        }

        AscCommon.History.Add(new CChangesPDFFormReadOnly(this, this._readOnly, bReadOnly));

        function update(widget) {
            widget.SetWasChanged(true);
            widget.AddToRedraw();
        };
    
        if (this.IsWidget()) {
            update(this);
        } else {
            this.GetAllWidgets().forEach(update);
        }

        this._readOnly = bReadOnly;
        this.SetWasChanged(true);
        this.AddToRedraw();

        return true;
    };
    CBaseField.prototype.IsReadOnly = function(bInherit) {
        let oParent = this.GetParent();
        if (bInherit !== false && oParent && oParent.IsAllKidsWidgets())
            return oParent.IsReadOnly();

        return this._readOnly;
    };

    CBaseField.prototype.SetNoExport = function(bNoExport) {
        let oParent = this.GetParent();
        if (oParent && oParent.IsAllKidsWidgets()) {
            return oParent.SetNoExport(bNoExport);
        }
    
        if (this._noExport === bNoExport) {
            return true;
        }

        AscCommon.History.Add(new CChangesPDFFormNoExport(this, this._noExport, bNoExport));
        this._noExport = bNoExport;
    
        this.SetWasChanged(true);
        this.SetNeedRecalc(true);

        return true;
    };
    CBaseField.prototype.IsNoExport = function(bInherit) {
        let oParent = this.GetParent();
        if (bInherit !== false && oParent && oParent.IsAllKidsWidgets())
            return oParent.IsNoExport();

        return this._noExport;
    };
    
    CBaseField.prototype.SetRequired = function(bRequired) {
        let oParent = this.GetParent();
        if (oParent && oParent.IsAllKidsWidgets())
            return oParent.SetRequired(bRequired);

        if (this._required === bRequired) {
            return true;
        }

        if (this.GetType() === AscPDF.FIELD_TYPES.button) {
            return false;
        }

        AscCommon.History.Add(new CChangesPDFFormRequired(this, this._required, bRequired));

        function update(widget) {
            widget.SetWasChanged(true);
            widget.AddToRedraw();
        };
    
        if (this.IsWidget()) {
            update(this);
        } else {
            this.GetAllWidgets().forEach(update);
        }

        this._required = bRequired;
        this.SetWasChanged(true);
        this.AddToRedraw();

        return true;
    };
    CBaseField.prototype.IsRequired = function(bInherit) {
        let oParent = this.GetParent();
        if (bInherit !== false && oParent && oParent.IsAllKidsWidgets())
            return oParent.IsRequired();

        return this._required;
    };
    CBaseField.prototype.SetBorderColor = function(aColor) {
        if (aColor && aColor.length == 0) {
            return false;
        }

        AscCommon.History.Add(new CChangesPDFFormBorderColor(this, this.GetBorderColor(), aColor));

        this._strokeColor = this._borderColor = aColor;
        this.SetWasChanged(true);
        this.AddToRedraw();
    };
    CBaseField.prototype.GetBorderColor = function() {
        return this._strokeColor;
    };
    CBaseField.prototype.SetBackgroundColor = function(aColor) {
        if (aColor && aColor.length == 0) {
            return false;
        }
        
        AscCommon.History.Add(new CChangesPDFFormBGrColor(this, this.GetBackgroundColor(), aColor));

        this._fillColor = this._bgColor = aColor;
        this.SetWasChanged(true);
        this.AddToRedraw();
    };
    CBaseField.prototype.GetBackgroundColor = function() {
        return this._fillColor;
    };
    CBaseField.prototype.SetHighlight = function(nType) {
        this._highlight = nType;
    };
    CBaseField.prototype.GetHighlight = function() {
        return this._highlight;
    };
    CBaseField.prototype.IsHidden = function() {
        let nType = this.GetDisplay();
        if (nType == window["AscPDF"].Api.Types.display["hidden"] || nType == window["AscPDF"].Api.Types.display["noView"])
            return true;

        return false;
    };
    CBaseField.prototype.SetDisplay = function(nType) {
        AscCommon.History.Add(new CChangesPDFFormDisplay(this, this._display, nType));

        this._display = nType;

        this.SetWasChanged(true);
        this.AddToRedraw();
    };
    CBaseField.prototype.GetDisplay = function() {
        return this._display;
    };
    CBaseField.prototype.GetDefaultValue = function(bInherit) {
        let oParent = this.GetParent();
        if (bInherit !== false && oParent && oParent.IsAllKidsWidgets())
            return oParent.GetDefaultValue();
        
        return this._defaultValue;
    };
    CBaseField.prototype.SetDefaultValue = function(value) {
        let oParent = this.GetParent();
        if (oParent && oParent.IsAllKidsWidgets())
            return oParent.SetDefaultValue(value);

        let sOldDefValue = this.GetDefaultValue();
        if (value === sOldDefValue) {
            return true;
        }

        AscCommon.History.Add(new CChangesPDFFormDefaultValue(this, sOldDefValue, value));

        this._defaultValue = value;
        this.SetWasChanged(true);

        let oWidget = this.IsWidget() ? this : this.GetKid(0);
        if (oWidget.IsWidget()) {
            if (!value && this.GetParentValue() == sOldDefValue) {
                oWidget.SetValue(value);
                oWidget.Commit();
            }
            else if (value && (!this.GetParentValue() || (this.GetType() == AscPDF.FIELD_TYPES.checkbox && this.GetParentValue() == "Off"))) {
                oWidget.SetValue(value);
                oWidget.Commit();
            }
        }

        return true;
    };
	
	CBaseField.prototype.canBeginCompositeInput = function() {
		return false;
	};
	CBaseField.prototype.beforeCompositeInput = function() {
	};
	CBaseField.prototype.EnterText = function(codePoints) {
	};
    /**
	 * Sets default value for form.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 */
    CBaseField.prototype.Reset = function() {
        let defValue = this.GetDefaultValue() || "";
        if (this.GetValue() != defValue) {
            this.SetValue(defValue);
            this.SetWasChanged(true);
            this.SetNeedRecalc(true);
        }
    };

    CBaseField.prototype.DrawBackground = function(oGraphicsPDF) {
        let aOrigRect       = this.GetOrigRect();
        let aBgColor        = this.GetBackgroundColor();
        let oBgRGBColor;

        if (null == aBgColor && this.IsPressed && this.IsPressed())
            aBgColor = [1];
        
        if (aBgColor && aBgColor.length != 0)
            oBgRGBColor = this.GetRGBColor(aBgColor);

        if (!oBgRGBColor || this.IsNeedDrawHighlight())
            return;

        let X       = aOrigRect[0];
        let Y       = aOrigRect[1];
        let nWidth  = aOrigRect[2] - aOrigRect[0];
        let nHeight = aOrigRect[3] - aOrigRect[1];
        
        oGraphicsPDF.SetGlobalAlpha(1);

        if (this.GetType() == AscPDF.FIELD_TYPES.radiobutton && this._chStyle == AscPDF.CHECKBOX_STYLES.circle) {
            if (this.IsHovered() && this.IsPressed()) {
                if (aBgColor.length == 1 && aBgColor[0] == 1) {
                    oBgRGBColor = {r: 191, g: 0, b: 0};
                }
                else {
                    if (this.GetBorderStyle() !== AscPDF.BORDER_TYPES.beveled)
                        oBgRGBColor = AscPDF.MakeColorMoreGray(oBgRGBColor, 50);
                }
            }

            oGraphicsPDF.BeginPath();
            oGraphicsPDF.SetFillStyle(oBgRGBColor.r, oBgRGBColor.g, oBgRGBColor.b);
            // выставляем в центр круга
            let centerX = X + nWidth / 2;
            let centerY = Y + nHeight / 2;
            let nRadius = Math.min(nWidth / 2, nHeight / 2);
            oGraphicsPDF.Arc(centerX, centerY, nRadius, 0, 2 * Math.PI, false);
            oGraphicsPDF.Fill();
            oGraphicsPDF.ClosePath();
        }
        else {
            if ((this.GetType() == AscPDF.FIELD_TYPES.radiobutton || this.GetType() == AscPDF.FIELD_TYPES.checkbox) && this.IsPressed() && this.GetBorderStyle() !== AscPDF.BORDER_TYPES.beveled) {
                oBgRGBColor = AscPDF.MakeColorMoreGray(this.GetRGBColor(aBgColor), 50);
            }
            
            oGraphicsPDF.SetFillStyle(oBgRGBColor.r, oBgRGBColor.g, oBgRGBColor.b);
            oGraphicsPDF.FillRect(X, Y, nWidth, nHeight);
        }
    };

    /**
	 * Gets Api class for this form.
	 * @memberof CBaseField
     * @param {number} nIdx - The 0-based index of the item in the list or -1 for the last item in the list.
     * @param {boolean} [bExportValue=true] - Specifies whether to return an export value.
	 * @typeofeditors ["PDF"]
     * @returns {ApiBaseField}
	 */
    CBaseField.prototype.GetFormApi = function() {
        if (this.api)
            return this.api;

        switch (this.GetType()) {
            case AscPDF.FIELD_TYPES.text:
                return new AscPDF.ApiTextField(this);
            case AscPDF.FIELD_TYPES.combobox:
                return new AscPDF.ApiComboBoxField(this);
            case AscPDF.FIELD_TYPES.listbox:
                return new AscPDF.ApiListBoxField(this);
            case AscPDF.FIELD_TYPES.checkbox:
                return new AscPDF.ApiCheckBoxField(this);
            case AscPDF.FIELD_TYPES.radiobutton:
                return new AscPDF.ApiRadioButtonField(this);
            case AscPDF.FIELD_TYPES.button:
                return new AscPDF.ApiPushButtonField(this);
        }

        return null;
    };

    // pdf api methods

    /**
	 * A string that sets the trigger for the action. Values are:
	 * @typedef {"MouseUp" | "MouseDown" | "MouseEnter" | "MouseExit" | "OnFocus" | "OnBlur" | "Keystroke" | "Validate" | "Calculate" | "Format"} cTrigger
	 * For a list box, use the Keystroke trigger for the Selection Change event.
     */
    CBaseField.prototype.RevertContentView = function() {
        this.content.ResetShiftView();
        this._curShiftView.x = this._oldShiftView.x;
        this._curShiftView.y = this._oldShiftView.y;

        this._bAutoShiftContentView = false;
        this.content.ShiftView(this._oldShiftView.x, this._oldShiftView.y);

        if (this._scrollInfo) {
            let nMaxShiftY                  = this._scrollInfo.scroll.maxScrollY;
            this._scrollInfo.scrollCoeff    = Math.abs(this._curShiftView.y / nMaxShiftY);
        }

        this.AddToRedraw();
    };
    CBaseField.prototype.IsWidget = function() {
        let aRect = this.GetRect();
        return !!(aRect && aRect.length == 4);
    };
    CBaseField.prototype.IsNeedRevertShiftView = function() {
        if (this._curShiftView.y != this._oldShiftView.y ||
            this._curShiftView.x != this._oldShiftView.x)
            return true;
    };
    CBaseField.prototype.GetBordersWidth = function() {
        let nLineWidth = this._lineWidth != undefined ? this._lineWidth : 1;

        if (nLineWidth == 0 || this._borderStyle == undefined) {
            return {
                left:     nLineWidth,
                top:      nLineWidth,
                right:    nLineWidth,
                bottom:   nLineWidth
            }
        }

        switch (this._borderStyle) {
            case AscPDF.BORDER_TYPES.solid:
            case AscPDF.BORDER_TYPES.dashed:
                return {
                    left:     nLineWidth,
                    top:      nLineWidth,
                    right:    nLineWidth,
                    bottom:   nLineWidth
                }
            case AscPDF.BORDER_TYPES.beveled:
            case AscPDF.BORDER_TYPES.inset:
                return {
                    left:     2 * nLineWidth,
                    top:      2 * nLineWidth,
                    right:    2 * nLineWidth,
                    bottom:   2 * nLineWidth
                }
            case AscPDF.BORDER_TYPES.underline:
                return {
                    left:     0,
                    top:      0,
                    right:    0,
                    bottom:   nLineWidth
                }
        }
    };
    CBaseField.prototype.GetMarginsFromBorders = function() {
        let oBorders = this.GetBordersWidth();

        switch (this._borderStyle) {
            case AscPDF.BORDER_TYPES.solid:
            case AscPDF.BORDER_TYPES.dashed:
            case AscPDF.BORDER_TYPES.underline:
                return {
                    left:     oBorders.bottom,
                    top:      oBorders.bottom,
                    right:    oBorders.bottom,
                    bottom:   oBorders.bottom
                }
            case AscPDF.BORDER_TYPES.inset:
            case AscPDF.BORDER_TYPES.beveled:
                return {
                    left:     oBorders.bottom,
                    top:      oBorders.bottom,
                    right:    oBorders.bottom,
                    bottom:   oBorders.bottom
                }
            default:
                return {
                    left:     oBorders.bottom,
                    top:      oBorders.bottom,
                    right:    oBorders.bottom,
                    bottom:   oBorders.bottom
                }
        };
    };
    CBaseField.prototype.HasShiftView = function() {
        if (this.content.ShiftViewX != 0 || this.content.ShiftViewY != 0)
            return true;

        return false;
    };
    CBaseField.prototype.MoveCursorToStartPos = function() {
        this.content.MoveCursorToStartPos();
    };
    
    CBaseField.prototype.SetBorderStyle = function(nStyle) {
        if (this._borderStyle != nStyle) {
            AscCommon.History.Add(new CChangesPDFFormBorderStyle(this, this._borderStyle, nStyle));

            this._borderStyle = nStyle;

            this.SetWasChanged(true);
            this.SetNeedRecalc(true);
            if (this.IsComb && this.IsComb() == true) {
                this.content.GetElement(0).Content.forEach(function(run) {
                    run.RecalcInfo.Measure = true;
                });
            }
        }
    };
    CBaseField.prototype.GetBorderStyle = function() {
        return this._borderStyle;
    };
    CBaseField.prototype.SetBorderWidth = function(nWidth) {
        if (this._borderWidth != nWidth) {
            AscCommon.History.Add(new CChangesPDFFormBorderWidth(this, this._borderWidth, nWidth));

            this._borderWidth = nWidth;
            this._lineWidth = nWidth;

            this.SetWasChanged(true);
            this.SetNeedRecalc(true);
        }
    };
    CBaseField.prototype.GetBorderWidth = function() {
        return this._borderWidth || this._lineWidth;
    };
    /**
     * Returns a canvas with origin view (from appearance stream) of current form.
	 * @memberof CBaseField
     * @param {number} nAPType - AscPDF.APPEARANCE_TYPES (type of AP)
	 * @typeofeditors ["PDF"]
     * @returns {canvas}
	 */
    CBaseField.prototype.GetOriginView = function(nAPType, nPageW, nPageH) {
        if (this._apIdx == -1)
            return null;

        let oViewer = editor.getDocumentRenderer();
        let oFile   = oViewer.file;
        
        let oApearanceInfo  = this.GetOriginViewInfo(nPageW, nPageH);
        let oSavedView, oApInfoTmp;
        if (!oApearanceInfo)
            return null;
            
        switch (nAPType) {
            case AscPDF.APPEARANCE_TYPES.normal:
                oApInfoTmp = oApearanceInfo["N"];
                oSavedView = this._originView.normal;
                break;
            case AscPDF.APPEARANCE_TYPES.rollover:
                if (oApearanceInfo["R"]) {
                    oApInfoTmp = oApearanceInfo["R"]
                    oSavedView = this._originView.rollover;
                }
                else {
                    oApInfoTmp = oApearanceInfo["N"]
                    oSavedView = this._originView.normal;
                }
                break;
            case AscPDF.APPEARANCE_TYPES.mouseDown:
                if (oApearanceInfo["D"]) {
                    oApInfoTmp = oApearanceInfo["D"]
                    oSavedView = this._originView.mouseDown;
                }
                else {
                    oApInfoTmp = oApearanceInfo["N"]
                    oSavedView = this._originView.normal;
                }
                break;
            default:
                oApInfoTmp = oApearanceInfo["N"];
                oSavedView = this._originView.normal;
                break;
        }

        if (oSavedView && oSavedView.width == oApearanceInfo["w"] && oSavedView.height == oApearanceInfo["h"])
            return oSavedView;
        
        let canvas  = document.createElement("canvas");
        let nWidth  = oApearanceInfo["w"];
        let nHeight = oApearanceInfo["h"];

        canvas.width    = nWidth;
        canvas.height   = nHeight;

        canvas.x    = oApearanceInfo["x"];
        canvas.y    = oApearanceInfo["y"];
        
        if (!oApInfoTmp)
            return null;

        let supportImageDataConstructor = (AscCommon.AscBrowser.isIE && !AscCommon.AscBrowser.isIeEdge) ? false : true;

        let ctx             = canvas.getContext("2d");
        let mappedBuffer    = oFile.getUint8ClampedArray(oApInfoTmp["retValue"], 4 * nWidth * nHeight);
        let imageData       = null;

        if (supportImageDataConstructor)
        {
            imageData = new ImageData(mappedBuffer, nWidth, nHeight);
        }
        else
        {
            imageData = ctx.createImageData(nWidth, nHeight);
            imageData.data.set(mappedBuffer, 0);                    
        }
        if (ctx)
            ctx.putImageData(imageData, 0, 0);
        
        oViewer.file.free(oApInfoTmp["retValue"]);

        switch (nAPType) {
            case AscPDF.APPEARANCE_TYPES.normal:
                this._originView.normal = canvas;
                break;
            case AscPDF.APPEARANCE_TYPES.rollover:
                if (oApearanceInfo["R"]) {
                    this._originView.rollover = canvas;
                }
                else {
                    this._originView.normal = canvas;
                }
                break;
            case AscPDF.APPEARANCE_TYPES.mouseDown:
                if (oApearanceInfo["D"]) {
                    this._originView.mouseDown = canvas;
                }
                else {
                    this._originView.normal = canvas;
                }
                break;
            default:
                this._originView.normal = canvas;
                break;
        }

        return canvas;
    };
    /**
     * Returns AP info of this field.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
     * @returns {Object}
	 */
    CBaseField.prototype.GetOriginViewInfo = function(nPageW, nPageH) {
        let oViewer     = editor.getDocumentRenderer();
        let oFile       = oViewer.file;
        let nPage       = this.GetOriginPage();
        let oOriginPage = oFile.pages.find(function(page) {
            return page.originIndex == nPage;
        });

        if (oOriginPage.fieldsAPInfo == null || oOriginPage.fieldsAPInfo.size.w != nPageW || oOriginPage.fieldsAPInfo.size.h != nPageH) {
            oOriginPage.fieldsAPInfo = {
                info: oFile.nativeFile["getInteractiveFormsAP"](nPage, nPageW, nPageH),
                size: {
                    w: nPageW,
                    h: nPageH
                }
            }
        }
        
        for (let i = 0; i < oOriginPage.fieldsAPInfo.info.length; i++) {
            if (oOriginPage.fieldsAPInfo.info[i]["i"] == this._apIdx)
                return oOriginPage.fieldsAPInfo.info[i];
        }

        return null;
    };
	
	CBaseField.prototype.DrawOnPage = function(pdfGraphics, textBoxGraphics, pageIndex) {
		if (this.IsHidden() && !this.IsEditMode())
			return;
		
        if (pdfGraphics.isThumbnails) {
            this.DrawBorders(pdfGraphics);
        }
        else {
            if (this.IsNeedDrawFromStream())
                this.DrawFromStream(pdfGraphics, textBoxGraphics);
            else
                this.DrawFromTextBox(pdfGraphics, textBoxGraphics, pageIndex);
        }
	};

    CBaseField.prototype.DrawFromStream = function(oGraphicsPDF, oGraphicsWord) {
        let nAPType = this.IsHovered && this.IsHovered() ? AscPDF.AscPDF.APPEARANCE_TYPES.rollover : undefined;

        let originView = this.GetOriginView(nAPType, oGraphicsPDF.GetDrawingPageW(), oGraphicsPDF.GetDrawingPageH());

        let aOringRect = this.GetOrigRect();

        let X = aOringRect[0];
        let Y = aOringRect[1];

        if (originView) {
            oGraphicsPDF.SetGlobalAlpha(1);
            oGraphicsPDF.DrawImageXY(originView, X, Y, undefined, true);
        }

        // oGraphicsPDF.SetLineWidth(1);
        // let nWidth  = aOringRect[2] - aOringRect[0];
        // let nHeight = aOringRect[3] - aOringRect[1];

        // Y += 1 / 2;
        // X += 1 / 2;
        // nWidth  -= 1;
        // nHeight -= 1;

        // oGraphicsPDF.SetStrokeStyle(0, 255, 255);
        // oGraphicsPDF.SetLineDash([]);
        // oGraphicsPDF.BeginPath();
        // oGraphicsPDF.Rect(X, Y, nWidth, nHeight);
        // oGraphicsPDF.Stroke();

        this.DrawLocks(oGraphicsPDF);
        this.DrawEdit(oGraphicsWord);
    };
    CBaseField.prototype.DrawLocks = function(oGraphicsPDF) {
        let aOrigRect   = this.GetRect();
        let nX          = aOrigRect[0];
        let nY          = aOrigRect[1];
        let nWidth      = (aOrigRect[2] - aOrigRect[0]);
        let nHeight     = (aOrigRect[3] - aOrigRect[1]);

        let aRegions = [[
            [nX, nY],
            [nX + nWidth, nY],
            [nX + nWidth, nY + nHeight],
            [nX, nY + nHeight]
        ]];

        oGraphicsPDF.DrawLockObjectRect(this.Lock.Get_Type(), aRegions);
    };
	CBaseField.prototype.DrawFromTextBox = function(pdfGraphics, textBoxGraphics, pageIndex) {
		this.Draw(pdfGraphics, textBoxGraphics);
	};
    CBaseField.prototype.GetParent = function() {
        return this._parent;
    };
    CBaseField.prototype.GetTopParent = function() {
        if (this._parent) 
        {
            if (this._parent._parent)
                return this._parent.GetTopParent();
            else
                return this._parent;
        }

        return null;
    };

    CBaseField.prototype.SetTextColor = function(aColor) {
        function applyColorToContent(content, color) {
            if (content) {
                content.SetApplyToAll(true);
                content.AddToParagraph(new ParaTextPr({Color: color}));
                content.SetApplyToAll(false);
            }
        }

        AscCommon.History.Add(new CChangesPDFFormTextColor(this, this._textColor, aColor));
        this._textColor = aColor;
        
        var oRGB = this.GetRGBColor(aColor);
        var oColor = new AscCommonWord.CDocumentColor(oRGB.r, oRGB.g, oRGB.b, false);
    
        applyColorToContent(this.content, oColor);
        applyColorToContent(this.contentFormat, oColor);
        
        this.SetWasChanged(true);
        this.SetNeedRecalc(true);
    };
    CBaseField.prototype.GetTextColor = function() {
        return this._textColor;
    };
    CBaseField.prototype.SetTextFont = function(sFontName) {
        if (typeof(sFontName) !== "string" && sFontName == "")
            return;

        this._textFont = sFontName;
        this.SetWasChanged(true);
        this.AddToRedraw();
    };
    CBaseField.prototype.SetFontKey = function(sKey) {
        this._fontKey = sKey;
    };
    CBaseField.prototype.GetFontKey = function() {
        return this._fontKey;
    };
    CBaseField.prototype.SetTextFontActual = function(sFontName) {
        if (typeof(sFontName) !== "string" && sFontName == "")
            return;
        
        AscCommon.History.Add(new CChangesPDFFormTextFont(this, this._textFontActual, sFontName));

        if (false == Asc.editor.getDocumentRenderer().IsOpenFormsInProgress) {
            this._textFont = sFontName;
        }
        
        this._textFontActual = sFontName;

        if (this.content)
			this.content.SetFont(sFontName);
		
		if (this.contentFormat)
			this.contentFormat.SetFont(sFontName);
        
        this.SetWasChanged(true);
        this.SetNeedRecalc();
    };
    CBaseField.prototype.GetTextFontActual = function() {
        return this._textFontActual;
    };
    CBaseField.prototype.GetTextFont = function() {
        return this._textFont;
    };
    CBaseField.prototype.SetFontStyle = function(oStyle) {
        this._fontStyle = oStyle;

        if (this.content) {
            this.content.SetBold(oStyle.bold);
        }
        if (this.contentFormat) {
            this.contentFormat.SetBold(oStyle.bold);
        }
        if (this.content) {
            this.content.SetItalic(oStyle.italic);
        }
        if (this.contentFormat) {
            this.contentFormat.SetItalic(oStyle.italic);
        }
    };
    CBaseField.prototype.GetAlign = function() {};
    CBaseField.prototype.GetFontStyle = function() {
        return this._fontStyle;
    };
    CBaseField.prototype.SetTextSize = function(nSize) {
        AscCommon.History.Add(new CChangesPDFFormTextSize(this, this._textSize, nSize));

        this._textSize = nSize;
        
        if (nSize != 0) {
            if (this.content) {
                this.content.SetFontSize(nSize);
            }
            if (this.contentFormat) {
                this.contentFormat.SetFontSize(nSize);
            }
        }
        
        this.SetWasChanged(true);
        this.SetNeedRecalc();
    };
    CBaseField.prototype.GetTextSize = function() {
        return this._textSize;
    };
    CBaseField.prototype.SetRect = function(aOrigRect) {
        AscCommon.History.Add(new CChangesPDFFormRect(this, this.GetRect(), aOrigRect));

        this._origRect = aOrigRect;
        this.SetWasChanged(true);
        this.SetNeedRecalc(true);

        if (this.IsEditMode()) {
            let oShape = this.GetEditShape();

            if (aOrigRect) {
                let offX = aOrigRect[0] * g_dKoef_pt_to_mm;
                let offY = aOrigRect[1] * g_dKoef_pt_to_mm;
                let extX = (aOrigRect[2] - aOrigRect[0]) * g_dKoef_pt_to_mm;
                let extY = (aOrigRect[3] - aOrigRect[1]) * g_dKoef_pt_to_mm;

                let oXfrm = oShape.getXfrm();
                oXfrm.setExtX(extX);
                oXfrm.setExtY(extY);
                oXfrm.setOffX(offX);
                oXfrm.setOffY(offY);
            }
        }

        if (this.GetType() == AscPDF.FIELD_TYPES.text && this.IsComb()) {
            function updateMeasure(widget) {
                widget.content.GetElement(0).Content.forEach(function(run) {
                    run.RecalcInfo.Measure = true;
                });
                widget.contentFormat.GetElement(0).Content.forEach(function(run) {
                    run.RecalcInfo.Measure = true;
                });
    
                widget.SetNeedRecalc(true);
                widget.SetWasChanged(true);
            };
        
            if (this.IsWidget()) {
                updateMeasure(this);
            } else {
                this.GetAllWidgets().forEach(updateMeasure);
            }
        }
        else if (this.GetType() == AscPDF.FIELD_TYPES.button) {
            this.SetNeedUpdateImage(true);
        }

        this.CalculateContentClipRect();
    };
    CBaseField.prototype.CalculateContentClipRect = function() {};
    CBaseField.prototype.SetPosition = function(x, y) {
        let nExtX = this.GetWidth();
        let nExtY = this.GetHeight();

        let aNewRect = [x, y, x + nExtX, y + nExtY];

        this.SetRect(aNewRect);
    };
    CBaseField.prototype.GetOrigRect = function() {
        return this._origRect;
    };
    CBaseField.prototype.GetRect = function() {
        return this.GetOrigRect();
    };
    CBaseField.prototype.GetWidth = function() {
        let aRect = this.GetRect();
        return aRect[2] - aRect[0];
    };
    CBaseField.prototype.GetHeight = function() {
        let aRect = this.GetRect();
        return aRect[3] - aRect[1];
    };

    // common triggers
    CBaseField.prototype.onMouseEnter = function() {
        this.AddActionsToQueue(AscPDF.FORMS_TRIGGERS_TYPES.MouseEnter);
    };
    CBaseField.prototype.onMouseExit = function() {
        this.AddActionsToQueue(AscPDF.FORMS_TRIGGERS_TYPES.MouseExit);
    };
    CBaseField.prototype.onFocus = function() {
        this.AddActionsToQueue(AscPDF.FORMS_TRIGGERS_TYPES.OnFocus);
    };
    CBaseField.prototype.onBlur = function() {
        this.AddActionsToQueue(AscPDF.FORMS_TRIGGERS_TYPES.OnBlur);
    };
    CBaseField.prototype.onMouseUp = function() {
        this.AddActionsToQueue(AscPDF.FORMS_TRIGGERS_TYPES.MouseUp);
    };
    /**
	 * Escape from form.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 */
    CBaseField.prototype.Blur = function() {
        let oDoc = this.GetDocument();
        oDoc.SetGlobalHistory();

        this.SetInForm(false);

        if (this.content && this.content.IsSelectionUse()) {
            this.content.RemoveSelection();
        }
        
        if (oDoc.activeForm == this) {
            oDoc.activeForm = null;
            this.onBlur();
        }
    };
    // export
    CBaseField.prototype["getType"] = function() {
        return this.type;
    };
    CBaseField.prototype["getPage"] = function() {
        return this.GetPage();
    };
    CBaseField.prototype["getPagePos"] = function() {
        let aOrigRect = this.GetOrigRect();
        return {
            "x" : aOrigRect[0],
            "y" : aOrigRect[1],
            "w" : (aOrigRect[2] - aOrigRect[0]),
            "h" : (aOrigRect[3] - aOrigRect[1])
        };
    };
    CBaseField.prototype.WriteToBinaryBase = function(memory) {
        // type
        memory.WriteByte(this.GetType());

        // apidx
        memory.WriteLong(this.GetApIdx());

        // annont flags
        let bHidden      = false;
        let bPrint       = false;
        let bNoView      = false;
        let ToggleNoView = false;
        let locked       = false;
        let lockedC      = false;
        let noZoom       = false;
        let noRotate     = false;

        let nDisplayType = this.GetDisplay();
        if (nDisplayType == 1) {
            bHidden = true;
        }
        else if (nDisplayType == 0 || nDisplayType == 3) {
            bPrint = true;
            if (nDisplayType == 3) {
                bNoView = true;
            }
        }
        let annotFlags = (bHidden << 1) |
        (bPrint << 2) |
        (noZoom << 3) |
        (noRotate << 4) |
        (bNoView << 5) |
        (locked << 7) |
        (ToggleNoView << 8) |
        (lockedC << 9);

        memory.WriteLong(annotFlags);

        // page
        memory.WriteLong(this.GetOriginPage());

        // rect
        let aOrigRect = this.GetOrigRect();
        memory.WriteDouble(aOrigRect[0]); // x1
        memory.WriteDouble(aOrigRect[1]); // y1
        memory.WriteDouble(aOrigRect[2]); // x2
        memory.WriteDouble(aOrigRect[3]); // y2

        // pos for flags
        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);
        
        annotFlags = 0;
        let nBorder = this.GetBorderStyle();
        let nBorderW = this.GetBorderWidth();
        if (nBorder != null || nBorderW != null) {
            annotFlags |= (1 << 4);
            memory.WriteByte(nBorder);
            memory.WriteDouble(nBorderW);
            if (nBorder == 2) {
                memory.WriteLong(1); memory.WriteDouble(3);
            }
        }
        
        // write flags
        let nEndPos = memory.GetCurPosition();
        memory.Seek(nStartPos);
        memory.WriteLong(annotFlags);
        memory.Seek(nEndPos);
    };
    CBaseField.prototype.CheckWidgetFlags = function(memory) {
        let oParent = this.GetParent();
        if (oParent && oParent.IsAllKidsWidgets()) {
            let nCurPos = memory.GetCurPosition();
            memory.Seek(memory.posForWidgetFlags);
            memory.WriteLong(-1);
            memory.Seek(nCurPos);
        }
    };
    CBaseField.prototype.GetFontSizeAP = function() {
        return 0;
    };
    CBaseField.prototype.SetEditMode = function(bEdit) {
        if (bEdit == false) {
            this.editShape = null;
            this.AddToRedraw();
            return;
        }

        AscCommon.History.StartNoHistoryMode();
        let oPdfShape = new AscPDF.CPdfShape();
        let aOrigRect = this.GetRect();
        let aRectMM = aOrigRect ? aOrigRect.map(function(measure) {
            return measure * g_dKoef_pt_to_mm;
        }) : [];

        let nOffX = aRectMM[0];
        let nOffY = aRectMM[1];
        let nExtX = aRectMM[2] - aRectMM[0];
        let nExtY = aRectMM[3] - aRectMM[1];

        oPdfShape.setSpPr(new AscFormat.CSpPr());
        oPdfShape.spPr.setLn(new AscFormat.CLn());
        oPdfShape.setFLocksText(true);
        let oFill = AscFormat.CreateSolidFillRGBA(0, 0, 0, 0);
        oFill.setTransparent(0);

        oPdfShape.spPr.ln.setFill(oFill);
        oPdfShape.spPr.setFill(oFill);
        oPdfShape.spPr.setParent(oPdfShape);
        oPdfShape.spPr.setXfrm(new AscFormat.CXfrm());
        oPdfShape.spPr.xfrm.setParent(oPdfShape.spPr);
        
        oPdfShape.spPr.xfrm.setOffX(nOffX);
        oPdfShape.spPr.xfrm.setOffY(nOffY);
        oPdfShape.spPr.xfrm.setExtX(nExtX);
        oPdfShape.spPr.xfrm.setExtY(nExtY);
        
        oPdfShape.spPr.setGeometry(AscFormat.CreateGeometry("rect"));

        oPdfShape.createTextBody();
        oPdfShape.txBody.bodyPr.setInsets(0.5,0.5,0.5,0.5);
        oPdfShape.txBody.bodyPr.horzOverflow = AscFormat.nHOTClip;
        oPdfShape.txBody.bodyPr.vertOverflow = AscFormat.nVOTClip;
        oPdfShape.setVerticalAlign(AscFormat.VERTICAL_ANCHOR_TYPE_CENTER);

        let oContent = oPdfShape.GetDocContent();
        let oPara = oContent.GetElement(0);
        let oRun = oPara.GetElement(0);
        oRun.AddText(this.GetFullName());
        oPara.SetParagraphAlign(AscCommon.align_Center);
        oPara.SetApplyToAll(true);
        oPara.Add(new ParaTextPr({HighlightColor: AscFormat.CreateUniColorRGB(0, 0, 0)}));
        oPara.SetApplyToAll(false);

        oPdfShape.setStyle(AscFormat.CreateDefaultShapeStyle());
        oPdfShape.setBDeleted(false);
        oPdfShape.recalculate();
        
        oPdfShape.SetEditField(this);
        this.editShape = oPdfShape;
        AscCommon.History.EndNoHistoryMode();

        this.AddToRedraw();
    };
    CBaseField.prototype.IsEditMode = function() {
        return !!this.editShape;
    };
    CBaseField.prototype.GetEditShape = function() {
        return this.editShape;
    };
    CBaseField.prototype.WriteToBinaryBase2 = function(memory) {
        // font name
        let sFontName = this.GetTextFont() || "";
        memory.WriteString(sFontName);

        // text size
        let nFontSize = this.GetTextSize() || 0;
        memory.WriteDouble(nFontSize);

        // text size for ap
        memory.WriteDouble(this.GetFontSizeAP(this.content));

        // font style
        let oStyle = this.GetFontStyle();
        let nStyle = 0;
        if (oStyle.bold) {
            nStyle |= (1 << 0);
        }
        if (oStyle.italic) {
            nStyle |= (1 << 1);
        }
        memory.WriteLong(nStyle);

        // text color
        let aTextColor = this.GetTextColor();
        if (aTextColor && aTextColor.length != 0) {
            memory.WriteLong(aTextColor.length);

            for (let i = 0; i < aTextColor.length; i++) {
                memory.WriteDouble(aTextColor[i]);
            }
        }

        // align 
        if ([AscPDF.FIELD_TYPES.text, AscPDF.FIELD_TYPES.combobox, AscPDF.FIELD_TYPES.listbox].includes(this.GetType())) {
            let nAlignType = this.GetAlign();
            memory.WriteByte(nAlignType);
        }

        // сюда пойдут 1ые флаги полей
        memory.widgetFlags   = 0;
        memory.posForWidgetFlags  = memory.GetCurPosition();
        memory.Skip(4);
        
        if (this.IsReadOnly(false)) {
            memory.widgetFlags |= (1 << 0);
        }

        if (this.IsRequired(false)) {
            memory.widgetFlags |= (1 << 1);
        }

        if (this.IsNoExport(false)) {
            memory.widgetFlags |= (1 << 2);
        }

        let oParent = this.GetParent();
        let nFieldType = this.GetType();
        let bWriteType = !oParent || !oParent.IsAllKidsWidgets();

        if (bWriteType) {
            switch (nFieldType) {
                case AscPDF.FIELD_TYPES.radiobutton: {
                    memory.widgetFlags |= (1 << 15);
                    break;
                }
                case AscPDF.FIELD_TYPES.button: {
                    memory.widgetFlags |= (1 << 16);
                    break;
                }
                case AscPDF.FIELD_TYPES.combobox: {
                    memory.widgetFlags |= (1 << 17);
                    break;
                }
            }    
        }

        // сюда пойдут 2ые флаги полей
        memory.fieldDataFlags   = 0;
        memory.posForFieldDataFlags  = memory.GetCurPosition();
        memory.Skip(4);
        
        //
        // username
        //

        //
        // default style
        //

        let sFontKey = this.GetFontKey();
        if (sFontKey) {
            memory.fieldDataFlags |= (1 << 2);
            memory.WriteString(sFontKey);
        }

        // highlight
        let nHighlightType = this.GetHighlight();
        if (nHighlightType != null) {
            memory.fieldDataFlags |= (1 << 3);
            memory.WriteByte(nHighlightType);
        }

        let aBorderColor = this.GetBorderColor();
        if (aBorderColor && aBorderColor.length != 0) {
            memory.fieldDataFlags |= (1 << 5);
            memory.WriteLong(aBorderColor.length);
            for (let i = 0; i < aBorderColor.length; i++) {
                memory.WriteDouble(aBorderColor[i]);
            }
        }

        //
        // rotate
        //

        let aBgColor = this.GetBackgroundColor();
        if (aBgColor && aBgColor.length != 0) {
            memory.fieldDataFlags |= (1 << 7);
            memory.WriteLong(aBgColor.length);
            for (let i = 0; i < aBgColor.length; i++) {
                memory.WriteDouble(aBgColor[i]);;
            }
        }

        // default value
        let defValue = this.GetDefaultValue(false);
        if (defValue != null) {
            memory.fieldDataFlags |= (1 << 8);
            memory.WriteString(defValue);
        }

        // parent
        if (oParent != null) {
            memory.fieldDataFlags |= (1 << 17);
            memory.WriteLong(oParent.GetApIdx());
        }

        // partial name
        let sName = this.GetPartialName();
        if (sName != null) {
            memory.fieldDataFlags |= (1 << 18);
            memory.WriteString(sName);
        }

        // meta data
        let oMeta = this.GetMeta();
        if (oMeta != null) {
            memory.fieldDataFlags |= (1 << 20);
            memory.WriteString(JSON.stringify(oMeta));
        }

        // actions
        let aActions = this.GetListActions(false);
        memory.WriteLong(aActions.length);
        
        for (let i = 0; i < aActions.length; i++) {
            aActions[i].WriteToBinary(memory);
        }
    };
    CBaseField.prototype.WriteToBinaryAsParent = function(memory) {
        memory.WriteLong(this.GetApIdx());
        // pos for flags
        let nStartPos   = memory.GetCurPosition();
        let nFlags      = 0;
        memory.Skip(4);

        // partial name
        let sName = this.GetPartialName();
        if (sName != null) {
            nFlags |= (1 << 0);
            memory.WriteString(sName);
        }

        // value
        let value = this.GetParentValue();
        if (value != null && Array.isArray(value) == false) {
            nFlags |= (1 << 1);
            memory.WriteString(value);
        }

        // default value
        let defValue = this.GetDefaultValue();
        if (defValue != null) {
            nFlags |= (1 << 2);
            memory.WriteString(defValue);
        }

        // combobox/listbox
        let curIdxs = [];
        if ([AscPDF.FIELD_TYPES.combobox, AscPDF.FIELD_TYPES.listbox].includes(this.GetType())) {
            curIdxs = this.GetParentCurIdxs();
        }
        if (curIdxs.length > 0) {
            nFlags |= (1 << 3);
            memory.WriteLong(curIdxs.length);
            for (let i = 0; i < curIdxs.length; i++) {
                memory.WriteLong(curIdxs[i]);
            }
        }

        // parent
        let oParent = this.GetParent();
        if (oParent != null) {
            nFlags |= (1 << 4);
            memory.WriteLong(oParent.GetApIdx());
        }

        if (value != null && Array.isArray(value) == true) {
            // флаг что значение - это массив
            nFlags |= (1 << 5);
            memory.WriteLong(value.length);
            for (let i = 0; i < value.length; i++) {
                memory.WriteString(value[i]);
            }
        }

        // combobox/listbox
        let aOptions = [];
        if ([AscPDF.FIELD_TYPES.combobox, AscPDF.FIELD_TYPES.listbox, AscPDF.FIELD_TYPES.checkbox, AscPDF.FIELD_TYPES.radiobutton].includes(this.GetType())) {
            aOptions = this.GetOptions(false);
            if (aOptions) {
                nFlags |= (1 << 6);

                memory.WriteLong(aOptions.length);
                for (let i = 0; i < aOptions.length; i++) {
                    memory.WriteString(Array.isArray(aOptions[i]) ? aOptions[i][1] : "");
                    memory.WriteString(Array.isArray(aOptions[i]) ? aOptions[i][0] : aOptions[i]);
                }
            }
        }

        // флаги полей
        nFlags |= (1 << 7);

        let nFieldType = this.GetType();
        let bWriteType = this.IsAllKidsWidgets();

        let nWidgetFlags = 0;
        if (bWriteType) {
            switch (nFieldType) {
                case AscPDF.FIELD_TYPES.radiobutton: {
                    nWidgetFlags |= (1 << 15);
                    break;
                }
                case AscPDF.FIELD_TYPES.button: {
                    nWidgetFlags |= (1 << 16);
                    break;
                }
                case AscPDF.FIELD_TYPES.combobox: {
                    nWidgetFlags |= (1 << 17);
                    break;
                }
            }    
        }

        if (this.IsReadOnly && this.IsReadOnly()) {
            nWidgetFlags |= (1 << 0);
        }
        if (this.IsRequired && this.IsRequired()) {
            nWidgetFlags |= (1 << 1);
        }
        if (this.IsNoExport && this.IsNoExport()) {
            nWidgetFlags |= (1 << 2);
        }
        if (this.IsMultiline && this.IsMultiline()) {
            nWidgetFlags |= (1 << 12);
        }
        if (this.IsPassword && this.IsPassword()) {
            nWidgetFlags |= (1 << 13);
        }
        if (this.IsNoToggleToOff && this.IsNoToggleToOff()) {
            nWidgetFlags |= (1 << 14);
        }
        if (this.IsEditable && this.IsEditable()) {
            nWidgetFlags |= (1 << 18);
        }
        if (this.IsFileSelect && this.IsFileSelect()) {
            nWidgetFlags |= (1 << 20);
        }
        if (this.IsMultipleSelection && this.IsMultipleSelection()) {
            nWidgetFlags |= (1 << 21);
        }
        if (this.IsDoNotSpellCheck && this.IsDoNotSpellCheck()) {
            nWidgetFlags |= (1 << 22);
        }
        if (this.IsDoNotScroll && this.IsDoNotScroll()) {
            nWidgetFlags |= (1 << 23);
        }
        if (this.IsComb && this.IsComb()) {
            nWidgetFlags |= (1 << 24);
        }
        if (this.IsRadiosInUnison && this.IsRadiosInUnison()) {
            nWidgetFlags |= (1 << 25);
        }
        if (this.IsCommitOnSelChange && this.IsCommitOnSelChange()) {
            nWidgetFlags |= (1 << 26);
        }
        memory.WriteLong(nWidgetFlags);

        // actions
        nFlags |= (1 << 8);
        let aActions = this.GetListActions();
        memory.WriteLong(aActions.length);
        
        for (let i = 0; i < aActions.length; i++) {
            aActions[i].WriteToBinary(memory);
        }

        // charLimit
        if (this.GetType() == AscPDF.FIELD_TYPES.text) {
            nFlags |= (1 << 9);
            memory.WriteLong(this.GetCharLimit());
        }

        // write flags
        let nEndPos = memory.GetCurPosition();
        memory.Seek(nStartPos);
        memory.WriteLong(nFlags);
        memory.Seek(nEndPos);
    };
    CBaseField.prototype.WriteRenderToBinary = function(memory) {
        // пока только для text, combobox
        if (false == [AscPDF.FIELD_TYPES.text, AscPDF.FIELD_TYPES.combobox, AscPDF.FIELD_TYPES.listbox].includes(this.GetType())) {
            return;
        }

        // тут будет длина комманд
        let nStartPos = memory.GetCurPosition();
        memory.Skip(4);

        let oContentToDraw = this.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Format) ? this.contentFormat : this.content;
        oContentToDraw.Draw(0, memory.docRenderer);

        // запись длины комманд
        let nEndPos = memory.GetCurPosition();
        memory.Seek(nStartPos);
        memory.WriteLong(nEndPos - nStartPos);
        memory.Seek(nEndPos);
    };

    // for format

    // private methods

    function private_GetFieldAlign(sJc)
	{
		if ("left" === sJc)
			return align_Left;
		else if ("right" === sJc)
			return align_Right;
		else if ("center" === sJc)
			return align_Center;

		return undefined;
	}

    if (!window["AscPDF"])
	    window["AscPDF"] = {};
    
	window["AscPDF"].ALIGN_TYPE         = ALIGN_TYPE;
	window["AscPDF"].VALID_ROTATIONS    = VALID_ROTATIONS;
    window["AscPDF"].MAX_TEXT_SIZE      = MAX_TEXT_SIZE;
    window["AscPDF"].FONT_STRETCH       = FONT_STRETCH;
    window["AscPDF"].FONT_STYLE         = FONT_STYLE;
    window["AscPDF"].FONT_WEIGHT        = FONT_WEIGHT;
	window["AscPDF"].DEFAULT_FIELD_FONT = DEFAULT_FIELD_FONT;

    window["AscPDF"].CBaseField = CBaseField;
    
})();

