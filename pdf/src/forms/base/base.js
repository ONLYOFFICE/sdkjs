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

    let BUTTON_PRESSED = {
        r: 153,
        g: 193,
        b: 218
    };

    let BORDER_TYPES = {
        solid:      0,
        beveled:    1,
        dashed:     2,
        inset:      3,
        underline:  4
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

    let APPEARANCE_TYPE = {
        normal:     0,
        rollover:   1,
        mouseDown:  2
    };

    const CHAR_LIM_MAX = 500; // to do проверить

    // For Span attributes (start)
    let FONT_STRETCH = ["ultra-condensed", "extra-condensed", "condensed", "semi-condensed", "normal",
        "semi-expanded", "expanded", "extra-expanded", "ultra-expanded"];

    let FONT_STYLE = {
        italic: "italic",
        normal: "normal"
    }

    let FONT_WEIGHT = [100, 200, 300, 400, 500, 600, 700, 800, 900];

    let LINE_WIDTH = {
        none:   0,
        thin:   1,
        medium: 2,
        thick:  3
    }

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
    function CBaseField(sName, nType, nPage, aRect, oDoc)
    {
        this.type = nType;

        this._kids          = [];
        this._borderStyle   = BORDER_TYPES.solid;
        this._delay         = false;
        this._display       = AscPDF.Api.Objects.display["visible"];
        this._doc           = oDoc;
        this._fillColor     = undefined;
        this._bgColor       = undefined;          // prop for old versions (fillColor)
        this._hidden        = false;             // This property has been superseded by the display property and its use is discouraged.
        this._lineWidth     = LINE_WIDTH.thin;  // In older versions of this specification, this property was borderWidth
        this._borderWidth   = undefined;       
        this._name          = sName;         // partial field name
        this._page          = nPage;        // integer | array
        this._print         = true;        // This property has been superseded by the display property and its use is discouraged.
        this._readonly      = false;
        this._rect          = aRect;         // scaled rect
        this._origRect      = [];           // orig rect as in file
        this._required      = false;       // for all except button
        this._rotation      = 0;
        this._strokeColor   = null;     // In older versions of this specification, this property was borderColor. The use of borderColor is now discouraged,
                                        // although it is still valid for backward compatibility.
        this._borderColor   = undefined;
        this._submitName    = "";
        this._textColor     = [0,0,0];
        this._fgColor       = undefined;
        this._textSize      = 10; // 0 == max text size // to do
        this._userName      = ""; // It is intended to be used as tooltip text whenever the cursor enters a field. 
        //It can also be used as a user-friendly name, instead of the field name, when generating error messages.
        this._parent        = null;
        
        this._triggers = new AscPDF.CFormTriggers();

        // internal
        this._id = AscCommon.g_oIdCounter.Get_NewId();
        
        this._isWidget = aRect && aRect.length == 4 ? true : false;

        this.contentRect = {
            X: 0,
            Y: 0,
            W: 0,
            H: 0,
            Page: nPage
        }
        this._formRect = {
            X: 0,
            Y: 0,
            W: 0,
            H: 0,
            Page: nPage
        }

        this._oldContentPos = {X: 0, Y: 0, XLimit: 0, YLimit: 0};
        this._curShiftView = { // смещение, когда мы скролим, т.е. активное смещение
            x: 0,
            y: 0
        }
        this._originShiftView = { // смещение, когда значение формы применено (т.е. форма не активна)
            x: 0,
            y: 0
        }

        this._apIdx     = undefined; // индекс формы на странице в исходном файле (в массиве метода getInteractiveForms), используется для получения appearance
        this._parentIdx = undefined; // индекс родителя во время чтения

        this.needCommit             = false;
        this._needDrawHighlight     = true;
        this._needDrawHoverBorder   = false;
        this._needRecalc            = true;
        this._wasChanged            = false; // была ли изменена форма
        this._bDrawFromStream       = false; // нужно ли рисовать из стрима
        this._originView = {
            normal:     null,
            mouseDown:  null,
            rollover:   null
        }

        editor.getDocumentRenderer().ImageMap = {};
        editor.getDocumentRenderer().InitDocument = function() {return};

        this._partialName = sName;
        this.api = this.GetFormApi();
        this["api"] = this.api;
    }
    
    CBaseField.prototype.AddToChildsMap = function(nIdx) {
        this.GetDocument().AddFieldToChildsMap(this, nIdx);
    };
    /**
	 * Can or not enter text into form.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 */
    CBaseField.prototype.IsEditable = function() {
        return false;
    };
	/**
	 * Check if fonts are available, if not then we download them
	 * @returns {boolean}
	 */
	CBaseField.prototype.checkFonts = function() {
		return this._doc && 1 === this._doc.defaultFontsLoaded;
	};
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
        let nCurPage = this.GetPage();
        if (nPage == nCurPage)
            return;


        let oViewer = editor.getDocumentRenderer();
        let nCurIdxOnPage = oViewer.pagesInfo.pages[nCurPage].fields ? oViewer.pagesInfo.pages[nCurPage].fields.indexOf(this) : -1;
        if (oViewer.pagesInfo.pages[nPage]) {
            if (oViewer.pagesInfo.pages[nPage].fields == null) {
                oViewer.pagesInfo.pages[nPage].fields = [];
            }

            if (nCurIdxOnPage != -1)
                oViewer.pagesInfo.pages[nCurPage].fields.splice(nCurIdxOnPage, 1);

            if (oViewer.pagesInfo.pages[nPage].fields.indexOf(this) == -1)
                oViewer.pagesInfo.pages[nPage].fields.push(this);

            this._page = nPage;
            this.selectStartPage = nPage;
        }
    };
    CBaseField.prototype.GetPage = function() {
        return this._page;
    };

    /**
	 * Gets the child field by the specified partial name.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 * @returns {?CBaseField}
	 */
    CBaseField.prototype.AddKid = function(oField) {
        this._kids.push(oField);
        oField._parent = this;
    };
    CBaseField.prototype.GetKids = function() {
        return this._kids;
    };

    CBaseField.prototype.Recalculate = function() {};
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
            oField._parent = null;
            return true;
        }

        return false;
    };

    CBaseField.prototype.GetDocContent = function() {
        return this.content;
    };

    CBaseField.prototype.getFormRelRect = function()
    {
        return this.contentRect;
    };
    CBaseField.prototype.getFormRect = function()
    {
        return this._formRect;
    };
    
    CBaseField.prototype.IntersectWithRect = function(X, Y, W, H)
    {
        let oViewer     = editor.getDocumentRenderer();
        var arrRects    = [];
        var oBounds     = this.getFormRelRect();

        let nPage = this.GetPage();
        let oPage = oViewer.pageDetector.pages.find(function(page) {
            return page.num == nPage;
        });
        
        let nPageIndX = oPage.x / AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_pix_to_mm;
        let nPageIndY = oPage.y / AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_pix_to_mm;

        var nLeft   = Math.max(X, oBounds.X) + nPageIndX / oViewer.zoom;
        var nRight  = Math.min(X + W, oBounds.X + oBounds.W) + nPageIndX / oViewer.zoom;
        var nTop    = Math.max(Y, oBounds.Y) + nPageIndY / oViewer.zoom;
        var nBottom = Math.min(Y + H, oBounds.Y + oBounds.H) + nPageIndY / oViewer.zoom;

        if (nLeft < nRight && nTop < nBottom)
        {
            arrRects.push({
                X : nLeft,
                Y : nTop,
                W : nRight - nLeft,
                H : nBottom - nTop
            });
        }

        return arrRects;
    };
    CBaseField.prototype.GetFullName = function() {
        if (this._parent)
        {
            if (this._partialName != undefined)
                return `${this._parent.GetFullName()}.${this._partialName}`
            else
                return this._parent.GetFullName();
        }

        return this._partialName ? this._partialName : "";
    };
    CBaseField.prototype.GetPartialName = function() {
        return this._partialName;
    };
    /**
	 * Sets the action of the field for a given trigger.
     * Note: This method will overwrite any action already defined for the chosen trigger.
	 * @memberof CBaseField
     * @param {number} nTriggerType - A string that sets the trigger for the action. (FORMS_TRIGGERS_TYPES)
	 * @param {Array} aActionsInfo - array with actions info for specified trigger. (info from openForms method)
     * @typeofeditors ["PDF"]
	 */
    CBaseField.prototype.SetActionsOnOpen = function(nTriggerType, aActionsInfo) {
        let oDocument = this.GetDocument();
        let aActions = [];
        for (let i = 0; i < aActionsInfo.length; i++) {
            let oAction;
            let aFields = [];
            switch (aActionsInfo[i]["S"]) {
                case AscPDF.ACTIONS_TYPES.JavaScript:
                    oAction = new AscPDF.CActionRunScript(aActionsInfo[i]["JS"]);
                    aActions.push(oAction);
                    oAction.SetField(this);
                    break;
                case AscPDF.ACTIONS_TYPES.ResetForm:
                    oAction = new AscPDF.CActionReset(aFields, aActionsInfo[i]["Fields"]);
                    aActions.push(oAction);
                    oAction.SetField(this);
                    break;
                case AscPDF.ACTIONS_TYPES.URI:
                    oAction = new AscPDF.CActionURI(aActionsInfo[i]["URI"]);
                    aActions.push(oAction);
                    oAction.SetField(this);
                    break;
                case AscPDF.ACTIONS_TYPES.HideShow:
                    oAction = new AscPDF.CActionHideShow(Boolean(aActionsInfo[i]["H"]), aActionsInfo[i]["T"]);
                    aActions.push(oAction);
                    oAction.SetField(this);
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
                    oAction.SetField(this);
                    break;
                case AscPDF.ACTIONS_TYPES.Named:
                    oAction = new AscPDF.CActionNamed(AscPDF.CActionNamed.GetInternalType(aActionsInfo[i]["N"]));
                    aActions.push(oAction);
                    oAction.SetField(this);
                    break;
            }
        }

        switch (nTriggerType) {
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseUp:
                this._triggers.MouseUp = new AscPDF.CFormTrigger(nTriggerType, aActions);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseDown:
                this._triggers.MouseDown = new AscPDF.CFormTrigger(nTriggerType, aActions);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseEnter:
                this._triggers.MouseEnter = new AscPDF.CFormTrigger(nTriggerType, aActions);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseExit:
                this._triggers.MouseExit = new AscPDF.CFormTrigger(nTriggerType, aActions);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.OnFocus:
                this._triggers.OnFocus = new AscPDF.CFormTrigger(nTriggerType, aActions);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.OnBlur:
                this._triggers.OnBlur = new AscPDF.CFormTrigger(nTriggerType, aActions);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Keystroke:
                this._triggers.Keystroke = new AscPDF.CFormTrigger(nTriggerType, aActions);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Validate:
                this._triggers.Validate = new AscPDF.CFormTrigger(nTriggerType, aActions);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Calculate:
                this._triggers.Calculate = new AscPDF.CFormTrigger(nTriggerType, aActions);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Format:
                this._triggers.Format = new AscPDF.CFormTrigger(nTriggerType, aActions);
                break;
        }

        aActions.forEach(function(oAction) {
            oAction.SetTrigger(nTriggerType);
        });

        return aActions;
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
        let oCalcInfo = this.GetDocument().GetCalculateInfo();
        let oAction = new AscPDF.CActionRunScript(sScript);
        oAction.SetField(this);

        switch (nTriggerType) {
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseUp:
                this._triggers.MouseUp = new AscPDF.CFormTrigger(nTriggerType, [oAction]);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseDown:
                this._triggers.MouseDown = new AscPDF.CFormTrigger(nTriggerType, [oAction]);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseEnter:
                this._triggers.MouseEnter = new AscPDF.CFormTrigger(nTriggerType, [oAction]);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseExit:
                this._triggers.MouseExit = new AscPDF.CFormTrigger(nTriggerType, [oAction]);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.OnFocus:
                this._triggers.OnFocus = new AscPDF.CFormTrigger(nTriggerType, [oAction]);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.OnBlur:
                this._triggers.OnBlur = new AscPDF.CFormTrigger(nTriggerType, [oAction]);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Keystroke:
                this._triggers.Keystroke = new AscPDF.CFormTrigger(nTriggerType, [oAction]);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Validate:
                this._triggers.Validate = new AscPDF.CFormTrigger(nTriggerType, [oAction]);
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Calculate:
                this._triggers.Calculate = new AscPDF.CFormTrigger(nTriggerType, [oAction]);
                oCalcInfo.RemoveFieldFromOrder(this.GetFullName());
                oCalcInfo.AddFieldToOrder(this.GetFullName());
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Format:
                this._triggers.Format = new AscPDF.CFormTrigger(nTriggerType, [oAction]);
                break;
        }
    };

    /**
	 * Gets the JavaScript action of the field for a given trigger.
	 * @memberof CBaseField
     * @param {number} nType - A string that sets the trigger for the action. (FORMS_TRIGGERS_TYPES)
	 * @typeofeditors ["PDF"]
     * @returns {CFormTrigger}
	 */
    CBaseField.prototype.GetTrigger = function(nType) {
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
    CBaseField.prototype.GetApiValue = function(bInherit) {
        let oParent = this.GetParent();
        if (oParent == null && this._value == null)
            return "";
        else if (bInherit === false || (this._value != null && this.GetPartialName() != null)) {
            return this._value;
        }
        
        if (oParent)
            return oParent.GetApiValue();
    };
    /**
	 * Sets api value of form.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 */
    CBaseField.prototype.SetApiValue = function(value) {
        let oParent = this.GetParent();
        if (oParent && this.IsWidget())
            oParent.SetApiValue(value);
        else
            this._value = value;
    };

    /**
     * Does the actions setted for specifed trigger type.
	 * @memberof CBaseField
     * @param {number} nType - trigger type (FORMS_TRIGGERS_TYPES)
	 * @typeofeditors ["PDF"]
     * @returns {canvas}
	 */
    CBaseField.prototype.AddActionsToQueue = function(nType) {
        let oDoc            = this.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();
        let oTrigger        = this.GetTrigger(nType);
        
        if (oTrigger && oTrigger.Actions.length > 0) {
            oActionsQueue.AddActions(oTrigger.Actions);
            oActionsQueue.Start();
        }
    };

    CBaseField.prototype.DrawHighlight = function(oCtx) {
        if (this.IsHidden() == true)
            return;

        let oViewer     = editor.getDocumentRenderer();
        let nScale      = AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom * (96 / oViewer.file.pages[this.GetPage()].Dpi);
        let aBgColor    = this.GetBackgroundColor();

        let xCenter = oViewer.width >> 1;
        if (oViewer.documentWidth > oViewer.width)
        {
            xCenter = (oViewer.documentWidth >> 1) - (oViewer.scrollX) >> 0;
        }
        let yPos    = oViewer.scrollY >> 0;
        let page    = oViewer.drawingPages[this.GetPage()];
        let w       = (page.W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
        let h       = (page.H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
        let indLeft = ((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - (w >> 1);
        let indTop  = ((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
        
        let aOrigRect = this.GetOrigRect();
        let X = aOrigRect[0] * nScale + indLeft;
        let Y = aOrigRect[1] * nScale + indTop;
        let W = (aOrigRect[2] - aOrigRect[0]) * nScale;
        let H = (aOrigRect[3] - aOrigRect[1]) * nScale;

        if (null == aBgColor)
            oCtx.globalAlpha = 0.8;

        oCtx.globalCompositeOperation = "destination-over";
        if (this.IsNeedDrawFromStream())
            oCtx.globalCompositeOperation = "multiply";
        
        if (this.GetType() == AscPDF.FIELD_TYPES.radiobutton && this._chStyle == AscPDF.CHECKBOX_STYLES.circle) {
            oCtx.beginPath();
            oCtx.fillStyle = `rgb(${FIELDS_HIGHLIGHT.r}, ${FIELDS_HIGHLIGHT.g}, ${FIELDS_HIGHLIGHT.b})`;
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
            oCtx.fillStyle = `rgb(${FIELDS_HIGHLIGHT.r}, ${FIELDS_HIGHLIGHT.g}, ${FIELDS_HIGHLIGHT.b})`;
            oCtx.fillRect(X, Y, W, H);
            oCtx.closePath();
        }
        oCtx.globalCompositeOperation = "source-over";
    };
    CBaseField.prototype.DrawBorders = function(oGraphicsPDF) {
        let oViewer     = editor.getDocumentRenderer();
        let aOringRect  = this.GetOrigRect();
        let aBgColor    = this.IsNeedDrawHighlight() == false ? (this.GetBackgroundColor() || [1]) : [1];
        let oBgRGBColor = this.GetRGBColor(aBgColor);

        if (aBgColor && aBgColor.length != 0)
            oBgRGBColor = AscPDF.MakeColorMoreGray(oBgRGBColor, 50);
        
        let nLineWidth = this._lineWidth;

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

        if (this.GetType() == AscPDF.FIELD_TYPES.radiobutton && this._chStyle == AscPDF.CHECKBOX_STYLES.circle) {
            // выставляем в центр круга
            let centerX = X + nWidth / 2;
            let centerY = Y + nHeight / 2;
            let nRadius = Math.min(nWidth / 2, nHeight / 2);

            // отрисовка
            switch (this._borderStyle) {
                case BORDER_TYPES.solid:
                case BORDER_TYPES.underline:
                    if (color == null)
                        break;
                    oGraphicsPDF.SetLineDash([]);
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.Arc(centerX, centerY, nRadius, 0, 2 * Math.PI, false);
                    oGraphicsPDF.Stroke();
                    break;
                case BORDER_TYPES.beveled:
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
                case BORDER_TYPES.dashed:
                    if (color == null)
                        break;

                    oGraphicsPDF.SetLineDash([5 * oViewer.zoom]);
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.Arc(centerX, centerY, nRadius, 0, 2 * Math.PI, false);
                    oGraphicsPDF.Stroke();
                    break;
                case BORDER_TYPES.inset:
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
            switch (this._borderStyle) {
                case BORDER_TYPES.solid:
                    if (color == null)
                        break;
                    
                    oGraphicsPDF.SetLineDash([]);
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.Rect(X, Y, nWidth, nHeight);
                    oGraphicsPDF.Stroke();

                    break;
                case BORDER_TYPES.beveled:
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
                case BORDER_TYPES.dashed:
                    if (color == null)
                        break;

                    oGraphicsPDF.SetLineDash([5 * oViewer.zoom]);
                    oGraphicsPDF.BeginPath();
                    oGraphicsPDF.Rect(X, Y, nWidth, nHeight);
                    oGraphicsPDF.Stroke();
                    break;
                case BORDER_TYPES.inset:
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
                case BORDER_TYPES.underline:
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
        if (this.GetType() == AscPDF.FIELD_TYPES.button && this.IsPressed() && this.GetHighlight() == AscPDF.BUTTON_HIGHLIGHT_TYPES.push && this._images.mouseDown == undefined) {
            switch (this._borderStyle) {
                case BORDER_TYPES.solid:
                case BORDER_TYPES.dashed:
                case BORDER_TYPES.underline: {
                    
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
        if ((this.GetType() == AscPDF.FIELD_TYPES.text && this.IsComb() == true) && this._borderColor != null && (this.GetBorderStyle() == BORDER_TYPES.solid || this.GetBorderStyle() == BORDER_TYPES.dashed)) {
            let nCombWidth = (nWidth / this._charLimit);
            let nIndentX = nCombWidth;
            
            for (let i = 0; i < this._charLimit - 1; i++) {
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
                r: aInternalColor[0] * 255,
                g: aInternalColor[0] * 255,
                b: aInternalColor[0] * 255
            }
        }
        else if (aInternalColor.length == 3) {
            oColor = {
                r: aInternalColor[0] * 255,
                g: aInternalColor[1] * 255,
                b: aInternalColor[2] * 255
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
    /**
	 * Creates new history point for cur field.
	 * @memberof CComboBoxField
	 * @typeofeditors ["PDF"]
     * @param {boolean} isCanUnion - is it possible to union a point with others
     * @returns {object}
	 */
    CBaseField.prototype.CreateNewHistoryPoint = function(isCanUnion) {
        if (isCanUnion == null)
            isCanUnion = false;

        if (AscCommon.History.IsOn() == false)
            AscCommon.History.TurnOn();
        AscCommon.History.Create_NewPoint();
        AscCommon.History.SetAdditionalFormFilling(this);
        AscCommon.History.Points[AscCommon.History.Points.length - 1].Additional.CanUnion = isCanUnion;
    };
    CBaseField.prototype.DrawSelected = function() {
        return;
        /*
        let oViewer     = editor.getDocumentRenderer();
        let nScale      = AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom;

        let X       = this._pagePos.x * nScale;
        let Y       = this._pagePos.y * nScale;
        let nWidth  = this._pagePos.w * nScale;
        let nHeight = this._pagePos.h * nScale;

        oCtx.globalAlpha = 1;
        oCtx.setLineDash([]);
        oCtx.strokeStyle = "rgb(0, 0, 0)";
        oCtx.lineWidth = Math.max(nScale, 1);
        
        oCtx.beginPath();
        oCtx.rect(X, Y, nWidth, nHeight);
        oCtx.stroke();
        oCtx.closePath();
        */
    };
    
    CBaseField.prototype.Get_Id = function() {
        return this._id;
    };
    CBaseField.prototype.SetNeedRecalc = function(bRecalc, bSkipAddToRedraw) {
        if (bRecalc == false) {
            this._needRecalc = !this.checkFonts();
        }
        else {
            this._needRecalc = true;
            if (bSkipAddToRedraw != true)
                this.AddToRedraw();
        }
    };
    CBaseField.prototype.IsNeedRecalc = function() {
        return this._needRecalc;
    };
    CBaseField.prototype.SetWasChanged = function(isChanged) {
        let oViewer = editor.getDocumentRenderer();

        if (oViewer.IsOpenFormsInProgress == false) {
            this._wasChanged = isChanged;
            this.SetDrawFromStream(!isChanged);
        }
    };
    CBaseField.prototype.ClearCache = function() {
        this._originView.normal     = null;
        this._originView.mouseDown  = null;
        this._originView.rollover   = null;
    };
    CBaseField.prototype.IsChanged = function() {
        return this._wasChanged;  
    };
    CBaseField.prototype.IsNeedDrawFromStream = function() {
        return this._bDrawFromStream;
    };
    CBaseField.prototype.SetDrawFromStream = function(bFromStream) {
        this._bDrawFromStream = bFromStream;
    };
    CBaseField.prototype.SetDrawHighlight = function(bDraw) {
        this._needDrawHighlight = bDraw;
    };
    CBaseField.prototype.IsNeedDrawHighlight = function() {
        return this._needDrawHighlight;
    };

    CBaseField.prototype.AddToRedraw = function() {
        let oViewer = editor.getDocumentRenderer();
        if (oViewer.pagesInfo.pages[this.GetPage()])
            oViewer.pagesInfo.pages[this.GetPage()].needRedrawForms = true;
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
        this._readonly = bReadOnly;
    };
    
    CBaseField.prototype.SetRequired = function(bRequired) {
        if (this.GetType() != AscPDF.FIELD_TYPES.button)
            this._required = bRequired;
    };
    CBaseField.prototype.SetBorderColor = function(aColor) {
        this._strokeColor = this._borderColor = aColor;
    };
    CBaseField.prototype.SetBackgroundColor = function(aColor) {
        this._fillColor = this._bgColor = aColor;
    };
    CBaseField.prototype.GetBackgroundColor = function() {
        return this._fillColor;
    };
    CBaseField.prototype.IsHidden = function() {
        let nType = this.GetDisplay();
        if (nType == window["AscPDF"].Api.Objects.display["hidden"] || nType == window["AscPDF"].Api.Objects.display["noView"])
            return true;

        return false;
    };
    CBaseField.prototype.SetDisplay = function(nType) {
        this._display = nType;
    };
    CBaseField.prototype.GetDisplay = function() {
        return this._display;
    };
    CBaseField.prototype.GetDefaultValue = function() {
        if (this._defaultValue == null && this.GetParent())
            return this.GetParent().GetDefaultValue();
        
        return this._defaultValue;
    };
    CBaseField.prototype.SetDefaultValue = function(value) {
        this._defaultValue = value;
    };

    /**
	 * Sets default value for form.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 */
    CBaseField.prototype.Reset = function() {
        let defValue = this.GetDefaultValue();
        if (this.GetValue() != defValue) {
            
            this.SetValue(defValue);
            this.SetApiValue(defValue);
            this.AddToRedraw();
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
                    if (this.GetBorderStyle() !== BORDER_TYPES.beveled)
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
            if ((this.GetType() == AscPDF.FIELD_TYPES.radiobutton || this.GetType() == AscPDF.FIELD_TYPES.checkbox) && this.IsPressed() && this.GetBorderStyle() !== BORDER_TYPES.beveled) {
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
    
    CBaseField.prototype.RevertContentViewToOriginal = function() {
        this.content.ResetShiftView();
        this._curShiftView.x = this._originShiftView.x;
        this._curShiftView.y = this._originShiftView.y;

        this._bAutoShiftContentView = false;
        this.content.ShiftView(this._originShiftView.x, this._originShiftView.y);

        if (this._scrollInfo) {
            let nMaxShiftY                  = this._scrollInfo.scroll.maxScrollY;
            this._scrollInfo.scrollCoeff    = Math.abs(this._curShiftView.y / nMaxShiftY);
        }

        this.AddToRedraw();
    };
    CBaseField.prototype.IsWidget = function() {
        return this._isWidget;
    };
    CBaseField.prototype.IsNeedRevertShiftView = function() {
        if (this._curShiftView.y != this._originShiftView.y ||
            this._curShiftView.x != this._originShiftView.x)
            return true;
    };
    CBaseField.prototype.GetBordersWidth = function(bScaled) {
        let oViewer = editor.getDocumentRenderer();
        let nLineWidth = bScaled == true ? 1.25 * this._lineWidth * AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom : 1.25 * this._lineWidth;

        if (nLineWidth == 0) {
            return {
                left:     0,
                top:      0,
                right:    0,
                bottom:   0
            }
        }

        switch (this._borderStyle) {
            case BORDER_TYPES.solid:
            case BORDER_TYPES.dashed:
                return {
                    left:     nLineWidth,
                    top:      nLineWidth,
                    right:    nLineWidth,
                    bottom:   nLineWidth
                }
            case BORDER_TYPES.beveled:
            case BORDER_TYPES.inset:
                return {
                    left:     2 * nLineWidth,
                    top:      2 * nLineWidth,
                    right:    2 * nLineWidth,
                    bottom:   2 * nLineWidth
                }
            case BORDER_TYPES.underline:
                return {
                    left:     0,
                    top:      0,
                    right:    0,
                    bottom:   nLineWidth
                }
        }
    };
    CBaseField.prototype.GetMarginsFromBorders = function(bScaled, bMM) {
        let oBorders    = this.GetBordersWidth(bScaled);
        let nKoeff      = bMM == true ? g_dKoef_pix_to_mm : 1;

        switch (this._borderStyle) {
            case BORDER_TYPES.solid:
            case BORDER_TYPES.dashed:
            case BORDER_TYPES.underline:
                return {
                    left:     oBorders.bottom * nKoeff,
                    top:      oBorders.bottom * nKoeff,
                    right:    oBorders.bottom * nKoeff,
                    bottom:   oBorders.bottom * nKoeff
                }
            case BORDER_TYPES.inset:
            case BORDER_TYPES.beveled:
                return {
                    left:     oBorders.bottom * nKoeff,
                    top:      oBorders.bottom * nKoeff,
                    right:    oBorders.bottom * nKoeff,
                    bottom:   oBorders.bottom * nKoeff
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
            this._borderStyle = nStyle;
            this.SetWasChanged(true);

            if (this.IsComb && this.IsComb() == true) {
                this.SetNeedRecalc(true);
                this.content.GetElement(0).Content.forEach(function(run) {
                    run.RecalcInfo.Measure = true;
                });
            }

            this.AddToRedraw();
        }
    };
    CBaseField.prototype.GetBorderStyle = function() {
        return this._borderStyle;
    };
    CBaseField.prototype.SetBorderWidth = function(nWidth) {
        if (this._borderWidth != nWidth) {
            this._borderWidth = nWidth;
            this._lineWidth = nWidth;

            this.SetWasChanged(true);

            this.AddToRedraw();
        }
    };
    /**
     * Returns a canvas with origin view (from appearance stream) of current form.
	 * @memberof CBaseField
     * @param {number} nAPType - APPEARANCE_TYPE (type of AP)
	 * @typeofeditors ["PDF"]
     * @returns {canvas}
	 */
    CBaseField.prototype.GetOriginView = function(nAPType) {
        if (this._apIdx == -1)
            return null;

        let oViewer = editor.getDocumentRenderer();
        let oFile   = oViewer.file;
        
        let oApearanceInfo  = this.GetOriginViewInfo();
        let oSavedView, oApInfoTmp;
        if (!oApearanceInfo)
            return null;
            
        switch (nAPType) {
            case APPEARANCE_TYPE.normal:
                oApInfoTmp = oApearanceInfo["N"];
                oSavedView = this._originView.normal;
                break;
            case APPEARANCE_TYPE.rollover:
                if (oApearanceInfo["R"]) {
                    oApInfoTmp = oApearanceInfo["R"]
                    oSavedView = this._originView.rollover;
                }
                else {
                    oApInfoTmp = oApearanceInfo["N"]
                    oSavedView = this._originView.normal;
                }
                break;
            case APPEARANCE_TYPE.mouseDown:
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
        let mappedBuffer    = new Uint8ClampedArray(oFile.memory().buffer, oApInfoTmp["retValue"], 4 * nWidth * nHeight);
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
            case APPEARANCE_TYPE.normal:
                this._originView.normal = canvas;
                break;
            case APPEARANCE_TYPE.rollover:
                if (oApearanceInfo["R"]) {
                    this._originView.rollover = canvas;
                }
                else {
                    this._originView.normal = canvas;
                }
                break;
            case APPEARANCE_TYPE.mouseDown:
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
    CBaseField.prototype.GetOriginViewInfo = function() {
        let oViewer     = editor.getDocumentRenderer();
        let oFile       = oViewer.file;
        let nPage       = this.GetOriginPage();
        let oOriginPage = oFile.pages.find(function(page) {
            return page.originIndex == nPage;
        });

        let w = ((oOriginPage.W * 96 / oOriginPage.Dpi) >> 0) * AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom >> 0;
        let h = ((oOriginPage.H * 96 / oOriginPage.Dpi) >> 0) * AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom >> 0;

        if (oOriginPage.fieldsAPInfo == null || oOriginPage.fieldsAPInfo.size.w != w || oOriginPage.fieldsAPInfo.size.h != h) {
            oOriginPage.fieldsAPInfo = {
                info: oFile.nativeFile["getInteractiveFormsAP"](nPage, w, h),
                size: {
                    w: w,
                    h: h
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
        if (this.IsHidden())
            return;
		
        if (this.IsNeedDrawFromStream()) {
            this.Recalculate();
            this.DrawFromStream(pdfGraphics);
        }
        else
            this.DrawFromTextBox(pdfGraphics, textBoxGraphics, pageIndex);
    };

    CBaseField.prototype.DrawFromStream = function(oGraphicsPDF) {
        let originView      = this.GetOriginView(this.IsHovered && this.IsHovered() ? AscPDF.APPEARANCE_TYPE.rollover : undefined);
        let nGrScale        = oGraphicsPDF.GetScale();

        if (originView) {
            oGraphicsPDF.DrawImage(originView, 0, 0, originView.width / nGrScale, originView.height / nGrScale, originView.x / nGrScale, originView.y / nGrScale, originView.width / nGrScale, originView.height / nGrScale);
        }
    };
	CBaseField.prototype.DrawFromTextBox = function(pdfGraphics, textBoxGraphics, pageIndex) {
		let _t = this;
		if (!this._doc.checkDefaultFieldFonts(function() {
			let viewer = _t._doc.GetDocumentRenderer();
			if (!viewer)
				return;
			
			viewer.repaintFormsOnPage(pageIndex);
		})) {
			return;
		}
		
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

    CBaseField.prototype.SetApiTextColor = function(aColor) {
        if ([AscPDF.FIELD_TYPES.radiobutton, AscPDF.FIELD_TYPES.checkbox].includes(this.GetType()))
            return;

        let color = AscPDF.Api.Objects.color;

        let oRGB = color.convert(aColor, "RGB");
        if (this.content) {
            let oPara       = this.content.GetElement(0);
            let oApiPara    = editor.private_CreateApiParagraph(oPara);

            oApiPara.SetColor(oRGB[1] * 255, oRGB[2] * 255, oRGB[3] * 255, false);
            oPara.RecalcCompiledPr(true);
        }
        if (this.contentFormat) {
            let oPara       = this.contentFormat.GetElement(0);
            let oApiPara    = editor.private_CreateApiParagraph(oPara);

            oApiPara.SetColor(oRGB[1] * 255, oRGB[2] * 255, oRGB[3] * 255, false);
            oPara.RecalcCompiledPr(true);
        }

        let oApiColor   = color.convert(oRGB, aColor[0]);
        this._textColor = oApiColor.slice(1);

        this.SetWasChanged(true);
        this.AddToRedraw();
    };
    
    CBaseField.prototype.SetTextColor = function(aColor) {
        this._textColor = aColor;
        
        let oRGB = this.GetRGBColor(aColor);
        if (this.content) {
            let oPara       = this.content.GetElement(0);
            let oApiPara    = editor.private_CreateApiParagraph(oPara);

            oApiPara.SetColor(oRGB.r, oRGB.g, oRGB.b, false);
            oPara.RecalcCompiledPr(true);
        }
        if (this.contentFormat) {
            let oPara       = this.contentFormat.GetElement(0);
            let oApiPara    = editor.private_CreateApiParagraph(oPara);

            oApiPara.SetColor(oRGB.r, oRGB.g, oRGB.b, false);
            oPara.RecalcCompiledPr(true);
        }
        
        this.SetWasChanged(true);
        this.AddToRedraw();
    };
    CBaseField.prototype.GetTextColor = function() {
        return this._textColor;
    };
    /**
     * Is the field completely within the window of view.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
     * @returns {boolean}
	 */
    CBaseField.prototype.IsInSight = function() {
        let oViewer     = editor.getDocumentRenderer();
        let aOrigRect   = this.GetOrigRect();
        let nPage       = this.GetPage();

        let oPage;
        for (let i = 0; i < oViewer.pageDetector.pages.length; i++) {
            if (oViewer.pageDetector.pages[i].num == nPage) {
                oPage = oViewer.pageDetector.pages[i];
                break;
            }
        }

        if (!oPage)
            return false;

        // координаты видимой части страницы
        let x1, x2, y1, y2;

        x1 = (-oPage.x / oPage.w) * oViewer.file.pages[nPage].W;
        y1 = (-oPage.y / oPage.h) * oViewer.file.pages[nPage].H;

        x2 = x1 + oViewer.canvas.width / (oPage.w) * oViewer.file.pages[nPage].W;
        y2 = y1 + oViewer.canvas.height / (oPage.h) * oViewer.file.pages[nPage].H;
        
        if (aOrigRect[0] >= x1 && aOrigRect[1] >= y1 && aOrigRect[2] <= x2 && aOrigRect[3] <= y2)
            return true;
        else
            return false;

    };
    CBaseField.prototype.GetOrigRect = function() {
        return this._origRect;
    };
    CBaseField.prototype.GetRect = function() {
        return this._rect;
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
        return this._page;
    };
    CBaseField.prototype["getPagePos"] = function() {
        if (!this._pagePos)
            return null;
        return {
            "x" : this._pagePos.x,
            "y" : this._pagePos.y,
            "w" : this._pagePos.w,
            "h" : this._pagePos.h
        };
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

    /**
	 * Converts global coords to page coords.
     * Note: use scaled coordinates like pagePos_ from field, and not original like _origRect from field.
     * @param {Number} x
     * @param {Number} y
     * @param {Number} nPage
     * @param {boolean} isNotMM - coordinates in millimeters or not 
	 * @typeofeditors ["PDF"]
	 */
    function GetPageCoordsByGlobalCoords(x, y, nPage, isNotMM) {
        // конвертация из глобальных x, y к mm кординатам самой страницы
        let oViewer = editor.getDocumentRenderer();
        var pageObject = oViewer.getPageByCoords(x - oViewer.x, y - oViewer.y);

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        if (!pageObject) {
            return {X: 0, Y: 0}
        }

        let result = {
            X : isNotMM ? (pageObject.x) * nScaleY : (pageObject.x) * g_dKoef_pix_to_mm * nScaleY,
            Y : isNotMM ? (pageObject.y) * nScaleX : (pageObject.y) * g_dKoef_pix_to_mm * nScaleX
        };

        result["X"] = result.X;
        result["Y"] = result.Y;

        return result;
    }
    /**
	 * Converts page coords to global coords.
     * Note: use scaled coordinates like pagePos_ from field, and not original like _origRect from field.
     * @param {Number} x
     * @param {Number} y
     * @param {Number} nPage
     * @param {boolean} isNotMM - coordinates in millimeters or not 
	 * @typeofeditors ["PDF"]
	 */
    function GetGlobalCoordsByPageCoords(x, y, nPage, isNotMM) {
        let oViewer = editor.getDocumentRenderer();

        let nScaleY = oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W / oViewer.zoom;

        let X = isNotMM ? x / nScaleX : x * g_dKoef_mm_to_pix / nScaleX;
        let Y = isNotMM ? y / nScaleY : y * g_dKoef_mm_to_pix / nScaleY;

        let pageCoords;
        for (let i = 0; i < oViewer.pageDetector.pages.length; i++) {
            if (oViewer.pageDetector.pages[i].num == nPage)
                pageCoords = oViewer.pageDetector.pages[i];
        }

        let result = {
            X : (X * pageCoords.w / oViewer.file.pages[nPage].W + pageCoords.x) / AscCommon.AscBrowser.retinaPixelRatio,
            Y : (Y * pageCoords.h / oViewer.file.pages[nPage].H + pageCoords.y) / AscCommon.AscBrowser.retinaPixelRatio
        };

        result["X"] = result.X;
        result["Y"] = result.Y;

        return result;
    }

    function invertRGB(oColor) {
        // Calculate the inverted components
        const invertedR = 255 - oColor.r;
        const invertedG = 255 - oColor.g;
        const invertedB = 255 - oColor.b;
      
        return {r: invertedR, g: invertedG, b: invertedB}
      }

    if (!window["AscPDF"])
	    window["AscPDF"] = {};
    
	window["AscPDF"].ALIGN_TYPE         = ALIGN_TYPE;
	window["AscPDF"].BORDER_TYPES       = BORDER_TYPES;
    window["AscPDF"].APPEARANCE_TYPE    = APPEARANCE_TYPE;
    window["AscPDF"].VALID_ROTATIONS    = VALID_ROTATIONS;
    window["AscPDF"].MAX_TEXT_SIZE      = MAX_TEXT_SIZE;
    window["AscPDF"].FONT_STRETCH       = FONT_STRETCH;
    window["AscPDF"].FONT_STYLE         = FONT_STYLE;
    window["AscPDF"].FONT_WEIGHT        = FONT_WEIGHT;
	window["AscPDF"].DEFAULT_FIELD_FONT = DEFAULT_FIELD_FONT;

    window["AscPDF"].CBaseField = CBaseField;
    window["AscPDF"]["GetGlobalCoordsByPageCoords"] = window["AscPDF"].GetGlobalCoordsByPageCoords = GetGlobalCoordsByPageCoords;
    window["AscPDF"]["GetPageCoordsByGlobalCoords"] = window["AscPDF"].GetPageCoordsByGlobalCoords = GetPageCoordsByGlobalCoords;
    
})();

