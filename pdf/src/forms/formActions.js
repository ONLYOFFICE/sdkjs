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
    let FORMS_TRIGGERS_TYPES = {
        MouseUp:    0,
        MouseDown:  1,
        MouseEnter: 2,
        MouseExit:  3,
        OnFocus:    4,
        OnBlur:     5,
        Keystroke:  6,
        Validate:   7,
        Calculate:  8,
        Format:     9
    }
    let ACTIONS_TYPES = {
        Unknown:        0,
        GoTo:           1,
        GoToR:          2,
        GoToE:          3,
        Launch:         4,
        Thread:         5,
        URI:            6,
        Sound:          7,
        Movie:          8,
        HideShow:       9,
        Named:          10,
        SubmitForm:     11,
        ResetForm:      12,
        ImportData:     13,
        JavaScript:     14,
        SetOCGState:    15,
        Rendition:      16,
        Trans:          17,
        GoTo3DView:     18
    }

    let ACTION_NAMED_TYPES = {
        NextPage:   0,
        PrevPage:   1,
        FirstPage:  2,
        LastPage:   3
    }

    let GOTO_TYPES = { // see description in pdf specification (table 151 destination syntax)
        xyz:    0,
        fit:    1,
        fitH:   2,
        fitV:   3,
        fitR:   4,
        fitB:   5,
        fitBH:  6,
        fitBV:  7
    }

    function CFormTriggers() {
        this.MouseUp = null; 
        this.MouseDown = null; 
        this.MouseEnter = null; 
        this.MouseExit = null; 
        this.OnFocus = null; 
        this.OnBlur = null; 
        this.Keystroke = null; 
        this.Validate = null; 
        this.Calculate = null; 
        this.Format = null;
    }
    CFormTriggers.prototype.Copy = function() {
        let newObj = new CFormTriggers();
        if (this.MouseUp != null)
            newObj.MouseUp = this.MouseUp.Copy(); 
        if (this.MouseDown != null)
            newObj.MouseDown = this.MouseDown.Copy(); 
        if (this.MouseEnter != null)
            newObj.MouseEnter = this.MouseEnter.Copy(); 
        if (this.MouseExit != null)
            newObj.MouseExit = this.MouseExit.Copy(); 
        if (this.OnFocus != null)
            newObj.OnFocus = this.OnFocus.Copy(); 
        if (this.OnBlur != null)
            newObj.OnBlur = this.OnBlur.Copy(); 
        if (this.Keystroke != null)
            newObj.Keystroke = this.Keystroke.Copy(); 
        if (this.Validate != null)
            newObj.Validate = this.Validate.Copy(); 
        if (this.Calculate != null)
            newObj.Calculate = this.Calculate.Copy(); 
        if (this.Format != null)
            newObj.Format = this.Format.Copy();

        return newObj;
    }

    function CFormTrigger(type, aActions) {
        this.type = type;

        // actions
        this.Actions = aActions;
    }
    CFormTrigger.prototype.Copy = function() {
        return new CFormTrigger(this.type, this.script);
    };
    CFormTrigger.prototype.GetActions = function() {
        return this.Actions;
    };

    function CActionBase(nType) {
        this.type = nType;
        this.field = null;
        this.triggerType = undefined;
    };
    CActionBase.prototype.GetType = function() {
        return this.type;
    };
    CActionBase.prototype.SetField = function(oField) {
        this.field = oField;
    };
    CActionBase.prototype.SetTrigger = function(nType) {
        this.triggerType = nType;
    };

    function CActionGoTo(nPage, nGoToType, nZoom, oRect) {
        CActionBase.call(this, ACTIONS_TYPES.GoTo);
        this.page       = nPage;
        this.goToType   = nGoToType;
        this.zoom       = nZoom;
        this.rect       = oRect; // top right bottom left
    };
    CActionGoTo.prototype = Object.create(CActionBase.prototype);
	CActionGoTo.prototype.constructor = CActionGoTo;

    CActionGoTo.prototype.GetZoom = function() {
        if (this.zoom != null)
            return this.zoom;

        let oViewer     = editor.getDocumentRenderer();
        let nNoZoomH    = oViewer.drawingPages[this.page].H / oViewer.zoom;
        let nNoZoomW    = oViewer.drawingPages[this.page].W / oViewer.zoom;

        let nScaleY = oViewer.drawingPages[this.page].H / oViewer.file.pages[this.page].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[this.page].W / oViewer.file.pages[this.page].W / oViewer.zoom;

        switch (this.goToType) {
            case GOTO_TYPES.xyz: // inherit zoom
                break;
            case GOTO_TYPES.fit:
            case GOTO_TYPES.fitB: { // fit to max of heigth/width
                let nVerZoom = ((oViewer.canvas.height / (nNoZoomH * AscCommon.AscBrowser.retinaPixelRatio)) * 100 >> 0) / 100;
                let nHorZoom = ((oViewer.canvas.width / (nNoZoomW * AscCommon.AscBrowser.retinaPixelRatio)) * 100 >> 0) / 100;

                this.zoom = Math.min(nHorZoom, nVerZoom);
                break;
            }
            case GOTO_TYPES.fitH:
            case GOTO_TYPES.fitBH: { // fit to width
                this.zoom = ((oViewer.canvas.width / (nNoZoomW * AscCommon.AscBrowser.retinaPixelRatio)) * 100 >> 0) / 100;
                break;
            }
            case GOTO_TYPES.fitV:
            case GOTO_TYPES.fitBV: { // fit to heigth
                this.zoom = ((oViewer.canvas.height / (nNoZoomH * AscCommon.AscBrowser.retinaPixelRatio)) * 100 >> 0) / 100;
                break;
            }
            case GOTO_TYPES.fitR: { // fit to rect
                let nRectW = (this.rect.right - this.rect.left) * nScaleX * AscCommon.AscBrowser.retinaPixelRatio;
                let nRectH = (this.rect.bottom - this.rect.top) * nScaleY * AscCommon.AscBrowser.retinaPixelRatio;

                let nVerZoom = ((oViewer.canvas.height / (nRectH)) * 100 >> 0) / 100;
                let nHorZoom = ((oViewer.canvas.width / (nRectW)) * 100 >> 0) / 100;

                let nMinZoom = Math.min(nHorZoom, nVerZoom);
                
                // далее вычисляем ширину с новым потенциальным зумом,
                // если при данных размерах будет добавлен скролл, то вычитаем его ширину и пересчитываем zoom
                let nNewPageW = oViewer.drawingPages[this.page].W = (oViewer.file.pages[this.page].W * 96 * nMinZoom / oViewer.file.pages[this.page].Dpi) >> 0;
                if (nNewPageW > oViewer.width) {
                    nVerZoom = (((oViewer.canvas.height - oViewer.scrollWidth) / (nRectH)) * 100 >> 0) / 100;
                }
                
                this.zoom = Math.min(nHorZoom, nVerZoom);
            }
        }

        return this.zoom;
    };

    CActionGoTo.prototype.Do = function() {
        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.field.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oActionsQueue.SetCurAction(this);
        
        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.triggerType == FORMS_TRIGGERS_TYPES.OnFocus && this.field != oDoc.activeForm)
            oActionsQueue.Continue();

        if (this.page >= oViewer.pagesInfo.countTextPages) {
            oActionsQueue.Continue();
            return;
        }

        let nZoom = this.GetZoom();
        if (nZoom && oViewer.zoom != nZoom)
            oViewer.setZoom(nZoom, true);

        // выставляем смещения
        let yOffset;
        let xOffset;
        if (this.rect.top != null) {
            yOffset = this.rect.top + oViewer.betweenPages / (oViewer.drawingPages[this.page].H / oViewer.file.pages[this.page].H);
        }
        else
            yOffset = oViewer.betweenPages / (oViewer.drawingPages[this.page].H / oViewer.file.pages[this.page].H);

        if (this.rect.left != null) {
            xOffset = this.rect.left;
        }

        if ((nZoom && oViewer.zoom != nZoom) || yOffset != undefined && xOffset != undefined || oViewer.currentPage != this.page) {
            oViewer.disabledPaintOnScroll = true; // вырубаем отрисовку на скроле
            oViewer.navigateToPage(this.page, yOffset, xOffset);
            oViewer.disabledPaintOnScroll = false;
            oViewer.needRedraw = true; // в конце Actions выполним отрисовку
        }

        oActionsQueue.Continue();
    };

    function CActionNamed(nType) {
        CActionBase.call(this, ACTIONS_TYPES.Named);
        this.nameType = nType;
    };
    CActionNamed.prototype = Object.create(CActionBase.prototype);
	CActionNamed.prototype.constructor = CActionNamed;

    CActionNamed.prototype.GetName = function() {
        switch (this.nameType) {
            case ACTION_NAMED_TYPES.NextPage:
                return "NextPage";
            case ACTION_NAMED_TYPES.PrevPage:
                return "PrevPage";
            case ACTION_NAMED_TYPES.FirstPage:
                return "FirstPage";
            case ACTION_NAMED_TYPES.LastPage:
                return "LastPage";
        }

        return "";
    };

    CActionNamed.GetInternalType = function(sType) {
        switch (sType) {
            case "NextPage":
                return ACTION_NAMED_TYPES.NextPage;
            case "PrevPage":
                return ACTION_NAMED_TYPES.PrevPage;
            case "FirstPage":
                return ACTION_NAMED_TYPES.FirstPage;
            case "LastPage":
                return ACTION_NAMED_TYPES.LastPage;
        }

        return -1;
    };

    CActionNamed.prototype.Do = function() {
        let Api             = editor;
        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.field.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oActionsQueue.SetCurAction(this);

        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.triggerType == FORMS_TRIGGERS_TYPES.OnFocus && this.field != oDoc.activeForm)
            oActionsQueue.Continue();

        switch (this.nameType) {
            case ACTION_NAMED_TYPES.FirstPage:
                Api.goToPage(0);
                break;
            case ACTION_NAMED_TYPES.NextPage:
                if (oViewer.currentPage + 1 <= oViewer.pagesInfo.countTextPages)
                    Api.goToPage(oViewer.currentPage + 1);
                break;
            case ACTION_NAMED_TYPES.PrevPage:
                if (oViewer.currentPage - 1 >= 0)
                    Api.goToPage(oViewer.currentPage - 1);
                break;
            case ACTION_NAMED_TYPES.LastPage:
                if (oViewer.currentPage != oViewer.pagesInfo.countTextPages)
                    Api.goToPage(oViewer.pagesInfo.countTextPages - 1);
                break;
        }

        oActionsQueue.Continue();
    };

    function CActionURI(sURI) {
        CActionBase.call(this, ACTIONS_TYPES.URI);
        this.uri = sURI;
    };
    CActionURI.prototype = Object.create(CActionBase.prototype);
	CActionURI.prototype.constructor = CActionURI;

    CActionURI.prototype.Do = function() {
        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.field.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oActionsQueue.SetCurAction(this);

        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.triggerType == FORMS_TRIGGERS_TYPES.OnFocus && this.field != oDoc.activeForm)
            oActionsQueue.Continue();

        editor.sendEvent("asc_onOpenLinkPdfForm", this.uri, this.OpenLink.bind(this), oActionsQueue.Continue.bind(oActionsQueue));
    };

    CActionURI.prototype.OpenLink = function() {
        window.open(this.uri, "_blank");

        this.field.GetDocument().GetActionsQueue().Continue();
    };

    function CActionHideShow(bHidden, aFieldsNames) {
        CActionBase.call(this, ACTIONS_TYPES.HideShow);
        this.hidden = bHidden;
        this.names = aFieldsNames;
    };

    CActionHideShow.prototype = Object.create(CActionBase.prototype);
	CActionHideShow.prototype.constructor = CActionHideShow;

    CActionHideShow.prototype.Do = function() {
        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.field.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oActionsQueue.SetCurAction(this);

        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.triggerType == FORMS_TRIGGERS_TYPES.OnFocus && this.field != oDoc.activeForm)
            oActionsQueue.Continue();

        oDoc.HideShowForms(this.hidden, this.names);
    };

    function CActionReset(aFieldsNames) {
        CActionBase.call(this, ACTIONS_TYPES.Reset);
        this.names = aFieldsNames;
    };
    CActionReset.prototype = Object.create(CActionBase.prototype);
	CActionReset.prototype.constructor = CActionReset;

    CActionReset.prototype.Do = function() {
        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.field.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oActionsQueue.SetCurAction(this);

        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.triggerType == FORMS_TRIGGERS_TYPES.OnFocus && this.field != oDoc.activeForm)
            oActionsQueue.Continue();
            
        oDoc.ResetForms(this.names);
    };

    function CActionRunScript(script) {
        CActionBase.call(this, ACTIONS_TYPES.JavaScript);
        this.script = script;
        this.bContinueAfterEval = true; // выключаем на асинхронных операциях
    };
    CActionRunScript.prototype = Object.create(CActionBase.prototype);
	CActionRunScript.prototype.constructor = CActionRunScript;

    CActionRunScript.prototype.Do = function() {
        let oViewer         = editor.getDocumentRenderer();
        let oDoc            = this.field.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oActionsQueue.SetCurAction(this);

        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.triggerType == FORMS_TRIGGERS_TYPES.OnFocus && this.field != oDoc.activeForm) {
            oActionsQueue.Continue();
            return;
        }

        oDoc.SetEvent({
            "target": this.field.GetFormApi(),
            "rc": true
        });

        try {
            EvalScript(this.script, oDoc);
        }
        catch (err) {
            console.log(err);
        }

        if (this.bContinueAfterEval == true)
            oActionsQueue.Continue();
    };

    CActionRunScript.prototype.RunScript = function() {
        let oDoc = this.field.GetDocument();

        oDoc.SetEvent({
            "target": this.field.GetFormApi(),
            "rc": true
        });

        try {
            EvalScript(this.script, oDoc);
        }
        catch (err) {
            console.log(err);
        }
    };
	
    function EvalScript(str, oParentDoc) {
        let aArgsNamesToDelete = [
            "window",
            "setTimeout",
            "setInterval",
            "XMLHttpRequest",
            "Promise",
            "console"
        ];
    
        let oApiConsole = {
            "println"(value) {
                console.log("\n");
                console.log(value);
            },
            "clear"() {
                console.clear();
            },
            "hide"() {
                return false;
            },
            "show"() {
                return false;
            }
        };
        Object.freeze(oApiConsole);
    
        let oApiObjects = AscPDF.Api.Objects;
        let aArgsNamesPdfApi = [
            "event",
            "color",

            "AFNumber_Format",
            "AFNumber_Keystroke",
            "AFPercent_Format",
            "AFPercent_Keystroke",
            "AFDate_Format",
            "AFDate_Keystroke",
            "AFDate_FormatEx",
            "AFDate_KeystrokeEx",
            "AFTime_Format",
            "AFTime_Keystroke",
            "AFTime_FormatEx",
            "AFTime_KeystrokeEx",
            "AFSpecial_Format",
            "AFSpecial_Keystroke",
            "AFSpecial_KeystrokeEx",
            "AFSimple_Calculate",
            "AFRange_Validate",
        ];
    
        let oApiFunc = AscPDF.Api.Functions;
        let aArgsPdfApi = [
            oParentDoc.event,
            oApiObjects.color,

            oApiFunc.AFNumber_Format,
            oApiFunc.AFNumber_Keystroke,
            oApiFunc.AFPercent_Format,
            oApiFunc.AFPercent_Keystroke,
            oApiFunc.AFDate_Format,
            oApiFunc.AFDate_Keystroke,
            oApiFunc.AFDate_FormatEx,
            oApiFunc.AFDate_KeystrokeEx,
            oApiFunc.AFTime_Format,
            oApiFunc.AFTime_Keystroke,
            oApiFunc.AFTime_FormatEx,
            oApiFunc.AFTime_KeystrokeEx,
            oApiFunc.AFSpecial_Format,
            oApiFunc.AFSpecial_Keystroke,
            oApiFunc.AFSpecial_KeystrokeEx,
            oApiFunc.AFSimple_Calculate,
            oApiFunc.AFRange_Validate
        ];
    
        let funcArgs = aArgsNamesToDelete.concat(aArgsNamesPdfApi);
        funcArgs.push(str);
    
        let func = Function.apply(null, funcArgs);
        func.bind(oParentDoc.GetDocumentApi()).apply(null, new Array(aArgsNamesToDelete.length - 1).concat(oApiConsole, aArgsPdfApi));
    }
    

    if (!window["AscPDF"])
	    window["AscPDF"] = {};
    
    window["AscPDF"].CFormTriggers      = CFormTriggers;
    window["AscPDF"].CFormTrigger       = CFormTrigger;
    window["AscPDF"].CActionGoTo        = CActionGoTo;
    window["AscPDF"].CActionNamed       = CActionNamed;
    window["AscPDF"].CActionURI         = CActionURI;
    window["AscPDF"].CActionHideShow    = CActionHideShow;
    window["AscPDF"].CActionReset       = CActionReset;
    window["AscPDF"].CActionRunScript   = CActionRunScript;
    
    window["AscPDF"].ACTIONS_TYPES          = ACTIONS_TYPES;
    window["AscPDF"].FORMS_TRIGGERS_TYPES   = FORMS_TRIGGERS_TYPES;

})();

