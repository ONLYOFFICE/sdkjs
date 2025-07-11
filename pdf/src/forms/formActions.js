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
    
    CFormTriggers.prototype.Copy = function(oParentField) {
        let oCopy = new CFormTriggers();
        if (this.MouseUp != null)
            oCopy.MouseUp = this.MouseUp.Copy(oParentField); 
        if (this.MouseDown != null)
            oCopy.MouseDown = this.MouseDown.Copy(oParentField); 
        if (this.MouseEnter != null)
            oCopy.MouseEnter = this.MouseEnter.Copy(oParentField); 
        if (this.MouseExit != null)
            oCopy.MouseExit = this.MouseExit.Copy(oParentField); 
        if (this.OnFocus != null)
            oCopy.OnFocus = this.OnFocus.Copy(oParentField); 
        if (this.OnBlur != null)
            oCopy.OnBlur = this.OnBlur.Copy(oParentField); 
        if (this.Keystroke != null)
            oCopy.Keystroke = this.Keystroke.Copy(oParentField); 
        if (this.Validate != null)
            oCopy.Validate = this.Validate.Copy(oParentField); 
        if (this.Calculate != null)
            oCopy.Calculate = this.Calculate.Copy(oParentField); 
        if (this.Format != null)
            oCopy.Format = this.Format.Copy(oParentField);

        return oCopy;
    }

    function CFormTrigger(type, aActions) {
        this.type = type;
        this.parentField = null;

        // actions
        this.Actions = aActions;

        let _t = this;
        this.Actions.forEach(function(action) {
            action.SetParent(_t);
        });
    }
    CFormTrigger.GetName = function(nType) {
        switch (nType) {
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseUp: {
                return "Mouse Up";
            }
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseDown: {
                return "Mouse Down";
            }
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseEnter: {
                return "Mouse Enter";
            }
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseExit: {
                return "Mouse Exit";
            }
            case AscPDF.FORMS_TRIGGERS_TYPES.OnFocus: {
                return "Focus";
            }
            case AscPDF.FORMS_TRIGGERS_TYPES.OnBlur: {
                return "Blur";
            }
            case AscPDF.FORMS_TRIGGERS_TYPES.Keystroke: {
                return "Keystroke";
            }
            case AscPDF.FORMS_TRIGGERS_TYPES.Validate: {
                return "Validate";
            }
            case AscPDF.FORMS_TRIGGERS_TYPES.Calculate: {
                return "Calculate";
            }
            case AscPDF.FORMS_TRIGGERS_TYPES.Format: {
                return "Format";
            }
        }
    };
    CFormTrigger.prototype.Copy = function(oParentField) {
        let aActionsCopies = [];
        for (let i = 0; i < this.Actions.length; i++) {
            let action = this.Actions[i];
            aActionsCopies.push(action.Copy());
        }

        let oCopy = new CFormTrigger(this.type, aActionsCopies);
        oCopy.SetParentField(oParentField); 

        return oCopy;
    };
    CFormTrigger.prototype.GetActions = function() {
        return this.Actions;
    };
    CFormTrigger.prototype.GetType = function() {
        return this.type;
    };
    CFormTrigger.prototype.SetParentField = function(oField) {
        this.parentField = oField;
    };
    CFormTrigger.prototype.GetParentField = function() {
        return this.parentField;
    };
    CFormTrigger.prototype.SetCallerField = function(oField) {
        this.callerField = oField;
    };
    CFormTrigger.prototype.GetCallerFiled = function() {
        return this.callerField || this.GetParentField();
    };
    CFormTrigger.prototype.WriteToBinary = function(memory) {
        let nType = this.GetType();
        switch (nType) {
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseUp:
                memory.WriteString("A");
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseDown:
                memory.WriteString("D");
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseEnter:
                memory.WriteString("E");
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.MouseExit:
                memory.WriteString("X");
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.OnFocus:
                memory.WriteString("Fo");
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.OnBlur:
                memory.WriteString("Bl");
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Keystroke:
                memory.WriteString("K");
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Validate:
                memory.WriteString("V");
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Calculate:
                memory.WriteString("C");
                break;
            case AscPDF.FORMS_TRIGGERS_TYPES.Format:
                memory.WriteString("F");
                break;
        }

        for (let i = 0; i < this.Actions.length; i++) {
            this.Actions[i].WriteToBinary(memory);
            if (this.Actions[i + 1])
                memory.WriteByte(1);
            else
                memory.WriteByte(0);
        }
    };

    function CActionBase(nType) {
        this.type = nType;
        this.parent = null;
    };
    CActionBase.prototype.GetType = function() {
        return this.type;
    };
    CActionBase.prototype.GetCallerFiled = function() {
        let oTrigger = this.GetParent();
        if (oTrigger) {
            return oTrigger.GetCallerFiled();
        }
    };
    CActionBase.prototype.SetParent = function(oTrigger) {
        this.parent = oTrigger;
    };
    CActionBase.prototype.GetParent = function() {
        return this.parent;
    };
    CActionBase.prototype.GetTriggerType = function() {
        let oTrigger = this.GetParent();
        if (oTrigger) {
            return oTrigger.GetType();
        }
    };
    CActionBase.prototype.GetTriggerName = function() {
        let oTrigger = this.GetParent();
        if (oTrigger) {
            return CFormTrigger.GetName(oTrigger.GetType());
        }
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

    CActionGoTo.prototype.Copy = function() {
        return new CActionGoTo(this.GetPage(), this.GetKind(), this.GetZoom(true), this.GetRect().slice());
    };
    CActionGoTo.prototype.GetZoom = function(bSource) {
        if (this.zoom != null || bSource)
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

    CActionGoTo.prototype.GetPage = function() {
        return this.page;
    };

    CActionGoTo.prototype.GetKind = function() {
        return this.goToType;
    };

    CActionGoTo.prototype.GetRect = function() {
        return this.rect;
    };

    CActionGoTo.prototype.Do = function() {
        let oViewer         = Asc.editor.getDocumentRenderer();
        let oField          = this.GetCallerFiled();
        let oDoc            = oField.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oActionsQueue.SetCurAction(this);
        
        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.GetTriggerType() == FORMS_TRIGGERS_TYPES.OnFocus && oField != oDoc.activeForm) {
            oActionsQueue.Continue();
            return;
        }
            
        if (this.page >= oViewer.pagesInfo.countTextPages) {
            oActionsQueue.Continue();
            return;
        }

        let nZoom = this.GetZoom();
        if (nZoom && oViewer.zoom != nZoom)
            oViewer.setZoom(nZoom, true);

        // выставляем смещения
        let yOffset = this.rect.top != null ? this.rect.top : 0;
        let xOffset = this.rect.left != null ? this.rect.left : 0;

        if ((nZoom && oViewer.zoom != nZoom) || yOffset != undefined && xOffset != undefined || oViewer.currentPage != this.page) {
            let oTr = oDoc.pagesTransform[this.page].invert;
            let oPos = oTr.TransformPoint(xOffset, yOffset);

            oViewer.disabledPaintOnScroll = true; // вырубаем отрисовку на скроле
            oViewer.scrollToXY(oViewer.scrollY + oPos.y, oViewer.scrollX + oPos.x);
            oViewer.disabledPaintOnScroll = false;
            oViewer.needRedraw = true; // в конце Actions выполним отрисовку
        }

        oActionsQueue.Continue();
    };
    
    CActionGoTo.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(this.GetType());
        memory.WriteLong(this.GetPage());

        let nKind = this.GetKind();
        memory.WriteByte(nKind);

        switch (nKind) {
            case 0:
            case 2:
            case 3:
            case 6:
            case 7:
            {
                let nFlag = 0;
                let nStartPos = memory.GetCurPosition();
                memory.Skip(4);

                if (this.rect.left != null) {
                    nFlag |= (1 << 4);
                    memory.WriteDouble(this.rect.left);
                }
                if (this.rect.top != null) {
                    nFlag |= (1 << 4);
                    memory.WriteDouble(this.rect.top);
                }
                if (this.zoom != null) {
                    nFlag |= (1 << 4);
                    memory.WriteDouble(this.zoom);
                }

                // write flags
                let nEndPos = memory.GetCurPosition();
                memory.Seek(nStartPos);
                memory.WriteLong(nFlag);
                memory.Seek(nEndPos);
                break;
            }
            case 4:
            {
                memory.WriteDouble(this.rect.left);
                memory.WriteDouble(this.rect.bottom);
                memory.WriteDouble(this.rect.right);
                memory.WriteDouble(this.rect.top);
                break;
            }
        }
    };

    function CActionNamed(nType) {
        CActionBase.call(this, ACTIONS_TYPES.Named);
        this.nameType = nType;
    };
    CActionNamed.prototype = Object.create(CActionBase.prototype);
	CActionNamed.prototype.constructor = CActionNamed;

    CActionNamed.prototype.Copy = function() {
        return new CActionNamed(this.GetNameType());
    };
    CActionNamed.prototype.GetNameStrType = function() {
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
    CActionNamed.prototype.GetNameType = function() {
        return this.nameType;
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
        let Api             = Asc.editor;
        let oViewer         = Api.getDocumentRenderer();
        let oField          = this.GetCallerFiled();
        let oDoc            = oField.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oActionsQueue.SetCurAction(this);

        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.GetTriggerType() == FORMS_TRIGGERS_TYPES.OnFocus && oField != oDoc.activeForm) {
            oActionsQueue.Continue();
            return;
        }

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

    CActionNamed.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(this.GetType());
        memory.WriteString(this.GetNameStrType());
    };

    function CActionURI(sURI) {
        CActionBase.call(this, ACTIONS_TYPES.URI);
        this.uri = sURI;
    };
    CActionURI.prototype = Object.create(CActionBase.prototype);
	CActionURI.prototype.constructor = CActionURI;

    CActionURI.prototype.Copy = function() {
        return new CActionURI(this.GetURI());
    };
    CActionURI.prototype.Do = function() {
        let oField          = this.GetCallerFiled();
        let oDoc            = oField.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oActionsQueue.SetCurAction(this);

        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.GetTriggerType() == FORMS_TRIGGERS_TYPES.OnFocus && oField != oDoc.activeForm) {
            oActionsQueue.Continue();
            return;
        }

        Asc.editor.sendEvent("asc_onOpenLinkPdfForm", this.uri, this.OpenLink.bind(this), oActionsQueue.Continue.bind(oActionsQueue));
    };
    CActionURI.prototype.GetURI = function() {
        return this.uri;
    };
    CActionURI.prototype.OpenLink = function() {
        window.open(this.uri, "_blank");

        let oField          = this.GetCallerFiled();
        let oDoc            = oField.GetDocument();
        oDoc.GetActionsQueue().Continue();
    };

    CActionURI.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(this.GetType());
        memory.WriteString(this.GetURI());
    };

    function CActionHideShow(bHidden, aFieldsNames) {
        CActionBase.call(this, ACTIONS_TYPES.HideShow);
        this.hidden = bHidden;
        this.names = aFieldsNames;
    };

    CActionHideShow.prototype = Object.create(CActionBase.prototype);
	CActionHideShow.prototype.constructor = CActionHideShow;

    CActionHideShow.prototype.Copy = function() {
        return new CActionHideShow(this.GetHidden(), this.GetNames().slice());
    };
    CActionHideShow.prototype.Do = function() {
        let oField          = this.GetCallerFiled();
        let oDoc            = oField.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oActionsQueue.SetCurAction(this);

        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.GetTriggerType() == FORMS_TRIGGERS_TYPES.OnFocus && oField != oDoc.activeForm) {
            oActionsQueue.Continue();
            return;
        }

        oDoc.HideShowForms(this.hidden, this.names);
    };

    CActionHideShow.prototype.GetNames = function() {
        return this.names;
    };
    CActionHideShow.prototype.GetHidden = function() {
        return this.hidden;
    };

    CActionHideShow.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(this.GetType());
        if (this.hidden)
            memory.WriteByte(1);
        else
            memory.WriteByte(0);

        if (this.names) {
            memory.WriteLong(this.names.length);
            for (let i = 0; i < this.names.length; i++) {
                memory.WriteString(this.names[i]);
            }
        }
    };

    function CActionReset(aFieldsNames, bAllExcept) {
        CActionBase.call(this, ACTIONS_TYPES.ResetForm);
        this.names      = aFieldsNames;
        this.bAllExcept = bAllExcept;
    };
    CActionReset.prototype = Object.create(CActionBase.prototype);
	CActionReset.prototype.constructor = CActionReset;

    CActionReset.prototype.Copy = function() {
        return new CActionReset(this.GetNames().slice(), this.GetNeedAllExcept());
    };
    CActionReset.prototype.Do = function() {
        let oField          = this.GetCallerFiled();
        let oDoc            = oField.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oActionsQueue.SetCurAction(this);

        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.GetTriggerType() == FORMS_TRIGGERS_TYPES.OnFocus && oField != oDoc.activeForm) {
            oActionsQueue.Continue();
            return;
        }
            
        oDoc.ResetForms(this.names, this.bAllExcept);
    };

    CActionReset.prototype.GetNames = function() {
        return this.names;
    };
    CActionReset.prototype.GetNeedAllExcept = function() {
        return this.bAllExcept;
    };

    CActionReset.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(this.GetType());

        if (this.bAllExcept)
            memory.WriteLong(1);
        else
            memory.WriteLong(0);

        if (this.names) {
            memory.WriteLong(this.names.length);
            for (let i = 0; i < this.names.length; i++) {
                memory.WriteString(this.names[i]);
            }
        }
    };

    function CActionRunScript(script) {
        CActionBase.call(this, ACTIONS_TYPES.JavaScript);
        this.script = script;
        this.bContinueAfterEval = true; // выключаем на асинхронных операциях
    };
    CActionRunScript.prototype = Object.create(CActionBase.prototype);
	CActionRunScript.prototype.constructor = CActionRunScript;

    CActionRunScript.prototype.Copy = function() {
        return new CActionRunScript(this.GetScript());
    };
    CActionRunScript.prototype.Do = function() {
        let oField          = this.GetCallerFiled();
        let oDoc            = oField.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oActionsQueue.SetCurAction(this);

        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.GetTriggerType() == FORMS_TRIGGERS_TYPES.OnFocus && oField != oDoc.activeForm) {
            oActionsQueue.Continue();
            return;
        }

        const oEvent = oDoc.CreateEvent({
            "name": this.GetTriggerName(),
            "target": oField.GetFormApi(),
            "rc": true
        });

        try {
            EvalScript(this.script, oDoc, oEvent);
        }
        catch (err) {
            console.log(err);
        }
        finally {
            oDoc.event = oDoc.eventsStack.pop();
        }

        if (this.bContinueAfterEval == true)
            oActionsQueue.Continue();
    };

    CActionRunScript.prototype.RunScript = function(oEventPr) {
        let oField = this.GetCallerFiled();
        let oDoc = oField.GetDocument();

        if (!oEventPr) {
            oEventPr = {
                "name": this.GetTriggerName(),
                "target": oField.GetFormApi(),
                "rc": true
            }
        }

        const oEvent = oDoc.CreateEvent(oEventPr);
        
        try {
            EvalScript(this.script, oDoc, oEvent);
        }
        catch (err) {
            console.log(err);
        }
        finally {
            oDoc.event = oDoc.eventsStack.pop();
        }

        return oEvent;
    };

    CActionRunScript.prototype.GetScript = function() {
        return this.script;
    }

    CActionRunScript.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(this.GetType());
        memory.WriteString(this.script);
    };
	
    function EvalScript(str, oParentDoc, oEvent) {
        let aArgsNamesToDelete = [
            "window",
            "setTimeout",
            "setInterval",
            "XMLHttpRequest",
            "Promise",
            "console"
        ];
    
        let oApiConsole = {
            "println": function(value) {
                console.log(value);
            },
            "clear": function() {
                console.clear();
            },
            "hide": function() {
                return false;
            },
            "show": function() {
                return false;
            }
        };
        Object.freeze(oApiConsole);
    
        let oApiObjects = AscPDF.Api.Types;
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
    
        if (!oParentDoc.globalEventStack) {
            oParentDoc.globalEventStack = [];    
        }

        oParentDoc.globalEventStack.push(oParentDoc.event);
        oParentDoc.event = oEvent;

        let oApiFunc = AscPDF.Api.Functions;
        let aArgsPdfApi = [
            oEvent,
            oApiObjects["color"],

            oApiFunc["AFNumber_Format"],
            oApiFunc["AFNumber_Keystroke"],
            oApiFunc["AFPercent_Format"],
            oApiFunc["AFPercent_Keystroke"],
            oApiFunc["AFDate_Format"],
            oApiFunc["AFDate_Keystroke"],
            oApiFunc["AFDate_FormatEx"],
            oApiFunc["AFDate_KeystrokeEx"],
            oApiFunc["AFTime_Format"],
            oApiFunc["AFTime_Keystroke"],
            oApiFunc["AFTime_FormatEx"],
            oApiFunc["AFTime_KeystrokeEx"],
            oApiFunc["AFSpecial_Format"],
            oApiFunc["AFSpecial_Keystroke"],
            oApiFunc["AFSpecial_KeystrokeEx"],
            oApiFunc["AFSimple_Calculate"],
            oApiFunc["AFRange_Validate"]
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

