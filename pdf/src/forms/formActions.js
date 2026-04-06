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
    let PDF_TRIGGERS_TYPES = {
        mouseUp:    0,
        mouseDown:  1,
        mouseEnter: 2,
        mouseExit:  3,
        onFocus:    4,
        onBlur:     5,
        keystroke:  6,
        validate:   7,
        calculate:  8,
        format:     9
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

    function CPdfTriggers() {
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
    
    CPdfTriggers.prototype.Copy = function(oParentField) {
        let oCopy = new CPdfTriggers();
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

    function CPdfTrigger(type, aActions) {
        this.type = type;
        this.parentField = null;

        // actions
        this.Actions = aActions;

        let _t = this;
        this.Actions.forEach(function(action) {
            action.SetParent(_t);
        });
    }
    CPdfTrigger.GetName = function(nType) {
        switch (nType) {
            case AscPDF.PDF_TRIGGERS_TYPES.mouseUp: {
                return "Mouse Up";
            }
            case AscPDF.PDF_TRIGGERS_TYPES.mouseDown: {
                return "Mouse Down";
            }
            case AscPDF.PDF_TRIGGERS_TYPES.mouseEnter: {
                return "Mouse Enter";
            }
            case AscPDF.PDF_TRIGGERS_TYPES.mouseExit: {
                return "Mouse Exit";
            }
            case AscPDF.PDF_TRIGGERS_TYPES.onFocus: {
                return "Focus";
            }
            case AscPDF.PDF_TRIGGERS_TYPES.onBlur: {
                return "Blur";
            }
            case AscPDF.PDF_TRIGGERS_TYPES.keystroke: {
                return "Keystroke";
            }
            case AscPDF.PDF_TRIGGERS_TYPES.validate: {
                return "Validate";
            }
            case AscPDF.PDF_TRIGGERS_TYPES.calculate: {
                return "Calculate";
            }
            case AscPDF.PDF_TRIGGERS_TYPES.format: {
                return "Format";
            }
        }
    };
    CPdfTrigger.prototype.Copy = function(oParentField) {
        let aActionsCopies = [];
        for (let i = 0; i < this.Actions.length; i++) {
            let action = this.Actions[i];
            aActionsCopies.push(action.Copy());
        }

        let oCopy = new CPdfTrigger(this.type, aActionsCopies);
        oCopy.SetParentField(oParentField); 

        return oCopy;
    };
    CPdfTrigger.prototype.GetActions = function() {
        return this.Actions;
    };
    CPdfTrigger.prototype.GetType = function() {
        return this.type;
    };
    CPdfTrigger.prototype.SetParentField = function(oField) {
        this.parentField = oField;
    };
    CPdfTrigger.prototype.GetParentField = function() {
        return this.parentField;
    };
    CPdfTrigger.prototype.SetCallerField = function(oField) {
        this.callerField = oField;
    };
    CPdfTrigger.prototype.GetCallerFiled = function() {
        return this.callerField || this.GetParentField();
    };
    CPdfTrigger.prototype.WriteToBinary = function(memory) {
        let nType = this.GetType();
        switch (nType) {
            case AscPDF.PDF_TRIGGERS_TYPES.mouseUp:
                memory.WriteString("A");
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.mouseDown:
                memory.WriteString("D");
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.mouseEnter:
                memory.WriteString("E");
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.mouseExit:
                memory.WriteString("X");
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.onFocus:
                memory.WriteString("Fo");
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.onBlur:
                memory.WriteString("Bl");
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.keystroke:
                memory.WriteString("K");
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.validate:
                memory.WriteString("V");
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.calculate:
                memory.WriteString("C");
                break;
            case AscPDF.PDF_TRIGGERS_TYPES.format:
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
            return CPdfTrigger.GetName(oTrigger.GetType());
        }
    };

    function CActionGoTo(sPageId, nGoToType, nZoom, aRect) {
        CActionBase.call(this, ACTIONS_TYPES.GoTo);
        this.pageId     = sPageId;
        this.goToType   = nGoToType;
        this.zoom       = nZoom;
        this.rect       = aRect; // top right bottom left
    };
    CActionGoTo.prototype = Object.create(CActionBase.prototype);
	CActionGoTo.prototype.constructor = CActionGoTo;

    CActionGoTo.prototype.Copy = function() {
        return new CActionGoTo(this.GetPageId(), this.GetKind(), this.GetZoom(true), this.GetRect().slice());
    };
    CActionGoTo.prototype.GetZoom = function(bSource) {
        if (this.zoom != null || bSource) {
            return this.zoom;
		}
		else if (this.calcedZoom) {
			return this.calcedZoom;
		}

        let nPageIdx    = this.GetPageIdx();
        let oViewer     = editor.getDocumentRenderer();
        let nNoZoomH    = oViewer.drawingPages[nPageIdx].H / oViewer.zoom;
        let nNoZoomW    = oViewer.drawingPages[nPageIdx].W / oViewer.zoom;

        let nScaleY = oViewer.drawingPages[nPageIdx].H / oViewer.file.pages[nPageIdx].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPageIdx].W / oViewer.file.pages[nPageIdx].W / oViewer.zoom;

        switch (this.goToType) {
            case AscPDF.GOTO_TYPES.xyz: // inherit zoom
                break;
            case AscPDF.GOTO_TYPES.fit:
            case AscPDF.GOTO_TYPES.fitB: { // fit to max of heigth/width
                let nVerZoom = ((oViewer.canvas.height / (nNoZoomH * AscCommon.AscBrowser.retinaPixelRatio)) * 100 >> 0) / 100;
                let nHorZoom = ((oViewer.canvas.width / (nNoZoomW * AscCommon.AscBrowser.retinaPixelRatio)) * 100 >> 0) / 100;

                this.calcedZoom = Math.min(nHorZoom, nVerZoom);
                break;
            }
            case AscPDF.GOTO_TYPES.fitH:
            case AscPDF.GOTO_TYPES.fitBH: { // fit to width
                this.calcedZoom = ((oViewer.canvas.width / (nNoZoomW * AscCommon.AscBrowser.retinaPixelRatio)) * 100 >> 0) / 100;
                break;
            }
            case AscPDF.GOTO_TYPES.fitV:
            case AscPDF.GOTO_TYPES.fitBV: { // fit to heigth
                this.calcedZoom = ((oViewer.canvas.height / (nNoZoomH * AscCommon.AscBrowser.retinaPixelRatio)) * 100 >> 0) / 100;
                break;
            }
            case AscPDF.GOTO_TYPES.fitR: { // fit to rect
                let nRectW = Math.abs((this.rect[0] - this.rect[2]) * nScaleX * AscCommon.AscBrowser.retinaPixelRatio);
                let nRectH = Math.abs((this.rect[3] - this.rect[1]) * nScaleY * AscCommon.AscBrowser.retinaPixelRatio);

                let nVerZoom = ((oViewer.canvas.height / (nRectH)) * 100 >> 0) / 100;
                let nHorZoom = ((oViewer.canvas.width / (nRectW)) * 100 >> 0) / 100;

                let nMinZoom = Math.min(nHorZoom, nVerZoom);
                
                // далее вычисляем ширину с новым потенциальным зумом,
                // если при данных размерах будет добавлен скролл, то вычитаем его ширину и пересчитываем zoom
                let nNewPageW = (oViewer.file.pages[nPageIdx].W * 96 * nMinZoom / oViewer.file.pages[nPageIdx].Dpi) >> 0;
                if (nNewPageW > oViewer.width) {
                    nVerZoom = (((oViewer.canvas.height - oViewer.scrollWidth) / (nRectH)) * 100 >> 0) / 100;
                }
                
                this.calcedZoom = Math.min(nHorZoom, nVerZoom);
            }
        }

        return this.calcedZoom;
    };

    CActionGoTo.prototype.GetPageId = function() {
        return this.pageId;
    };
    CActionGoTo.prototype.GetPageIdx = function() {
        let oPageInfo = AscCommon.g_oTableId.GetById(this.GetPageId());
        return oPageInfo.GetIndex();
    };
    CActionGoTo.prototype.SetPageIdx = function(pageIdx) {
        let oDoc = Asc.editor.getPDFDoc();
		let oPage = oDoc.GetPageInfo(pageIdx);

        this.pageId = oPage.GetId();
    };

    CActionGoTo.prototype.GetKind = function() {
        return this.goToType;
    };
    CActionGoTo.prototype.SetKind = function() {
        this.goToType = kind;
    };

    CActionGoTo.prototype.GetRect = function() {
        return this.rect;
    };
    CActionGoTo.prototype.SetRect = function(rect) {
        this.rect = rect;
    };

    CActionGoTo.prototype.Do = function() {
        let oViewer         = Asc.editor.getDocumentRenderer();
        let oField          = this.GetCallerFiled();
        let oDoc            = oField.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oActionsQueue.SetCurAction(this);
        
        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.GetTriggerType() == PDF_TRIGGERS_TYPES.onFocus && oField != oDoc.activeForm) {
            oActionsQueue.Continue();
            return;
        }
        
        let nPageIdx = this.GetPageIdx();
        if (nPageIdx == -1) {
            oActionsQueue.Continue();
            return;
        }

        let nZoom = this.GetZoom();
        if (nZoom && oViewer.zoom != nZoom)
            oViewer.setZoom(nZoom, true);

        // выставляем смещения
        let yOffset = this.rect[1] != null ? this.rect[1] : 0;
        let xOffset = this.rect[0] != null ? this.rect[0] : 0;

        if ((nZoom && oViewer.zoom != nZoom) || yOffset != undefined && xOffset != undefined || oViewer.currentPage != nPageIdx) {
            let oTr = oDoc.pagesTransform[nPageIdx].invert;
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
        memory.WriteLong(this.GetPageIdx());

        let nKind = this.GetKind();
        memory.WriteByte(nKind);

        switch (nKind) {
            case AscPDF.GOTO_TYPES.xyz:
            case AscPDF.GOTO_TYPES.fitH:
            case AscPDF.GOTO_TYPES.fitV:
            case AscPDF.GOTO_TYPES.fitBH:
            case AscPDF.GOTO_TYPES.fitBV:
            {
                let nFlag = 0;
                let nStartPos = memory.GetCurPosition();
                memory.Skip(4);

                if (this.rect[0] != null) {
                    nFlag |= (1 << 0);
                    memory.WriteDouble(this.rect[0]);
                }
                if (this.rect[1] != null) {
                    nFlag |= (1 << 1);
                    memory.WriteDouble(this.rect[1]);
                }
                if (this.zoom != null) {
                    nFlag |= (1 << 2);
                    memory.WriteDouble(this.zoom);
                }

                // write flags
                let nEndPos = memory.GetCurPosition();
                memory.Seek(nStartPos);
                memory.WriteLong(nFlag);
                memory.Seek(nEndPos);
                break;
            }
            case AscPDF.GOTO_TYPES.fitR:
            {
                memory.WriteDouble(this.rect[0]);
                memory.WriteDouble(this.rect[1]);
                memory.WriteDouble(this.rect[2]);
                memory.WriteDouble(this.rect[3]);
                break;
            }
        }
    };

    function CActionNamed(nType) {
        CActionBase.call(this, ACTIONS_TYPES.Named);
        this.name = nType;
    };
    CActionNamed.prototype = Object.create(CActionBase.prototype);
	CActionNamed.prototype.constructor = CActionNamed;

    CActionNamed.prototype.Copy = function() {
        return new CActionNamed(this.Getname());
    };
    CActionNamed.prototype.GetName = function() {
        return this.name;
    };
    CActionNamed.prototype.SetName = function(type) {
        this.name = type;
    };

    CActionNamed.prototype.Do = function() {
        let Api             = Asc.editor;
        let oViewer         = Api.getDocumentRenderer();
        let oField          = this.GetCallerFiled();
        let oDoc            = oField.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();
		let nPagesCount		= oDoc.GetPagesCount();

        oActionsQueue.SetCurAction(this);

        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.GetTriggerType() == PDF_TRIGGERS_TYPES.onFocus && oField != oDoc.activeForm) {
            oActionsQueue.Continue();
            return;
        }

        switch (this.name) {
            case AscPDF.ACTION_NAMED_TYPES.FirstPage:
                Api.goToPage(0);
                break;
            case AscPDF.ACTION_NAMED_TYPES.NextPage:
                if (oViewer.currentPage + 1 <= nPagesCount)
                    Api.goToPage(oViewer.currentPage + 1);
                break;
            case AscPDF.ACTION_NAMED_TYPES.PrevPage:
                if (oViewer.currentPage - 1 >= 0)
                    Api.goToPage(oViewer.currentPage - 1);
                break;
            case AscPDF.ACTION_NAMED_TYPES.LastPage:
                if (oViewer.currentPage != nPagesCount)
                    Api.goToPage(nPagesCount - 1);
                break;
        }

        oActionsQueue.Continue();
    };

    CActionNamed.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(this.GetType());
        memory.WriteString(this.GetName());
    };

    function CActionURI(sURI) {
        CActionBase.call(this, ACTIONS_TYPES.URI);
        this.uri = sURI;
    };
    CActionURI.prototype = Object.create(CActionBase.prototype);
	CActionURI.prototype.constructor = CActionURI;

    CActionURI.prototype.Copy = function() {
        return new CActionURI(this.GetUri());
    };
    CActionURI.prototype.Do = function() {
        let oField          = this.GetCallerFiled();
        let oDoc            = oField.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oActionsQueue.SetCurAction(this);

        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.GetTriggerType() == PDF_TRIGGERS_TYPES.onFocus && oField != oDoc.activeForm) {
            oActionsQueue.Continue();
            return;
        }

        Asc.editor.sendEvent("asc_onOpenLinkPdfForm", this.uri, this.OpenLink.bind(this), oActionsQueue.Continue.bind(oActionsQueue));
    };
    CActionURI.prototype.GetUri = function() {
        return this.uri;
    };
    CActionURI.prototype.SetUri = function(uri) {
        this.uri = uri;
    };
    CActionURI.prototype.OpenLink = function() {
        window.open(this.uri, "_blank");

        let oField          = this.GetCallerFiled();
        let oDoc            = oField.GetDocument();
        oDoc.GetActionsQueue().Continue();
    };

    CActionURI.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(this.GetType());
        memory.WriteString(this.GetUri());
    };

    function CActionHideShow(bHidden, aFieldsNames) {
        CActionBase.call(this, ACTIONS_TYPES.HideShow);
        this.hidden = bHidden;
        this.names = aFieldsNames;
    };

    CActionHideShow.prototype = Object.create(CActionBase.prototype);
	CActionHideShow.prototype.constructor = CActionHideShow;

    CActionHideShow.prototype.Copy = function() {
        return new CActionHideShow(this.IsHide(), this.GetNames().slice());
    };
    CActionHideShow.prototype.Do = function() {
        let oField          = this.GetCallerFiled();
        let oDoc            = oField.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oActionsQueue.SetCurAction(this);

        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.GetTriggerType() == PDF_TRIGGERS_TYPES.onFocus && oField != oDoc.activeForm) {
            oActionsQueue.Continue();
            return;
        }

        oDoc.HideShowForms(this.hidden, this.names);
    };

    CActionHideShow.prototype.GetNames = function() {
        return this.names.slice();
    };
    CActionHideShow.prototype.SetNames = function(names) {
        this.names = names.slice();
    };
    CActionHideShow.prototype.IsHide = function() {
        return this.hidden;
    };
    CActionHideShow.prototype.SetHide = function(isHide) {
        this.hidden = isHide;
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

    function CActionReset(aFieldsNames, isAllExcept) {
        CActionBase.call(this, ACTIONS_TYPES.ResetForm);
        this.names      = aFieldsNames;
        this.isAllExcept = isAllExcept;
    };
    CActionReset.prototype = Object.create(CActionBase.prototype);
	CActionReset.prototype.constructor = CActionReset;

    CActionReset.prototype.Copy = function() {
        return new CActionReset(this.GetNames().slice(), this.IsAllExcept());
    };
    CActionReset.prototype.Do = function() {
        let oField          = this.GetCallerFiled();
        let oDoc            = oField.GetDocument();
        let oActionsQueue   = oDoc.GetActionsQueue();

        oActionsQueue.SetCurAction(this);

        // если onFocus но форма не активна, то скипаем дейсвтие
        if (this.GetTriggerType() == PDF_TRIGGERS_TYPES.onFocus && oField != oDoc.activeForm) {
            oActionsQueue.Continue();
            return;
        }
            
        oDoc.ResetForms(this.names, this.isAllExcept);
    };

    CActionReset.prototype.GetNames = function() {
        return this.names.slice();
    };
    CActionReset.prototype.SetNames = function(names) {
        this.names = names.slice();
    };
    CActionReset.prototype.IsAllExcept = function() {
        return this.isAllExcept;
    };
    CActionReset.prototype.SetAllExcept = function(isAll) {
        this.isAllExcept = isAll;
    };

    CActionReset.prototype.WriteToBinary = function(memory) {
        memory.WriteByte(this.GetType());

        if (this.isAllExcept)
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
        if (this.GetTriggerType() == PDF_TRIGGERS_TYPES.onFocus && oField != oDoc.activeForm) {
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
    CActionRunScript.prototype.SetScript = function(script) {
        this.script = script;
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
    
		let oApiApp = {
			"alert": function(message) {
				window.alert(message);
			}
		}
		Object.freeze(oApiApp);

        let oApiObjects = AscPDF.Api.Types;
        let aArgsNamesPdfApi = [
            "event",
            "color",
			"app",

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
            "AFMakeNumber"
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
			oApiApp,

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
            oApiFunc["AFRange_Validate"],
            oApiFunc["AFMakeNumber"]
        ];
    
        let funcArgs = aArgsNamesToDelete.concat(aArgsNamesPdfApi);
        
		// make document methods global
		let oApiDoc = oParentDoc.GetDocumentApi();
		const aOwnMethods = Object.getOwnPropertyNames(AscPDF.ApiDocument.prototype).filter(function(key) {
			return key !== "constructor" && typeof oApiDoc[key] === 'function';
		});
		aArgsNamesPdfApi = aArgsNamesPdfApi.concat(aOwnMethods);

		aOwnMethods.forEach(function(key) {
			funcArgs.push(key);
			aArgsPdfApi.push(oApiDoc[key].bind(oApiDoc));
		});

		funcArgs.push(str);

        let func = Function.apply(null, funcArgs);
        func.bind(oApiDoc).apply(null, new Array(aArgsNamesToDelete.length - 1).concat(oApiConsole, aArgsPdfApi));
    }
    
	function extractArguments(str) {
		const start = str.indexOf('(');
		if (start === -1) {
			return [];
		}

		let end = -1;
		let depth = 0;
		let quote = null;
		let escaped = false;

		for (let i = start; i < str.length; i++) {
			const ch = str[i];

			if (quote) {
				if (escaped) {
					escaped = false;
				} else if (ch === '\\') {
					escaped = true;
				} else if (ch === quote) {
					quote = null;
				}
				continue;
			}

			if (ch === '"' || ch === "'") {
				quote = ch;
				continue;
			}

			if (ch === '(') depth++;
			if (ch === ')') depth--;

			if (depth === 0) {
				end = i;
				break;
			}
		}

		if (end === -1) {
			return [];
		}

		const argsString = str.slice(start + 1, end);
		const args = splitTopLevel(argsString);

		return args.map(parseArgument);
	}

	function splitTopLevel(str) {
		const result = [];
		let current = '';
		let quote = null;
		let escaped = false;
		let paren = 0;
		let bracket = 0;
		let brace = 0;

		for (let i = 0; i < str.length; i++) {
			const ch = str[i];

			if (quote) {
				current += ch;

				if (escaped) {
					escaped = false;
				} else if (ch === '\\') {
					escaped = true;
				} else if (ch === quote) {
					quote = null;
				}
				continue;
			}

			if (ch === '"' || ch === "'") {
				quote = ch;
				current += ch;
				continue;
			}

			if (ch === '(') paren++;
			else if (ch === ')') paren--;
			else if (ch === '[') bracket++;
			else if (ch === ']') bracket--;
			else if (ch === '{') brace++;
			else if (ch === '}') brace--;

			if (ch === ',' && paren === 0 && bracket === 0 && brace === 0) {
				result.push(current.trim());
				current = '';
				continue;
			}

			current += ch;
		}

		if (current.trim() !== '') {
			result.push(current.trim());
		}

		return result;
	}

	function parseArgument(arg) {
		if (arg === 'true') return true;
		if (arg === 'false') return false;
		if (arg === 'null') return null;

		if (arg !== '' && !Number.isNaN(Number(arg))) {
			return Number(arg);
		}

		if (
			(arg.startsWith('"') && arg.endsWith('"')) ||
			(arg.startsWith("'") && arg.endsWith("'"))
		) {
			return arg.slice(1, -1);
		}

		return arg;
	}

	function getJsonActionInfo(action) {
		let actionInfo = {};

		switch (action.GetType()) {
			case AscPDF.ACTIONS_TYPES.JavaScript:
				actionInfo["S"] = AscPDF.ACTIONS_TYPES.JavaScript;
				actionInfo["JS"] = action.GetScript();
				break;
			case AscPDF.ACTIONS_TYPES.ResetForm:
				actionInfo["S"] = AscPDF.ACTIONS_TYPES.ResetForm;
				actionInfo["Fields"] = action.GetNames();
				actionInfo["Flags"] = Number(action.IsAllExcept());
				break;
			case AscPDF.ACTIONS_TYPES.URI:
				actionInfo["S"] = AscPDF.ACTIONS_TYPES.URI;
				actionInfo["URI"] = action.GetUri();
				break;
			case AscPDF.ACTIONS_TYPES.HideShow:
				actionInfo["S"] = AscPDF.ACTIONS_TYPES.HideShow;
				actionInfo["H"] = action.IsHide();
				actionInfo["T"] = action.GetNames();
				break;
			case AscPDF.ACTIONS_TYPES.GoTo:
				actionInfo["S"] = AscPDF.ACTIONS_TYPES.GoTo;
				actionInfo["page"] = action.GetPageIdx();
				actionInfo["pageId"] = action.GetPageId();
				actionInfo["kind"] = action.GetKind();
				actionInfo["zoom"] = action.GetZoom();
				let aRect = action.GetRect();
				actionInfo["top"] = aRect[1];
				actionInfo["right"] = aRect[2];
				actionInfo["bottom"] = aRect[3];
				actionInfo["left"] = aRect[0];
				break;
			case AscPDF.ACTIONS_TYPES.Named:
				actionInfo["S"] = AscPDF.ACTIONS_TYPES.Named;
				actionInfo["N"] = action.GetName();
				break;
			default:
				// If the type is not recognized, add handling or skip
				break;
		}

		return actionInfo;
	}

    if (!window["AscPDF"])
	    window["AscPDF"] = {};
    
    window["AscPDF"].CPdfTriggers      = CPdfTriggers;
    window["AscPDF"].CPdfTrigger       = CPdfTrigger;
    window["AscPDF"].CActionGoTo        = CActionGoTo;
    window["AscPDF"].CActionNamed       = CActionNamed;
    window["AscPDF"].CActionURI         = CActionURI;
    window["AscPDF"].CActionHideShow    = CActionHideShow;
    window["AscPDF"].CActionReset       = CActionReset;
    window["AscPDF"].CActionRunScript   = CActionRunScript;
    window["AscPDF"].extractArguments   = extractArguments;
    window["AscPDF"].getJsonActionInfo	= getJsonActionInfo;
	
    window["AscPDF"].ACTIONS_TYPES          = ACTIONS_TYPES;
    window["AscPDF"].PDF_TRIGGERS_TYPES   = PDF_TRIGGERS_TYPES;

})();

