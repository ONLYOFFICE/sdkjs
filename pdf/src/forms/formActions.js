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

	PDF_TRIGGERS_TYPES["mouseUp"]		= PDF_TRIGGERS_TYPES.mouseUp;
	PDF_TRIGGERS_TYPES["mouseDown"]		= PDF_TRIGGERS_TYPES.mouseDown;
	PDF_TRIGGERS_TYPES["mouseEnter"]	= PDF_TRIGGERS_TYPES.mouseEnter;
	PDF_TRIGGERS_TYPES["mouseExit"]		= PDF_TRIGGERS_TYPES.mouseExit;
	PDF_TRIGGERS_TYPES["onFocus"]		= PDF_TRIGGERS_TYPES.onFocus;
	PDF_TRIGGERS_TYPES["onBlur"]		= PDF_TRIGGERS_TYPES.onBlur;
	PDF_TRIGGERS_TYPES["keystroke"]		= PDF_TRIGGERS_TYPES.keystroke;
	PDF_TRIGGERS_TYPES["validate"]		= PDF_TRIGGERS_TYPES.validate;
	PDF_TRIGGERS_TYPES["calculate"]		= PDF_TRIGGERS_TYPES.calculate;
	PDF_TRIGGERS_TYPES["format"]		= PDF_TRIGGERS_TYPES.format;
	
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

	ACTIONS_TYPES["mouseUp"]		= ACTIONS_TYPES.mouseUp;
	ACTIONS_TYPES["Unknown"]		= ACTIONS_TYPES.Unknown;
	ACTIONS_TYPES["GoTo"]			= ACTIONS_TYPES.GoTo;
	ACTIONS_TYPES["GoToR"]			= ACTIONS_TYPES.GoToR;
	ACTIONS_TYPES["GoToE"]			= ACTIONS_TYPES.GoToE;
	ACTIONS_TYPES["Launch"]			= ACTIONS_TYPES.Launch;
	ACTIONS_TYPES["Thread"]			= ACTIONS_TYPES.Thread;
	ACTIONS_TYPES["URI"]			= ACTIONS_TYPES.URI;
	ACTIONS_TYPES["Sound"]			= ACTIONS_TYPES.Sound;
	ACTIONS_TYPES["Movie"]			= ACTIONS_TYPES.Movie;
	ACTIONS_TYPES["HideShow"]		= ACTIONS_TYPES.HideShow;
	ACTIONS_TYPES["Named"]			= ACTIONS_TYPES.Named;
	ACTIONS_TYPES["SubmitForm"]		= ACTIONS_TYPES.SubmitForm;
	ACTIONS_TYPES["ResetForm"]		= ACTIONS_TYPES.ResetForm;
	ACTIONS_TYPES["ImportData"]		= ACTIONS_TYPES.ImportData;
	ACTIONS_TYPES["JavaScript"]		= ACTIONS_TYPES.JavaScript;
	ACTIONS_TYPES["SetOCGState"]	= ACTIONS_TYPES.SetOCGState;
	ACTIONS_TYPES["Rendition"]		= ACTIONS_TYPES.Rendition;
	ACTIONS_TYPES["Trans"]			= ACTIONS_TYPES.Trans;
	ACTIONS_TYPES["GoTo3DView"]		= ACTIONS_TYPES.GoTo3DView;

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
            case AscPDF.GOTO_TYPES.fitB: { // fit to page
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
                
                // next we calculate width with new potential zoom,
                // if scroll will be added at these dimensions, then we subtract its width and recalculate zoom
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
    CActionGoTo.prototype.SetKind = function(kind) {
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
        let oDoc            = oField ? oField.GetDocument() : Asc.editor.getPDFDoc();
        let oActionsQueue   = oDoc.GetActionsQueue();

		if (oField) {
	        oActionsQueue.SetCurAction(this);
		}
        
        // if onFocus but form is not active, then skip action
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

        // set offsets
        let yOffset = this.rect[1] != null ? this.rect[1] : 0;
        let xOffset = this.rect[0] != null ? this.rect[0] : 0;

        if ((nZoom && oViewer.zoom != nZoom) || yOffset != undefined && xOffset != undefined || oViewer.currentPage != nPageIdx) {
            let oTr = oDoc.pagesTransform[nPageIdx].invert;
            let oPos = oTr.TransformPoint(xOffset, yOffset);

            oViewer.disabledPaintOnScroll = true; // disable drawing on scroll
            oViewer.scrollToXY(oViewer.scrollY + oPos.y, oViewer.scrollX + oPos.x);
            oViewer.disabledPaintOnScroll = false;
            oViewer.needRedraw = true; // at the end of Actions we'll perform repaint
        }

		if (oField) {
	        oActionsQueue.Continue();
		}
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

        // if onFocus but form is not active, then skip action
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

        // if onFocus but form is not active, then skip action
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
		function normalizeUrl(uri) {
			const value = String(uri || "").trim();

			if (!value) {
				return "";
			}

			if (/^[a-z][a-z0-9+.-]*:/i.test(value)) {
				return value;
			}

			return `https://${value}`;
		}

		window.open(normalizeUrl(this.uri), "_blank");

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

        // if onFocus but form is not active, then skip action
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

        // if onFocus but form is not active, then skip action
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
        this.bContinueAfterEval = true; // disable on asynchronous operations
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

        // if onFocus but form is not active, then skip action
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
		if (/\bimport\s*\(/.test(str)) {
			throw new Error("Security: dynamic import() is not allowed in scripts");
		}
		if (/\bimport\s+/.test(str)) {
			throw new Error("Security: static import is not allowed in scripts");
		}

		let oShadowedGlobals = {
			"Function":         undefined,
			"eval":             undefined,
			"window":           undefined,
			"globalThis":       undefined,
			"self":             undefined,
			"document":         undefined,
			"AscDesktopEditor": undefined,
			"setTimeout":       undefined,
			"setInterval":      undefined,
			"XMLHttpRequest":   undefined,
			"Promise":          undefined,
			"console":          undefined
		};
		let aArgsNamesToDelete = Object.keys(oShadowedGlobals);

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
		oShadowedGlobals["console"] = oApiConsole;

		let aArgsValuesToDelete = Object.values(oShadowedGlobals);

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

		oParentDoc.eventsStack.push(oParentDoc.event);
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

		let protoFunc = Object.getPrototypeOf(function(){});
		let protoFuncGen = null;
		try {
			protoFuncGen = eval("Object.getPrototypeOf(function*(){})");
		} catch (err) {
			protoFuncGen = null;
		}
		const normalConstructor = protoFunc.constructor;
		const generatorNext = protoFuncGen ? protoFuncGen.prototype.next : null;
		protoFunc.constructor = function(){};
		if (protoFuncGen) {
			protoFuncGen.prototype.next = function(){};
		}

		try {
			let func = Function.apply(null, funcArgs);
			func.bind(oApiDoc).apply(null, aArgsValuesToDelete.concat(aArgsPdfApi));
		} finally {
			protoFunc.constructor = normalConstructor;
			if (protoFuncGen)
				protoFuncGen.prototype.next = generatorNext;
		}
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

    window["AscPDF"]["ACTIONS_TYPES"]		= window["AscPDF"].ACTIONS_TYPES		= ACTIONS_TYPES;
    window["AscPDF"]["PDF_TRIGGERS_TYPES"]	= window["AscPDF"].PDF_TRIGGERS_TYPES	= PDF_TRIGGERS_TYPES;

})();

