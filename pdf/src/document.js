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

// TODO: Временно
var CPresentation = CPresentation || function(){};

(function(){

    /**
	 * event properties.
	 * @typedef {Object} oEventPr
	 * @property {string} [change=""] - A string specifying the change in value that the user has just typed. A JavaScript may replace part or all of this string with different characters. The change may take the form of an individual keystroke or a string of characters (for example, if a paste into the field is performed).
	 * @property {boolean} [rc=true] - Used for validation. Indicates whether a particular event in the event chain should succeed. Set to false to prevent a change from occurring or a value from committing. The default is true.
	 * @property {object} [target=undefined] - The target object that triggered the event. In all mouse, focus, blur, calculate, validate, and format events, it is the Field object that triggered the event. In other events, such as page open and close, it is the Doc or this object.
	 * @property {any} value ->
     *  This property has different meanings for different field events:
     *    For the Field/Validate event, it is the value that the field contains when it is committed. For a combo box, it is the face value, not the export value.  
     *    For a Field/Calculate event, JavaScript should set this property. It is the value that the field should take upon completion of the event.    
     *    For a Field/Format event, JavaScript should set this property. It is the value used when generating the appearance for the field. By default, it contains the value that the user has committed. For a combo box, this is the face value, not the export value.   
     *    For a Field/Keystroke event, it is the current value of the field. If modifying a text field, for example, this is the text in the text field before the keystroke is applied.
     *    For Field/Blur and Field/Focus events, it is the current value of the field. During these two events, event.value is read only. That is, the field value cannot be changed by setting event.value.
     * @property {boolean} willCommit -  Verifies the current keystroke event before the data is committed. It can be used to check target form field values to verify, for example, whether character data was entered instead of numeric data. JavaScript sets this property to true after the last keystroke event and before the field is validated.
	 */

    let AscPDF = window["AscPDF"];

    function CCalculateInfo(oDoc) {
        this.names = [];
        this.document = oDoc;
        this.isInProgress = false;
        this.sourceField = null; // поле вызвавшее calculate
    };

    CCalculateInfo.prototype.AddFieldToOrder = function(sName) {
        if (this.names.includes(sName) == false)
            this.names.push(sName);
    };
    CCalculateInfo.prototype.RemoveFieldFromOrder = function(sName) {
        let nIdx = this.names.indexOf(sName);
        if (nIdx != -1) {
            this.names.splice(nIdx, 1);
        }
    };
    CCalculateInfo.prototype.SetIsInProgress = function(bValue) {
        this.isInProgress = bValue;
    };
    CCalculateInfo.prototype.IsInProgress = function() {
        return this.isInProgress;
    };
    CCalculateInfo.prototype.SetCalculateOrder = function(aNames) {
        this.names = aNames.slice();
    };
    /**
	 * Sets field to calc info, which caused the recalculation.
     * Note: This field cannot be changed in scripts.
	 * @memberof CBaseField
	 * @typeofeditors ["PDF"]
	 */
    CCalculateInfo.prototype.SetSourceField = function(oField) {
        this.sourceField = oField;
    };
    CCalculateInfo.prototype.GetSourceField = function() {
        return this.sourceField;
    };
	
	/**
	 * Main class for working with PDF structure
	 * @constructor
	 */
    function CPDFDoc(viewer) {
        this.rootFields = new Map(); // root поля форм
        this.widgets    = []; // непосредственно сами поля, которые отрисовываем (дочерние без потомков)
        this.annots     = [];

        this.theme = new AscFormat.CTheme();
        this.actionsInfo = new CActionQueue(this);
        this.calculateInfo = new CCalculateInfo(this);
        this.fieldsToCommit = [];
        this.event = {};
        this.AutoCorrectSettings = new AscCommon.CAutoCorrectSettings();
        Object.defineProperties(this.event, {
            "change": {
                set(value) {
                    if (value != null && value.toString)
                        this._change = value.toString();
                },
                get() {
                    return this._change;
                }
            }
        });

        this._parentsMap = {}; // map при открытии форм
        this.api = this.GetDocumentApi();
		
        // internal
        this._id = AscCommon.g_oIdCounter.Get_NewId();

        this.History    = AscCommon.History;
		this.Spelling   = new AscCommonWord.CDocumentSpellChecker();
        this.Viewer     = viewer;

        this.annotsHidden = false;
		
		this.fontLoader = AscCommon.g_font_loader;
		this.defaultFontsLoaded = -1; // -1 не загружены и не грузим, 0 - грузим, 1 - загружены
		this.fontLoaderCallbacks = [];
    }

    /////////// методы для открытия //////////////
    CPDFDoc.prototype.AddFieldToChildsMap = function(oField, nParentIdx) {
        if (this._parentsMap[nParentIdx] == null)
            this._parentsMap[nParentIdx] = [];

        this._parentsMap[nParentIdx].push(oField);
    };
    CPDFDoc.prototype.GetParentsMap = function() {
        return this._parentsMap;
    };
    CPDFDoc.prototype.OnEndFormsActions = function() {
        let oViewer = editor.getDocumentRenderer();
        if (oViewer.needRedraw == true) { // отключали отрисовку на скроле из ActionToGo, поэтому рисуем тут
            oViewer._paint();
            oViewer.needRedraw = false;
        }
        else {
            oViewer._paint();
        }
    };
    CPDFDoc.prototype.FillFormsParents = function(aParentsInfo) {
        let oChilds = this.GetParentsMap();
        let oParents = {};

        for (let i = 0; i < aParentsInfo.length; i++) {
            let nIdx = aParentsInfo[i]["i"];
            if (!oChilds[nIdx])
                continue;

            let sType = oChilds[nIdx][0].GetType();

            let oParent = private_createField(aParentsInfo[i]["name"], sType, undefined, undefined, this);
            if (aParentsInfo[i]["value"] != null)
                oParent.SetApiValue(aParentsInfo[i]["value"]);
            if (aParentsInfo[i]["Parent"] != null)
                this.AddFieldToChildsMap(oParent, aParentsInfo[i]["Parent"]);
            if (aParentsInfo[i]["defaultValue"] != null)
                oParent.SetDefaultValue(aParentsInfo[i]["defaultValue"]);
            oParents[nIdx] = oParent;

            this.rootFields.set(oParent.GetPartialName(), oParent);
        }

        for (let nParentIdx in oParents) {
            oChilds[nParentIdx].forEach(function(child) {
                oParents[nParentIdx].AddKid(child);
            });
        }
    };
    CPDFDoc.prototype.OnAfterFillFormsParents = function() {
        let bInberitValue = false;
        let value;
        for (let i = 0; i < this.widgets.length; i++) {
            let oField = this.widgets[i];
            if ((oField.GetPartialName() == null || oField.GetApiValue(bInberitValue) == null) && oField.GetParent()) {
                value = oField.GetParent().GetApiValue();
                if (value != null && value.toString)
                    value = value.toString();

                oField.SetValue(value);
            }
        }
    };
    CPDFDoc.prototype.FillButtonsIconsOnOpen = function() {
        let oViewer = editor.getDocumentRenderer();
        let oDoc = this;

        oViewer.IsOpenFormsInProgress = true;
        for (let i = 0; i < oViewer.pagesInfo.pages.length; i++) {
            let oPage = oViewer.drawingPages[i];

            let w = (oPage.W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
            let h = (oPage.H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

            let oFile = oViewer.file;
            let aIconsInfo = oFile.nativeFile["getButtonIcons"](i, w, h);

            if (aIconsInfo["View"] == null)
                return;
                
            let aIconsToLoad = [];
            let oIconsMap = {};

            // load images
            for (let nIcon = 0; nIcon < aIconsInfo["View"].length; nIcon++) {
                let canvas  = document.createElement("canvas");
                let ctx     = canvas.getContext("2d");
                let nWidth  = aIconsInfo["View"][nIcon]["w"];
                let nHeight = aIconsInfo["View"][nIcon]["h"];
                
                canvas.width    = nWidth;
                canvas.height   = nHeight;

                let nRetValue = aIconsInfo["View"][nIcon]["retValue"];

                let supportImageDataConstructor = (AscCommon.AscBrowser.isIE && !AscCommon.AscBrowser.isIeEdge) ? false : true;
                let mappedBuffer    = new Uint8ClampedArray(oFile.memory().buffer, nRetValue, 4 * nWidth * nHeight);
                let imageData       = null;

                if (supportImageDataConstructor) {
                    imageData = new ImageData(mappedBuffer, nWidth, nHeight);
                }
                else {
                    imageData = ctx.createImageData(nWidth, nHeight);
                    imageData.data.set(mappedBuffer, 0);                    
                }

                if (ctx) {
                    ctx.putImageData(imageData, 0, 0);
                }
                
                oFile.free(nRetValue);

                aIconsToLoad.push({
                    Image: {
                        width: nWidth,
                        height: nHeight,
                    },
                    src: canvas.toDataURL()
                });

                for (let nField = 0; nField < aIconsInfo["MK"].length; nField++) {
                    if (aIconsInfo["MK"][nField]["I"] == aIconsInfo["View"][nIcon]["j"]) {
                        aIconsInfo["MK"][nField]["I"] = aIconsToLoad[aIconsToLoad.length - 1];
                    }
                    else if (aIconsInfo["MK"][nField]["RI"] == aIconsInfo["View"][nIcon]["j"]) {
                        aIconsInfo["MK"][nField]["RI"] = aIconsToLoad[aIconsToLoad.length - 1];
                    }
                    else if (aIconsInfo["MK"][nField]["IX"] == aIconsInfo["View"][nIcon]["j"]) {
                        aIconsInfo["MK"][nField]["IX"] = aIconsToLoad[aIconsToLoad.length - 1];
                    }
                        //aIconsToLoad[aIconsToLoad.length - 1].fields.push(this.GetFieldBySourceIdx(aIconsInfo["MK"][nField]["i"]));
                }
            }

            editor.ImageLoader.LoadImagesWithCallback(aIconsToLoad.map(function(info) {
                return info.src;
            }), function() {

                oViewer.IsOpenFormsInProgress = true;
                for (let nField = 0; nField < aIconsInfo["MK"].length; nField++) {
                    let oField = oDoc.GetFieldBySourceIdx(aIconsInfo["MK"][nField]["i"]);

                    oField.Recalculate();
                    if (aIconsInfo["MK"][nField]["I"]) {
                        oField.AddImage(aIconsInfo["MK"][nField]["I"]);
                    }
                    if (aIconsInfo["MK"][nField]["RI"]) {
                        oField.AddImage(aIconsInfo["MK"][nField]["RI"], AscPDF.APPEARANCE_TYPE.rollover);
                    }
                    if (aIconsInfo["MK"][nField]["IX"]) {
                        oField.AddImage(aIconsInfo["MK"][nField]["IX"], AscPDF.APPEARANCE_TYPE.mouseDown);
                    }
                }

                oViewer.IsOpenFormsInProgress = false;
            });
        }
    };
    CPDFDoc.prototype.GetFieldBySourceIdx = function(nIdx) {
        for (let i = 0; i < this.widgets.length; i++) {
            if (this.widgets[i]._apIdx == nIdx) {
                return this.widgets[i];
            }
        }
    };
    ////////////////////////////////////


    CPDFDoc.prototype.GetId = function() {
        return this._id;
    };
    CPDFDoc.prototype.Get_Id = function() {
        return this._id;
    };
    CPDFDoc.prototype.GetDrawingDocument = function() {
		if (!editor || !editor.WordControl)
			return null;
		
		return editor.WordControl.m_oDrawingDocument;
	};
	CPDFDoc.prototype.GetDocumentRenderer = function() {
		if (!editor)
			return null;
		
		return editor.getDocumentRenderer();
	};
    CPDFDoc.prototype.CommitFields = function() {
        this.skipHistoryOnCommit = true;
        this.fieldsToCommit.forEach(function(field) {
            field.Commit();
        });
        
        this.ClearFieldsToCommit();
        this.skipHistoryOnCommit = false;
    };
    CPDFDoc.prototype.ClearCache = function(nPageIndex) {
        let oViewer = editor.getDocumentRenderer();

        if (oViewer.pagesInfo.pages[nPageIndex].fields != null) {
            oViewer.pagesInfo.pages[nPageIndex].fields.forEach(function(field) {
                field.ClearCache();
            });
        }
        if (oViewer.pagesInfo.pages[nPageIndex].annots != null) {
            oViewer.pagesInfo.pages[nPageIndex].annots.forEach(function(annot) {
                annot.ClearCache();
            });
        }
        
        oViewer.file.pages[nPageIndex].fieldsAPInfo = null;
        oViewer.file.pages[nPageIndex].annotsAPInfo = null;
    };
    CPDFDoc.prototype.IsNeedSkipHistory = function() {
        return !!this.skipHistoryOnCommit;
    };
    CPDFDoc.prototype.AddFieldToCommit = function(oField) {
        this.fieldsToCommit.push(oField);
    };
    CPDFDoc.prototype.ClearFieldsToCommit = function() {
        this.fieldsToCommit = [];
    };
    CPDFDoc.prototype.SelectNextField = function() {
        let oViewer         = editor.getDocumentRenderer();
        let aWidgetForms    = this.widgets;
        let oActionsQueue   = this.GetActionsQueue();
        let isNeedRedraw    = false;
		
		let _t = this;
		if (!this.checkDefaultFieldFonts(function(){_t.SelectNextField();}))
			return;
		
        if (aWidgetForms.length == 0)
            return;

        let nCurIdx = this.widgets.indexOf(this.activeForm);
        let oCurForm = this.widgets[nCurIdx];
        let oNextForm = this.widgets[nCurIdx + 1] || this.widgets[0];

        if (oCurForm && oNextForm) {
            if (oCurForm.IsNeedCommit()) {
                isNeedRedraw = true;

                let isValid = true;
                if ([AscPDF.FIELD_TYPES.text, AscPDF.FIELD_TYPES.combobox].includes(oCurForm.GetType())) {
                    isValid = oCurForm.DoValidateAction(oCurForm.GetValue());
                }

                if (isValid) {
                    oCurForm.needValidate = false; 
                    oCurForm.Commit();
                    if (this.event["rc"] == true) {
                        this.DoCalculateFields(oCurForm);
                        this.AddFieldToCommit(oCurForm);
                        this.CommitFields();
                    }
                }
                else {
                    oNextForm = null;
                    oCurForm.UndoNotAppliedChanges();
                    if (oCurForm.IsChanged() == false) {
                        oCurForm.SetDrawFromStream(true);
                    }
                }

                oCurForm.SetNeedCommit(false);
            }
            else if (oCurForm.IsChanged() == false) {
                isNeedRedraw = true;
                oCurForm.SetDrawFromStream(true);
            }

            oCurForm.SetDrawHighlight(true);
            oCurForm.Blur();
        }
        
        if (!oNextForm)
            return;

        this.activeForm = oNextForm;
        
        oNextForm.SetDrawHighlight(false);
        
        if (oNextForm.IsNeedDrawFromStream() == true && oNextForm.GetType() != AscPDF.FIELD_TYPES.button) {
            isNeedRedraw = true;
            oNextForm.SetDrawFromStream(false);
            oNextForm.AddToRedraw();
        }
        
        oNextForm.onFocus();

        let callBackAfterFocus = function() {
            switch (oNextForm.GetType()) {
                case AscPDF.FIELD_TYPES.text:
                case AscPDF.FIELD_TYPES.combobox:
                    oViewer.Api.WordControl.m_oDrawingDocument.UpdateTargetFromPaint = true;
                    oViewer.Api.WordControl.m_oDrawingDocument.m_lCurrentPage = 0;
                    oViewer.Api.WordControl.m_oDrawingDocument.m_lPagesCount = oViewer.file.pages.length;
                    oViewer.Api.WordControl.m_oDrawingDocument.showTarget(true);
                    oViewer.Api.WordControl.m_oDrawingDocument.TargetStart();
                    if (oNextForm.content.IsSelectionUse())
                        oNextForm.content.RemoveSelection();
    
                    oNextForm.content.GetElement(0).MoveCursorToStartPos();
                    oNextForm.content.RecalculateCurPos();
                    
                    break;
                default:
                    oViewer.Api.WordControl.m_oDrawingDocument.TargetEnd();
                    break;
            }
        };

        
        if (false == oNextForm.IsInSight())
            this.NavigateToField(oNextForm);
        else {
            oViewer._paint();
        }
        
        if (oActionsQueue.IsInProgress() == true)
            oActionsQueue.callBackAfterFocus = callBackAfterFocus;
        else
            callBackAfterFocus();
    };
    CPDFDoc.prototype.SelectPrevField = function() {
        let oViewer         = editor.getDocumentRenderer();
        let aWidgetForms    = this.widgets;
        let oActionsQueue   = this.GetActionsQueue();
        let isNeedRedraw    = false;
		
		let _t = this;
		if (!this.checkDefaultFieldFonts(function(){_t.SelectNextField();}))
			return;
		
        if (aWidgetForms.length == 0)
            return;

        let nCurIdx = this.widgets.indexOf(this.activeForm);
        let oCurForm = this.widgets[nCurIdx];
        let oNextForm = this.widgets[nCurIdx - 1] || this.widgets[this.widgets.length - 1];

        if (oCurForm && oNextForm) {
            if (oCurForm.IsNeedCommit()) {
                isNeedRedraw = true;
                oCurForm.Commit();
            }
            else if (oCurForm.IsChanged() == false) {
                isNeedRedraw = true;
                oCurForm.SetDrawFromStream(true);
            }

            oCurForm.SetDrawHighlight(true);
            oCurForm.Blur();
        }
        
        if (!oNextForm)
            return;
        
        this.activeForm = oNextForm;
        oNextForm.SetDrawHighlight(false);
        
        if (oNextForm.IsNeedDrawFromStream() == true && oNextForm.GetType() != AscPDF.FIELD_TYPES.button) {
            isNeedRedraw = true;
            oNextForm.SetDrawFromStream(false);
            oNextForm.AddToRedraw();
        }
        
        oNextForm.onFocus();

        let callBackAfterFocus = function() {
            switch (oNextForm.GetType()) {
                case AscPDF.FIELD_TYPES.text:
                case AscPDF.FIELD_TYPES.combobox:
                    oViewer.Api.WordControl.m_oDrawingDocument.UpdateTargetFromPaint = true;
                    oViewer.Api.WordControl.m_oDrawingDocument.m_lCurrentPage = 0;
                    oViewer.Api.WordControl.m_oDrawingDocument.m_lPagesCount = oViewer.file.pages.length;
                    oViewer.Api.WordControl.m_oDrawingDocument.showTarget(true);
                    oViewer.Api.WordControl.m_oDrawingDocument.TargetStart();
                    if (oNextForm.content.IsSelectionUse())
                        oNextForm.content.RemoveSelection();
    
                    oNextForm.content.GetElement(0).MoveCursorToStartPos();
                    oNextForm.content.RecalculateCurPos();
                    
                    break;
                default:
                    oViewer.Api.WordControl.m_oDrawingDocument.TargetEnd();
                    break;
            }
        };

        
        if (false == oNextForm.IsInSight())
            this.NavigateToField(oNextForm);
        else {
            oViewer._paint();
        }
        
        if (oActionsQueue.IsInProgress() == true)
            oActionsQueue.callBackAfterFocus = callBackAfterFocus;
        else
            callBackAfterFocus();
    };
    CPDFDoc.prototype.NavigateToField = function(oField) {
        let oViewer = editor.getDocumentRenderer();
        let aOrigRect = oField.GetOrigRect();
        let nPage = oField.GetPage();
        
        let nBetweenPages = oViewer.betweenPages / (oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H);

        let nPageHpx = (oViewer.drawingPages[nPage].H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
        let nPageWpx = (oViewer.drawingPages[nPage].W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

        // находим видимый размер от страницы в исходных размерах 
        let nViewedH = (oViewer.canvas.height / nPageHpx) * oViewer.file.pages[nPage].H;
        let nViewedW = (oViewer.canvas.width / nPageWpx) * oViewer.file.pages[nPage].W;
        
        // выставляем смещение до формы страницу
        let yOffset = aOrigRect[1] + (aOrigRect[3] - aOrigRect[1]) / 2 - nViewedH / 2 + nBetweenPages;
        let xOffset = aOrigRect[0] + (aOrigRect[2] - aOrigRect[0]) / 2 - nViewedW / 2;

        oViewer.navigateToPage(nPage, yOffset > 0 ? yOffset : undefined, xOffset > 0 ? xOffset : undefined);
    };
    CPDFDoc.prototype.EnterDownActiveField = function() {
        let oViewer = editor.getDocumentRenderer();
        let oField = this.activeForm;

        if (oField.IsNeedDrawHighlight())
            return;

        if ([AscPDF.FIELD_TYPES.checkbox, AscPDF.FIELD_TYPES.radiobutton].includes(oField.GetType())) {
            oField.onMouseUp();
        }
        else {
            oField.SetDrawHighlight(true);
            oField.UpdateScroll && oField.UpdateScroll(false); // убираем скролл

            if (oField.IsNeedRevertShiftView()) {
                oField.RevertContentViewToOriginal();
            }

            if (oField.IsNeedCommit()) {
                let isValid = true;
                if ([AscPDF.FIELD_TYPES.text, AscPDF.FIELD_TYPES.combobox].includes(oField.GetType())) {
                    isValid = oField.DoValidateAction(oField.GetValue());
                }
                if (isValid) {
                    oField.needValidate = false; 
                    oField.Commit();
                    if (this.event["rc"] == true) {
                        this.DoCalculateFields(oField);
                        this.AddFieldToCommit(oField);
                        this.CommitFields();
                    }
                }
                else {
                    oField.UndoNotAppliedChanges();
                    if (oField.IsChanged() == false) {
                        oField.SetDrawFromStream(true);
                    }
                }

                oField.SetNeedCommit(false);
            }
            else if (oField.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Format) && oField.GetValue() != "") {
                oField.AddToRedraw();

                if (oField.IsChanged() == false)
                    oField.SetDrawFromStream(true);
            }
            
            if (oField && oField.content && oField.content.IsSelectionUse()) {
                oField.content.RemoveSelection();
                oViewer.onUpdateOverlay();
            }

            oViewer.Api.WordControl.m_oDrawingDocument.TargetEnd(); // убираем курсор
        }

        // суть в том, что мы рисуем background только когда форма активна, если неактивна - рисуем highlight вместо него.
        if (oField.GetBackgroundColor())
            oField.AddToRedraw();

        oViewer._paint();
    };
    CPDFDoc.prototype.OnExitFieldByClick = function() {
        let oViewer         = editor.getDocumentRenderer();
        let oActiveForm     = this.activeForm;
        let oActionsQueue   = this.GetActionsQueue();

        oActiveForm.UpdateScroll && oActiveForm.UpdateScroll(false); // убираем скрол
        oActiveForm.SetDrawHighlight(true);

        // суть в том, что мы рисуем background только когда форма активна, если неактивна - рисуем highlight вместо него.
        if (oActiveForm.GetBackgroundColor())
            oActiveForm.AddToRedraw();

        // если чекбокс то выходим сразу
        if ([AscPDF.FIELD_TYPES.checkbox, AscPDF.FIELD_TYPES.radiobutton].includes(oActiveForm.GetType())) {
            oActiveForm.SetPressed(false);
            oActiveForm.SetHovered(false);
            oActiveForm.AddToRedraw();
            oActiveForm.Blur();
            
            if (oActionsQueue.IsInProgress() == false) {
                oViewer._paint();
            }

            return;
        }
        
        if (oActiveForm.IsNeedCommit()) {
            if (oActiveForm.IsNeedRevertShiftView()) {
                oActiveForm.RevertContentViewToOriginal();
            }

            let isValid = true;
            if ([AscPDF.FIELD_TYPES.text, AscPDF.FIELD_TYPES.combobox].includes(oActiveForm.GetType())) {
                isValid = oActiveForm.DoValidateAction(oActiveForm.GetValue());
            }

            if (isValid) {
                oActiveForm.needValidate = false; 
                oActiveForm.Commit();
                if (this.event["rc"] == true) {
                    this.DoCalculateFields(oActiveForm);
                    this.AddFieldToCommit(oActiveForm);
                    this.CommitFields();
                }
            }
            else {
                oActiveForm.UndoNotAppliedChanges();
                if (oActiveForm.IsChanged() == false) {
                    oActiveForm.SetDrawFromStream(true);
                }
            }

            oActiveForm.SetNeedCommit(false);
        }
        else {
            if (oActiveForm.IsChanged() == false) {
                oActiveForm.SetDrawFromStream(true);
                oActiveForm.AddToRedraw();
            }

            if (oActiveForm.IsNeedRevertShiftView()) {
                oActiveForm.RevertContentViewToOriginal();
            }
            
            if ([AscPDF.FIELD_TYPES.text, AscPDF.FIELD_TYPES.combobox].includes(oActiveForm.GetType())) {
                if (oActiveForm.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Format)) {
                    oActiveForm.AddToRedraw();
                }
            }
        }
        
        oActiveForm.Blur();
        oViewer.Api.WordControl.m_oDrawingDocument.TargetEnd();
        if (oActionsQueue.IsInProgress() == false && this.mouseDownField == null) {
            oViewer._paint();
        }

        if (oActiveForm && oActiveForm.content && oActiveForm.content.IsSelectionUse() && this.mouseDownField == null) {
            oActiveForm.content.RemoveSelection();
            oViewer.onUpdateOverlay();
        }
    };
    CPDFDoc.prototype.OnMouseDownField = function(oField, event) {
        let oViewer         = editor.getDocumentRenderer();
        let oActionsQueue   = this.GetActionsQueue();
		
		let _t = this;
		if (!this.checkDefaultFieldFonts(function(){_t.OnMouseDownField(oField, event)}))
			return;

        // суть в том, что мы рисуем background только когда форма активна, если неактивна - рисуем highlight вместо него.
        if (oField.GetBackgroundColor())
            oField.AddToRedraw();

        switch (oField.GetType())
        {
            case AscPDF.FIELD_TYPES.text:
            case AscPDF.FIELD_TYPES.combobox:
                oField.SetDrawHighlight(false);
                oField.onMouseDown(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y, event);
                oViewer.onUpdateOverlay();
                break;
            case AscPDF.FIELD_TYPES.listbox:
                oField.SetDrawHighlight(false);
                oField.onMouseDown(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y, event);
                
                oViewer.Api.WordControl.m_oDrawingDocument.TargetEnd();
                break;
            case AscPDF.FIELD_TYPES.button:
            case AscPDF.FIELD_TYPES.radiobutton:
            case AscPDF.FIELD_TYPES.checkbox:
                oField.SetDrawHighlight(false);
                oField.onMouseDown(event);
                break;
        }

        if (oActionsQueue.IsInProgress() == false && oViewer.pagesInfo.pages[oField.GetPage()].needRedrawForms)
            oViewer._paint();

        // нажали мышь - запомнили координаты и находимся ли на ссылке
        // при выходе за epsilon на mouseMove - сэмулируем нажатие
        // так что тут только курсор
        var cursorType;
        if (this.activeForm)
        {
            switch (this.activeForm.GetType())
            {
                case AscPDF.FIELD_TYPES.text:
                    cursorType = "text";
                    break;
                case AscPDF.FIELD_TYPES.combobox:
                    cursorType = "text";
                    break;
                default:
                    cursorType = "pointer";
                    break;
            }
        }

        oViewer.setCursorType(cursorType);
    };
    CPDFDoc.prototype.OnMouseDown = function(e) {
        let oViewer         = editor.getDocumentRenderer();
        let oDrawingObjects = oViewer.DrawingObjects;
        let oDrDoc          = this.GetDrawingDocument();
        let IsOnDrawer      = oViewer.Api.isDrawInkMode();
        let IsOnEraser      = oViewer.Api.isEraseInkMode();

        oViewer.mouseDownLinkObject = oViewer.getPageLinkByMouse();
        let oMouseDownField         = oViewer.getPageFieldByMouse();
        let oMouseDownAnnot         = oViewer.getPageAnnotByMouse();

        if (IsOnEraser) {
            if (oMouseDownAnnot && oMouseDownAnnot.IsInk())
                this.EraseInk(oMouseDownAnnot);
            return;
        }

        // если курсор меняется на resize, то клик по нему выходит за область поля или аннотации, отслеживаем этот момент и не убираем поле/аннотацию из активных
        let oPos    = oDrDoc.ConvertCoordsFromCursor2(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
        let X       = oPos.X;
        let Y       = oPos.Y;

        let bInRect = oDrawingObjects.updateCursorType(oPos.DrawPage, X, Y, e, false);
        
        if (IsOnDrawer == false && false == IsOnEraser) {
            if (bInRect) {
                if (oMouseDownField && (this.mouseDownAnnot != oMouseDownAnnot && !oDrawingObjects.selectedObjects.includes(this.mouseDownAnnot))) {
                    this.mouseDownField = oMouseDownField;
                    this.mouseDownAnnot = null;
                }
                else if (oMouseDownAnnot && !oDrawingObjects.selectedObjects.includes(oMouseDownAnnot)) {
                    this.mouseDownAnnot = oMouseDownAnnot;
                    this.mouseDownField = null;
                }
            }
            else {
                if (oMouseDownField) {
                    this.mouseDownField = oMouseDownField;
                    this.mouseDownAnnot = null;
                }
                else if (oMouseDownAnnot) {
                    this.mouseDownAnnot = oMouseDownAnnot;
                    this.mouseDownField = null;
                }
                else {
                    this.mouseDownAnnot = null;
                    this.mouseDownField = null;
                }
            }
        }

        if (IsOnDrawer == true || oViewer.Api.isMarkerFormat) {
            this.mouseDownAnnot = null;
            this.mouseDownField = null
            oDrawingObjects.OnMouseDown(e, X, Y, oPos.DrawPage);
            return;
        }
        
        // если попали в другую форму или никуда, то выход из текущей формы
        if (this.activeForm && this.mouseDownField != this.activeForm) {
            this.OnExitFieldByClick();
        }
        
        if (this.mouseDownField)
            this.OnMouseDownField(this.mouseDownField, e); 
        else if (this.mouseDownAnnot)
            this.mouseDownAnnot.onMouseDown(e);
        else
            oDrawingObjects.OnMouseDown(e, X, Y, oPos.DrawPage);

        if (!oViewer.MouseHandObject && (!oViewer.mouseDownLinkObject))
        {
            // ждать смысла нет
            oViewer.isMouseMoveBetweenDownUp = true;
            oViewer.onMouseDownEpsilon();
        }
        else if (this.mouseDownAnnot) {
            oViewer.onUpdateOverlay();
        }
        
        if (false == oDrawingObjects.selectedObjects.includes(this.mouseDownAnnot)) {
            oDrawingObjects.resetSelection();
        }
    };
    CPDFDoc.prototype.EraseInk = function(oInk) {
        this.RemoveAnnot(oInk.GetId());
    };

    CPDFDoc.prototype.OnMouseMove = function(x, y, e) {
        let oViewer         = editor.getDocumentRenderer();
        let oDrawingObjects = oViewer.DrawingObjects;
        let oDrDoc          = this.GetDrawingDocument();
        let oAPI            = oViewer.Api;
        
        let oPos    = oDrDoc.ConvertCoordsFromCursor2(x, y);
        let X       = oPos.X;
        let Y       = oPos.Y;

        let pageObject = oViewer.getPageByCoords3(AscCommon.global_mouseEvent.X - oViewer.x, AscCommon.global_mouseEvent.Y - oViewer.y);
        console.log(pageObject);
        
        if (oViewer.isMouseDown)
        {
            if (oAPI.isEraseInkMode()) {
                let oMouseMoveAnnot = oViewer.getPageAnnotByMouse();
                if (oMouseMoveAnnot && oMouseMoveAnnot.IsInk()) {
                    this.EraseInk(oMouseMoveAnnot);
                    return;
                }
            }
            else if (oAPI.isDrawInkMode() || this.mouseDownAnnot) {
                e.IsLocked = true;
                oViewer.overlay.ClearAll = true;
                oViewer.overlay.max_x = 0;
                oViewer.overlay.max_y = 0;
                

                if (oDrawingObjects.arrTrackObjects.length != 0 && this.mouseDownAnnot && this.mouseDownAnnot.IsInk() == true)
                    oDrawingObjects.updateCursorType(oPos.DrawPage, X, Y, e, false);

                oDrawingObjects.OnMouseMove(e, X, Y, oPos.DrawPage);
            }
            else if (this.activeForm)
            {
                // селект текста внутри формы
                if (this.activeForm.GetType() == AscPDF.FIELD_TYPES.text || this.activeForm.GetType() == AscPDF.FIELD_TYPES.combobox)
                {
                    this.activeForm.SelectionSetEnd(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y, e);
                    if (this.activeForm.content.IsSelectionEmpty() == false) {
                        oAPI.WordControl.m_oDrawingDocument.TargetEnd();
                    }
                    else {
                        oAPI.WordControl.m_oDrawingDocument.TargetStart();
                        oAPI.WordControl.m_oDrawingDocument.showTarget(true);
                    }
                    
                    oViewer.onUpdateOverlay();
                }
                // отрисовка нажатого/отжатого состояния кнопок/чекбоксов
                else if ([AscPDF.FIELD_TYPES.button, AscPDF.FIELD_TYPES.checkbox, AscPDF.FIELD_TYPES.radiobutton].includes(this.activeForm.GetType())) {
                    let mouseMoveFieldObject = oViewer.getPageFieldByMouse();
                    if (mouseMoveFieldObject != this.activeForm && this.activeForm.IsHovered()) {
                        this.activeForm.SetHovered(false);
                        this.activeForm.DrawUnpressed();
                    }
                    else if (mouseMoveFieldObject == this.activeForm && this.activeForm.IsHovered() == false) {
                        this.activeForm.SetHovered(true);
                        this.activeForm.DrawPressed();
                    }
                }
            }
            else if (oViewer.mouseDownLinkObject)
            {
                // не меняем курсор с "ссылочного", если зажимали на ссылке
                oViewer.setCursorType("pointer");
            }
            else
            {
                // даже если не двигали еще и ждем eps, все равно курсор меняем на зажатый
                oViewer.setCursorType(AscCommon.Cursors.Grabbing);
            }
        }
        else
        {
            if (oAPI.isInkDrawerOn())
                return;
            
            // просто водим мышкой - тогда смотрим, на ссылке или поле, чтобы выставить курсор
            var mouseMoveLinkObject = oViewer.getPageLinkByMouse();
            var mouseMoveFieldObject = oViewer.getPageFieldByMouse();
            var mouseMoveAnnotObject = oViewer.getPageAnnotByMouse();
            
            if (mouseMoveFieldObject && mouseMoveFieldObject != oViewer.mouseMoveFieldObject) {
                mouseMoveFieldObject._needDrawHoverBorder = true;
                if (oViewer.mouseMoveFieldObject)
                    oViewer.mouseMoveFieldObject._needDrawHoverBorder = false;

                oViewer.mouseMoveFieldObject && oViewer.mouseMoveFieldObject.onMouseExit();
                oViewer.mouseMoveFieldObject = mouseMoveFieldObject;
                mouseMoveFieldObject.onMouseEnter();
            }
            else if (mouseMoveFieldObject == null && oViewer.mouseMoveFieldObject) {
                oViewer.mouseMoveFieldObject.onMouseExit();
                oViewer.mouseMoveFieldObject._needDrawHoverBorder = false;
                oViewer.mouseMoveFieldObject = null;
            }


            let cursorType = "default";
            if (oViewer.MouseHandObject)
                cursorType = AscCommon.Cursors.Grab;

            if (mouseMoveLinkObject)
                cursorType = "pointer";
            else if (mouseMoveFieldObject)
            {
                switch (mouseMoveFieldObject.GetType())
                {
                    case AscPDF.FIELD_TYPES.text:
                        cursorType = "text";
                        break;
                    case AscPDF.FIELD_TYPES.combobox:
                        if (!pageObject)
                            return null;

                        if (pageObject.x >= mouseMoveFieldObject._markRect.x1 && pageObject.x <= mouseMoveFieldObject._markRect.x2 && pageObject.y >= mouseMoveFieldObject._markRect.y1 && pageObject.y <= mouseMoveFieldObject._markRect.y2 && mouseMoveFieldObject._options.length != 0) {
                            cursorType = "pointer";
                        }
                        else
                            cursorType = "text";
                        break;
                    default:
                        cursorType = "pointer";
                }
            }
            else if (mouseMoveAnnotObject && (mouseMoveAnnotObject.GetType() != AscPDF.ANNOTATIONS_TYPES.Ink && mouseMoveAnnotObject.IsTextMarkup() == false)) {
            	cursorType = "move";
            }
            else if (mouseMoveAnnotObject && mouseMoveAnnotObject.IsTextMarkup()) {
            	cursorType = "default";
            }

            oViewer.setCursorType(cursorType);

            if (!mouseMoveAnnotObject || mouseMoveAnnotObject.IsInk() == true)
                oDrawingObjects.updateCursorType(oPos.DrawPage, X, Y, e, false);
        }
    };
    CPDFDoc.prototype.OnMouseUp = function(e) {
        let oViewer         = editor.getDocumentRenderer();
        let oDrawingObjects = oViewer.DrawingObjects;
        let oDrDoc          = this.GetDrawingDocument();

        if (oViewer.Api.isEraseInkMode())
            return;

        let bUpdateOverlay = false;
        if (oDrawingObjects.arrTrackObjects.length != 0) {
            bUpdateOverlay = true;
        }

        let oMouseUpField = oViewer.getPageFieldByMouse();
		let oMouseUpAnnot = oViewer.getPageAnnotByMouse();
        
        let oPos    = oDrDoc.ConvertCoordsFromCursor2(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.X);
        let X       = oPos.X;
        let Y       = oPos.Y;

        let pageObject = oViewer.getPageByCoords3(AscCommon.global_mouseEvent.X - oViewer.x, AscCommon.global_mouseEvent.Y - oViewer.y);

        // если рисование, то просто заканчиваем его
        if (oViewer.Api.isDrawInkMode()) {
            oDrawingObjects.OnMouseUp(e, X, Y, pageObject.index);
            if (true == bUpdateOverlay) {
                oViewer.overlay.ClearAll = true;
                oViewer.overlay.max_x = 0;
                oViewer.overlay.max_y = 0;
                oViewer.onUpdateOverlay();
            }
            return;
        }

        // в приоритете обрабатываем формы или аннотации, потом уже ссылки
        if (this.mouseDownAnnot || this.mouseDownField) {
            oViewer.isMouseMoveBetweenDownUp = false;
            if (oViewer.MouseHandObject)
                oViewer.MouseHandObject.Active = false;
            oViewer.mouseDownLinkObject = null;
        }

        oDrawingObjects.OnMouseUp(e, X, Y, pageObject.index);
        
        if (this.mouseDownField)
        {
            if (oMouseUpField == this.mouseDownField)
                this.OnMouseUpField(oMouseUpField, e);
            else if (this.mouseDownField.GetType() == AscPDF.FIELD_TYPES.button) {
                this.mouseDownField.SetPressed(false);
            }
        }
        else if (this.mouseDownAnnot) {
            
            if (bUpdateOverlay == true) {
                oViewer.overlay.ClearAll = true;
                oViewer.overlay.max_x = 0;
                oViewer.overlay.max_y = 0;
                oViewer.onUpdateOverlay();
            }
            if (this.mouseDownAnnot == oMouseUpAnnot)
                this.mouseDownAnnot.onMouseUp();
        }
        else if (!oViewer.isMouseMoveBetweenDownUp)
        {
            oViewer.setCursorType(AscCommon.Cursors.Grab);

            // делаем клик в логическом документе, чтобы сбросить селект, если он был
            let pageObjectLogic = oViewer.getPageByCoords2(AscCommon.global_mouseEvent.X - oViewer.x, AscCommon.global_mouseEvent.Y - oViewer.y);
            oViewer.file.onMouseDown(pageObjectLogic.index, pageObjectLogic.x, pageObjectLogic.y);
            oViewer.file.onMouseUp(pageObjectLogic.index, pageObjectLogic.x, pageObjectLogic.y);
        }
        
        e.IsLocked = false;
        this.UpdateCopyCutState();
    };

    CPDFDoc.prototype.OnMouseUpField = function(oField, event) {
        let oViewer         = editor.getDocumentRenderer();
        let oActionsQueue   = this.GetActionsQueue();

        if (global_mouseEvent.ClickCount == 2 && (oField.GetType() == AscPDF.FIELD_TYPES.text || oField.GetType() == AscPDF.FIELD_TYPES.combobox))
        {
            oField.content.SelectAll();
            if (oField.content.IsSelectionEmpty() == false)
                oViewer.Api.WordControl.m_oDrawingDocument.TargetEnd();
            else
                oField.content.RemoveSelection();

            oViewer.onUpdateOverlay();
        }

        switch (oField.GetType())
        {
            case AscPDF.FIELD_TYPES.checkbox:
            case AscPDF.FIELD_TYPES.radiobutton:
                oViewer.Api.WordControl.m_oDrawingDocument.TargetEnd();
                
                oField.onMouseUp();

                if (oField.IsNeedCommit()) {
                    let oDoc = oField.GetDocument();
                    oDoc.DoCalculateFields();
                    oDoc.CommitFields();
                }
                break;
            default:
                oField.onMouseUp();
                break;
        }

        if (oActionsQueue.IsInProgress() == false && oViewer.pagesInfo.pages[oField.GetPage()].needRedrawForms)
            oViewer._paint();
    };
    CPDFDoc.prototype.DoUndo = function() {
        let oViewer = editor.getDocumentRenderer();

        if (AscCommon.History.Can_Undo())
        {
            this.currInkInDrawingProcess = null;

            let oCurPoint = AscCommon.History.Points[AscCommon.History.Index];
            let nCurPoindIdx = AscCommon.History.Index;

            oViewer.isOnUndoRedo = true;
            
            AscCommon.History.Undo();
            let oParentForm = oCurPoint.Additional.FormFilling;
            if (oParentForm) {
                // если форма активна, то изменения (undo) применяются только для неё
                // иначе для всех с таким именем (для checkbox и radiobutton всегда применяем для всех)
                // так же применяем для всех, если добрались до точки, общей для всех форм, а не примененнёые изменения удаляем (для всех кроме checkbox и radiobutton)
                if ((this.activeForm == null || oCurPoint.Additional && oCurPoint.Additional.CanUnion === false || nCurPoindIdx == 0) || 
                    (oParentForm.GetType() == AscPDF.FIELD_TYPES.checkbox || oParentForm.GetType() == AscPDF.FIELD_TYPES.radiobutton)) {
                        oViewer.Api.WordControl.m_oDrawingDocument.TargetEnd(); // убираем курсор
                        
                        if (oParentForm.GetType() == AscPDF.FIELD_TYPES.listbox) {
                            oParentForm.Commit(null);
                        }
                        // для радиокнопок храним все изменения, т.к. значения не идентичны для каждой формы из группы
                        // восстанавлием все состояния из истории полностью, поэтому значение формы не нужно применять.
                        else if (AscPDF.FIELD_TYPES.radiobutton != oParentForm.GetType())
                            oParentForm.Commit();

                        // вызываем calculate actions
                        let oDoc = oParentForm.GetDocument();
                        oDoc.DoCalculateFields();
                        oDoc.CommitFields();

                        // выход из формы
                        if (this.activeForm)
                        {
                            this.activeForm.SetDrawHighlight(true);
                            this.activeForm = null;
                        }
                }

                oParentForm.SetNeedRecalc(true);
                oParentForm.AddToRedraw();

                // Перерисуем страницу, на которой произошли изменения
                // oViewer._paint();
                oViewer.onUpdateOverlay();
            }
            
            oViewer._paint();
            oViewer.onUpdateOverlay();
            oViewer.isOnUndoRedo = false;
        }
    };
    CPDFDoc.prototype.DoRedo = function() {
        let oViewer = editor.getDocumentRenderer();

        if (AscCommon.History.Can_Redo())
        {
            this.currInkInDrawingProcess = null;
            oViewer.isOnUndoRedo = true;

            AscCommon.History.Redo();
            let nCurPoindIdx = AscCommon.History.Index;
            let oCurPoint = AscCommon.History.Points[nCurPoindIdx];

            let oParentForm = oCurPoint.Additional.FormFilling;
            if (oParentForm) {
                // если мы в форме, то изменения (undo) применяются только для неё
                // иначе для всех с таким именем
                if (this.activeForm == null || oCurPoint.Additional && oCurPoint.Additional.CanUnion === false) {
                    oViewer.Api.WordControl.m_oDrawingDocument.TargetEnd(); // убираем курсор
                        
                    if (oParentForm.GetType() == AscPDF.FIELD_TYPES.listbox) {
                        oParentForm.Commit(null);
                    }
                    // для радиокнопок храним все изменения, т.к. значения не идентичны для каждой формы из группы
                    // восстанавлием все состояния из истории полностью, поэтому значение формы не нужно применять.
                    else if (AscPDF.FIELD_TYPES.radiobutton != oParentForm.GetType())
                        oParentForm.Commit();

                    // вызываем calculate actions
                    let oDoc = oParentForm.GetDocument();
                    oDoc.DoCalculateFields();
                    oDoc.CommitFields();

                    // выход из формы
                    if (this.activeForm)
                    {
                        this.activeForm.SetDrawHighlight(true);
                        this.activeForm = null;
                    }
                }

                oParentForm.SetNeedRecalc(true);
                oParentForm.AddToRedraw()
                
                // Перерисуем страницу, на которой произошли изменения
                // oViewer._paint();
                oViewer.onUpdateOverlay();
            }

            oViewer._paint();
            oViewer.onUpdateOverlay();
            oViewer.isOnUndoRedo = false;
        }
    };
    CPDFDoc.prototype.UpdateUndoRedo = function() {
		editor.sync_CanUndoCallback(this.History.Can_Undo());
		editor.sync_CanRedoCallback(this.History.Can_Redo());
    };
    CPDFDoc.prototype.SetEvent = function(oEventPr) {
        if (oEventPr["target"] != null && oEventPr["target"] != this.event["target"])
            this.event["target"] = oEventPr["target"];

        if (oEventPr["rc"] != null)
            this.event["rc"] = oEventPr["rc"];
        else
            this.event["rc"] = true;

        if (oEventPr["change"] != null && oEventPr["change"] != this.event["change"])
            this.event["change"] = oEventPr["change"];
            
        if (oEventPr["value"] != null && oEventPr["value"] != this.event["value"])
            this.event["value"] = oEventPr["value"];

        if (oEventPr["willCommit"] != null)
            this.event["willCommit"] = oEventPr["willCommit"];

        if (oEventPr["willCommit"] != null)
            this.event["willCommit"] = oEventPr["willCommit"];

        if (oEventPr["selStart"] != null)
            this.event["selStart"] = oEventPr["selStart"];

        if (oEventPr["selEnd"] != null)
            this.event["selEnd"] = oEventPr["selEnd"];
    };
    CPDFDoc.prototype.SetWarningInfo = function(oInfo) {
        this.warningInfo = oInfo;
    };
    CPDFDoc.prototype.GetWarningInfo = function() {
        return this.warningInfo;
    };

    CPDFDoc.prototype.DoCalculateFields = function(oSourceField) {
        // при изменении любого поля (с коммитом) вызывается calculate у всех
        let oThis = this;
        this.calculateInfo.SetIsInProgress(true);
        this.calculateInfo.SetSourceField(oSourceField);
        this.calculateInfo.names.forEach(function(name) {
            let oField = oThis.GetField(name);

            let oFormatTrigger = oField.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Calculate);
            let oActionRunScript = oFormatTrigger ? oFormatTrigger.GetActions()[0] : null;

            if (oActionRunScript) {
                oThis.activeForm = oField;
                oActionRunScript.RunScript();
                if (oField.IsNeedCommit()) {
                    oField.SetNeedRecalc(true);
                    oThis.fieldsToCommit.push(oField);
                }
            }
        });
        this.calculateInfo.SetIsInProgress(false);
        this.calculateInfo.SetSourceField(null);
    };

    CPDFDoc.prototype.GetCalculateInfo = function() {
        return this.calculateInfo;
    };

    CPDFDoc.prototype.GetActionsQueue = function() {
        return this.actionsInfo;
    };
        
    /**
	 * Adds a new page to the active document.
	 * @memberof CPDFDoc
	 * @typeofeditors ["PDF"]
     * @param {number} [nPos] - (optional) The page after which to add the new page in a 1-based page numbering
     * system. The default is the last page of the document. Use 0 to add a page before the
     * first page. An invalid page range is truncated to the valid range of pages.
     * @param {points} [nWidth=612] - (optional) The width of the page in points. The default value is 612.
     * @param {points} [nHeight=792] - (optional) The height of the page in points. The default value is 792.
	 * @returns {boolean}
	 */
    CPDFDoc.prototype.AddPage = function(nPos, nWidth, nHeight) {
        let oViewer = editor.getDocumentRenderer();
        let oFile   = oViewer.file;

        if (nPos === undefined || -1 === nPos)
            nPos = oFile.pages.length;
        if (nWidth === undefined)
            nWidth = 612;
        if (nHeight === undefined)
            nHeight = 792;

        oFile.pages.splice(nPos, 0, {
            W: nWidth,
            H: nHeight,
            fonts: [],
            Dpi: 72
        });
	
		oViewer.drawingPages.splice(nPos, 0, {
			X : 0,
			Y : 0,
			W : (oFile.pages[nPos].W * 96 / oFile.pages[nPos].Dpi) >> 0,
			H : (oFile.pages[nPos].H * 96 / oFile.pages[nPos].Dpi) >> 0,
			Image : undefined
		});

        if (oViewer.pagesInfo.pages.length == 0)
            oViewer.pagesInfo.setCount(1);
        else
            oViewer.pagesInfo.pages.splice(nPos, 0, new AscPDF.CPageInfo());

        for (let nPage = nPos + 1; nPage < oViewer.pagesInfo.pages.length; nPage++) {
            if (oViewer.pagesInfo.pages[nPage].fields) {
                oViewer.pagesInfo.pages[nPage].fields.forEach(function(field) {
                    field.SetPage(nPage);
                });
            }
            if (oViewer.pagesInfo.pages[nPage].annots) {
                oViewer.pagesInfo.pages[nPage].annots.forEach(function(annot) {
                    annot.SetPage(nPage);
                });
            }
        }
            
        oViewer.resize();
        oViewer.sendEvent("onPagesCount", oFile.pages.length);
    };
    /**
	 * Adds an interactive field to document.
	 * @memberof CPDFDoc
	 * @typeofeditors ["PDF"]
     * @param {String} cName - The name of the new field to create.
     * @param {"button" | "checkbox" | "combobox" | "listbox" | "radiobutton" | "signature" | "text"} cFieldType - The type of form field to create.
     * @param {Number} nPageNum - The 0-based index of the page to which to add the field.
     * @param {Array} aCoords - An array of four numbers in rotated user space that specifies the size and placement
        of the form field. These four numbers are the coordinates of the bounding rectangle,
        in the following order: upper-left x, upper-left y, lower-right x and lower-right y 
	 * @returns {AscPDF.CBaseField}
	 */
    CPDFDoc.prototype.AddField = function(cName, cFieldType, nPageNum, aCoords) {
        function checkValidParams(cFieldType, nPageNum, aCoords) {
            if (Object.values(AscPDF.FIELD_TYPES).includes(cFieldType) == false)
                return false;
            if (typeof(nPageNum) !== "number" || nPageNum < 0)
                return false;
            let isValidRect = true;
            if (Array.isArray(aCoords)) {
                for (let i = 0; i < 4; i++) {
                    if (typeof(aCoords[i]) != "number") {
                        isValidRect = false;
                        break;
                    }
                }
            }
            else
                isValidRect = false;

            if (!isValidRect)
                return false;
        }
        if (false == checkValidParams(cFieldType, nPageNum, aCoords))
            return null;

        let oViewer = editor.getDocumentRenderer();
        let nScaleY = oViewer.drawingPages[nPageNum].H / oViewer.file.pages[nPageNum].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPageNum].W / oViewer.file.pages[nPageNum].W / oViewer.zoom;

        let aScaledCoords = [aCoords[0] * nScaleX, aCoords[1] * nScaleY, aCoords[2] * nScaleX, aCoords[3] * nScaleY];

        let oPagesInfo = oViewer.pagesInfo;
        if (!oPagesInfo.pages[nPageNum])
            return null;
        
        let oField = private_createField(cName, cFieldType, nPageNum, aScaledCoords, this);
        if (!oField)
            return null;

        oField._origRect = aCoords;

        this.widgets.push(oField);
        oField.SetNeedRecalc(true);

        if (oPagesInfo.pages[nPageNum].fields == null) {
            oPagesInfo.pages[nPageNum].fields = [];
        }
        oPagesInfo.pages[nPageNum].fields.push(oField);

        if (AscCommon.History.IsOn() == true)
            AscCommon.History.TurnOff();

        if (oViewer.IsOpenFormsInProgress == false) {
            oField.SyncField();
            oField.SetDrawFromStream(false);
        }

        return oField;
    };

    /**
	 * Adds an interactive field to document.
	 * @memberof CPDFDoc
	 * @typeofeditors ["PDF"]
     * @param {object} oProps - Annot props 
	 * @returns {AscPDF.CAnnotationBase}
	 */
    CPDFDoc.prototype.AddAnnot = function(oProps) {
        let oViewer = editor.getDocumentRenderer();
        let nPageNum = oProps.page;

        let oPagesInfo = oViewer.pagesInfo;
        if (!oPagesInfo.pages[nPageNum])
            return null;
        
        let oAnnot = CreateAnnotByProps(oProps, this);

        this.annots.push(oAnnot);
        oAnnot.SetNeedRecalc && oAnnot.SetNeedRecalc(true);

        oAnnot.SetDisplay(this.IsAnnotsHidden() ? window["AscPDF"].Api.Objects.display["hidden"] : window["AscPDF"].Api.Objects.display["visible"]);

        if (oPagesInfo.pages[nPageNum].annots == null) {
            oPagesInfo.pages[nPageNum].annots = [];
        }
        oPagesInfo.pages[nPageNum].annots.push(oAnnot);

        if (AscCommon.History.IsOn() == true)
            AscCommon.History.TurnOff();

        if (oViewer.IsOpenAnnotsInProgress == false) {
            this.CreateNewHistoryPoint();
            this.History.Add(new CChangesPDFDocumentAddItem(this, this.annots.length - 1, [oAnnot]));
            this.TurnOffHistory();
        }
        
        // if (oViewer.IsOpenFormsInProgress == false) {
        //     oAnnot.SetDrawFromStream(false);
        // }

        oAnnot.AddToRedraw();
        return oAnnot;
    };
    CPDFDoc.prototype.SetHighlight = function(r, g, b, opacity) {
        this.HighlightColor = {
            r: r,
            g: g,
            b: b,
            a: opacity
        };

        let oViewer         = editor.getDocumentRenderer();
        let oFile           = oViewer.file;
        let aSelQuads       = oFile.getSelectionQuads();

        if (aSelQuads.length == 0)
            return;

        for (let nInfo = 0; nInfo < aSelQuads.length; nInfo++) {
            let nPage   = aSelQuads[nInfo].page;
            let aQuads  = aSelQuads[nInfo].quads;

            let aAllPoints = [];
            aQuads.forEach(function(rect) {
                aAllPoints = aAllPoints.concat(rect);
            });

            let aMinRect = getMinRect(aAllPoints);
            let MinX = aMinRect[0];
            let MinY = aMinRect[1];
            let MaxX = aMinRect[2];
            let MaxY = aMinRect[3];

            let oProps = {
                rect:       [MinX - 3, MinY - 1, MaxX + 3, MaxY + 1],
                page:       nPage,
                name:       AscCommon.CreateGUID(),
                type:       AscPDF.ANNOTATIONS_TYPES.Highlight,
                hidden:     false
            }

            let oAnnot = this.AddAnnot(oProps);

            oAnnot.SetQuads(aQuads);
            oAnnot.SetStrokeColor([r/255, g/255, b/255]);
        }

        editor.sendEvent("asc_onMarkerFormatChanged", AscPDF.ANNOTATIONS_TYPES.Highlight, false);
        editor.SetMarkerFormat(AscPDF.ANNOTATIONS_TYPES.Highlight, false);
    };
    CPDFDoc.prototype.SetUnderline = function(r, g, b, opacity) {
        this.UnderlineColor = {
            r: r,
            g: g,
            b: b,
            a: opacity
        };

        let oViewer         = editor.getDocumentRenderer();
        let oFile           = oViewer.file;
        let aSelQuads       = oFile.getSelectionQuads();
        if (aSelQuads.length == 0)
            return;

        for (let nInfo = 0; nInfo < aSelQuads.length; nInfo++) {
            let nPage   = aSelQuads[nInfo].page;
            let aQuads  = aSelQuads[nInfo].quads;

            let aAllPoints = [];
            aQuads.forEach(function(rect) {
                aAllPoints = aAllPoints.concat(rect);
            });

            let aMinRect = getMinRect(aAllPoints);
            let MinX = aMinRect[0];
            let MinY = aMinRect[1];
            let MaxX = aMinRect[2];
            let MaxY = aMinRect[3];

            let oProps = {
                rect:       [MinX - 3, MinY - 1, MaxX + 3, MaxY + 1],
                page:       nPage,
                name:       AscCommon.CreateGUID(),
                type:       AscPDF.ANNOTATIONS_TYPES.Underline,
                hidden:     false
            }

            let oAnnot = this.AddAnnot(oProps);

            oAnnot.SetQuads(aQuads);
            oAnnot.SetStrokeColor([r/255, g/255, b/255]);
        }

        editor.sendEvent("asc_onMarkerFormatChanged", AscPDF.ANNOTATIONS_TYPES.Underline, false);
        editor.SetMarkerFormat(AscPDF.ANNOTATIONS_TYPES.Underline, false);
    };
    CPDFDoc.prototype.SetStrikeout = function(r, g, b, opacity) {
        this.StrikeoutColor = {
            r: r,
            g: g,
            b: b,
            a: opacity
        };

        let oViewer         = editor.getDocumentRenderer();
        let oFile           = oViewer.file;
        let aSelQuads       = oFile.getSelectionQuads();
        if (aSelQuads.length == 0)
            return;

        for (let nInfo = 0; nInfo < aSelQuads.length; nInfo++) {
            let nPage   = aSelQuads[nInfo].page;
            let aQuads  = aSelQuads[nInfo].quads;

            let aAllPoints = [];
            aQuads.forEach(function(rect) {
                aAllPoints = aAllPoints.concat(rect);
            });

            let aMinRect = getMinRect(aAllPoints);
            let MinX = aMinRect[0];
            let MinY = aMinRect[1];
            let MaxX = aMinRect[2];
            let MaxY = aMinRect[3];

            let oProps = {
                rect:       [MinX - 3, MinY - 1, MaxX + 3, MaxY + 1],
                page:       nPage,
                name:       AscCommon.CreateGUID(),
                type:       AscPDF.ANNOTATIONS_TYPES.Strikeout,
                hidden:     false
            }

            let oAnnot = this.AddAnnot(oProps);

            oAnnot.SetQuads(aQuads);
            oAnnot.SetStrokeColor([r/255, g/255, b/255]);
        }

        editor.sendEvent("asc_onMarkerFormatChanged", AscPDF.ANNOTATIONS_TYPES.Strikeout, false);
        editor.SetMarkerFormat(AscPDF.ANNOTATIONS_TYPES.Strikeout, false);
    };
    CPDFDoc.prototype.GetMarkerColor = function(nType) {
        switch (nType) {
            case AscPDF.ANNOTATIONS_TYPES.Highlight:
                return this.HighlightColor;
            case AscPDF.ANNOTATIONS_TYPES.Underline:
                return this.UnderlineColor;
            case AscPDF.ANNOTATIONS_TYPES.Strikeout:
                return this.StrikeoutColor;
        }

        return null;
    };
    CPDFDoc.prototype.AddComment = function(AscCommentData) {
        let oViewer     = editor.getDocumentRenderer();
        let pageObject  = oViewer.getPageByCoords3(AscCommon.global_mouseEvent.X - oViewer.x, AscCommon.global_mouseEvent.Y - oViewer.y);
        let nGrScale    = AscCommon.AscBrowser.retinaPixelRatio * (96 / oViewer.file.pages[pageObject.index].Dpi);
        let posToAdd    = this.anchorPositionToAdd ? this.anchorPositionToAdd : {x: 10, y: 10};
        
        let X2 = posToAdd.x + 40 / nGrScale;
        let Y2 = posToAdd.y + 40 / nGrScale;

        let oProps = {
            rect:       [posToAdd.x, posToAdd.y, X2, Y2],
            page:       pageObject.index,
            name:       AscCommon.CreateGUID(),
            type:       AscPDF.ANNOTATIONS_TYPES.Text,
            author:     AscCommentData.m_sUserName,
            modDate:    AscCommentData.m_sOOTime,
            contents:   AscCommentData.m_sText,
            hidden:     false
        }

        this.anchorPositionToAdd = null;

        let oAnnot;
        if (this.mouseDownAnnot && this.mouseDownAnnot.IsComment() == false) {
            oAnnot = CreateAnnotByProps(oProps, this);
            if (this.mouseDownAnnot.GetContents()) {
                let newCommentData = new AscCommon.CCommentData();
                newCommentData.Read_FromAscCommentData(AscCommentData);

                let curCommentData = new AscCommon.CCommentData();
                curCommentData.Read_FromAscCommentData(this.mouseDownAnnot.GetAscCommentData());
                curCommentData.Add_Reply(newCommentData);

                this.EditComment(this.mouseDownAnnot.GetId(), curCommentData);
            }
            else
                this.mouseDownAnnot.SetContents(oAnnot);
        }
        else {
            oAnnot = this.AddAnnot(oProps);
            editor.sendEvent("asc_onAddComment", oAnnot.GetId(), AscCommentData);    
        }
        
        return oAnnot;
    };
    CPDFDoc.prototype.CreateNewHistoryPoint = function() {
        if (AscCommon.History.IsOn() == false)
            AscCommon.History.TurnOn();
        AscCommon.History.Create_NewPoint();
    };
    CPDFDoc.prototype.EditComment = function(Id, CommentData) {
        let oAnnotToEdit = this.annots.find(function(annot) {
            return annot.GetId() === Id;
        });

        let oCurData = oAnnotToEdit.GetAscCommentData();

        this.CreateNewHistoryPoint();
        this.History.Add(new CChangesPDFCommentData(oAnnotToEdit, oCurData, CommentData));
        
        oAnnotToEdit.EditCommentData(CommentData);

        editor.sync_ChangeCommentData(Id, CommentData);
    };
    CPDFDoc.prototype.TurnOffHistory = function() {
        if (AscCommon.History.IsOn() == true)
            AscCommon.History.TurnOff();
    }
    CPDFDoc.prototype.ShowComment = function(arrId)
    {
        let oPos;
        var arrCommentsId = [];

        for (var nIndex = 0, nCount = arrId.length; nIndex < nCount; ++nIndex)
        {
            var Comment = this.GetCommentById(arrId[nIndex]);
            if (Comment)
            {
                if (null == oPos)
                {
                    oPos = AscPDF.GetGlobalCoordsByPageCoords(Comment._pagePos.x + Comment._pagePos.w, Comment._pagePos.y + Comment._pagePos.h / 2, Comment.GetPage(), true);
                }

                arrCommentsId.push(Comment.GetId());
            }

        }

        if (null != oPos && arrCommentsId.length > 0)
        {
            editor.sync_ShowComment(arrCommentsId, oPos["X"], oPos["Y"]);
        }
        else
        {
            editor.sync_HideComment();
        }
    };
    CPDFDoc.prototype.UpdateCopyCutState = function() {
        editor.sync_CanCopyCutCallback(this.CanCopyCut());
    };
    CPDFDoc.prototype.CanCopyCut = function() {
        let oViewer = editor.getDocumentRenderer();

        let isHasSelect = false;

        let oSelection = oViewer.file.Selection;
        if (oSelection.Glyph1 != oSelection.Glyph2 || oSelection.Line1 != oSelection.Line2 ||
            oSelection.Page1 != oSelection.Page2) {
                isHasSelect = true;
            }
        

        if (this.activeForm && this.activeForm.content && this.activeForm.content.IsSelectionUse() && 
            this.activeForm.content.IsSelectionEmpty() == false) {
                isHasSelect = true;
            }

        return isHasSelect;
    };
    CPDFDoc.prototype.RemoveComment = function(Id) {
        let oViewer = editor.getDocumentRenderer();
        let oAnnot = this.annots.find(function(annot) {
            return annot.GetId() === Id;
        });

        if (!oAnnot)
            return;

        if (oAnnot.IsComment()) {
            let nPage = oAnnot.GetPage();
            oAnnot.AddToRedraw();
            this.annots.splice(this.annots.indexOf(oAnnot), 1);
            oViewer.pagesInfo.pages[nPage].annots.splice(oViewer.pagesInfo.pages[nPage].annots.indexOf(oAnnot), 1);
            editor.sync_RemoveComment(Id);
            oViewer._paint();
        }
        else {
            oAnnot.SetContents(null);
        }
    };
    CPDFDoc.prototype.RemoveAnnot = function(Id) {
        let oViewer = editor.getDocumentRenderer();
        let oAnnot = this.annots.find(function(annot) {
            return annot.GetId() === Id;
        });

        if (!oAnnot)
            return;

        if (oAnnot.IsComment())
            return this.RemoveComment(Id);

        let nPage = oAnnot.GetPage();
        oAnnot.AddToRedraw();

        let nPos        = this.annots.indexOf(oAnnot);
        let nPosInPage  = oViewer.pagesInfo.pages[nPage].annots.indexOf(oAnnot);

        this.annots.splice(nPos, 1);
        oViewer.pagesInfo.pages[nPage].annots.splice(nPosInPage, 1);
        
        if (this.mouseDownAnnot == oAnnot)
            this.mouseDownAnnot = null;

        if (oViewer.IsOpenAnnotsInProgress == false) {
            this.CreateNewHistoryPoint();
            this.History.Add(new CChangesPDFDocumentRemoveItem(this, [nPos, nPosInPage], [oAnnot]));
            this.TurnOffHistory();
        }

        oViewer._paint();
        oViewer.onUpdateOverlay();
    };
    CPDFDoc.prototype.HideComments = function() {
        editor.sync_HideComment();
    };
    CPDFDoc.prototype.GetCommentById = function(sId) {
        return this.annots.find(function(annot) {
            return annot.GetId() == sId;
        });
    };
    
    /**
	 * Changes the interactive field name.
	 * @memberof CPDFDoc
	 * @typeofeditors ["PDF"]
     * @param {AscPDF.CBaseField} oField - source field.
     * @param {String} cName - the new field name.
	 * @returns {AscPDF.CBaseField}
	 */
    CPDFDoc.prototype.private_changeFieldName = function(oField, cName) {
        while (cName.indexOf('..') != -1)
            cName = cName.replace(new RegExp("\.\.", "g"), ".");

        let oExistsWidget = this.GetField(cName);
        // если есть виджет-поле с таким именем то не добавляем 
        if (oExistsWidget && oExistsWidget.GetType() != oField.GetType())
            return null; // to do выдавать ошибку создания поля

        // получаем partial names
        let aPartNames = cName.split('.').filter(function(item) {
            if (item != "")
                return item;
        })

        // по формату не больше 20 вложенностей
        if (aPartNames.length > 20)
            return null;

        if (!oField._parent)
            return false;

        let oFieldParent = oField._parent;
        // удаляем поле из родителя
        oFieldParent.RemoveKid(oField);

        // создаем родительские поля, последнее будет виджет-полем
        if (aPartNames.length > 1) {
            if (this.rootFields.get(aPartNames[0]) == null) { // root поле
                this.rootFields.set(aPartNames[0], private_createField(aPartNames[0], oField.GetType(), oField.GetPage(), []));
            }

            let oParentField = this.rootFields.get(aPartNames[0]);
            
            for (let i = 1; i < aPartNames.length; i++) {
                // добавляем виджет-поле (то, которое рисуем)
                if (i == aPartNames.length - 1) {
                    oParentField.AddKid(oField);
                }
                else {
                    // если есть поле с таким именем (part name), то двигаемся дальше, если нет, то создаем
                    let oExistsField = oParentField.GetField(aPartNames[i]);
                    if (oExistsField)
                        oParentField = oExistsField;
                    else {
                        let oNewParent = private_createField(aPartNames[i], oField.GetType(), oField.GetPage(), []);
                        oParentField.AddKid(oNewParent);
                        oParentField = oNewParent;
                    }
                }
            }
        }

        this.private_checkField(oFieldParent);
        oField.SyncField();
        return oField;
    };
    CPDFDoc.prototype.DoTest = function() {
        let pdfDoc = this;
        let oViewer = editor.getDocumentRenderer();
	    	
        function CreateTextForm(name, aRect)
        {
            return pdfDoc.AddField(name, "text", 0, aRect);
        }
        function EnterTextToForm(form, text)
        {
            let chars = text.codePointsArray();
            pdfDoc.activeForm = form;
            form.EnterText(chars);
            pdfDoc.EnterDownActiveField();
        }
        function AddJsAction(form, trigger, script)
        {
            form.SetAction(trigger, script);
        }
	
        let textForm1 = CreateTextForm("TextForm1", [0, 0, 50, 50]);
		let textForm2 = CreateTextForm("TextForm2", [60, 0, 110, 50]);
		let textForm3 = CreateTextForm("TextForm3", [120, 0, 170, 50]);
		
		textForm1.GetFormApi().value = "1";
		textForm2.GetFormApi().value = "2";
		textForm3.GetFormApi().value = "3";
		
		AddJsAction(textForm1, AscPDF.FORMS_TRIGGERS_TYPES.Calculate, "this.getField('TextForm2').value += 1");
		AddJsAction(textForm2, AscPDF.FORMS_TRIGGERS_TYPES.Calculate, "this.getField('TextForm3').value += 1");
		AddJsAction(textForm3, AscPDF.FORMS_TRIGGERS_TYPES.Calculate, "this.getField('TextForm1').value += 1");
		
        textForm2.MoveCursorRight();
		EnterTextToForm(textForm2, "2");
		console.log(textForm1.GetValue(), "2", "Check form1 value");
		console.log(textForm2.GetValue(), "22", "Check form2 value");
		console.log(textForm3.GetValue(), "4", "Check form3 value");

        textForm3.MoveCursorRight();
		EnterTextToForm(textForm3, "3");
		
		console.log(textForm1.GetValue(), "3", "Check form1 value");
		console.log(textForm2.GetValue(), "23", "Check form2 value");
		console.log(textForm3.GetValue(), "43", "Check form3 value");
    }

    /**
	 * Changes the interactive field name.
     * Note: This method used by forms actions.
	 * @memberof CPDFDoc
     * @param {CBaseField[]} aNames - array with forms names to reset. If param is undefined or array is empty then resets all forms.
	 * @typeofeditors ["PDF"]
	 */
    CPDFDoc.prototype.ResetForms = function(aNames) {
        let oActionsQueue = this.GetActionsQueue();
        let oThis = this;

        if (aNames.length > 0) {
            aNames.forEach(function(name) {
                let aFields = oThis.GetFields(name);
                if (aFields.length > 0)
                    AscCommon.History.Clear()

                aFields.forEach(function(field) {
                    field.Reset();
                });
            });
        }
        else {
            this.widgets.forEach(function(field) {
                field.Reset();
            });
            if (this.widgets.length > 0)
                AscCommon.History.Clear()
        }

        oActionsQueue.Continue();
    };
    /**
	 * Hides/shows forms by names
	 * @memberof CPDFDoc
     * @param {boolean} bHidden
     * @param {AscPDF.CBaseField[]} aNames - array with forms names to reset. If param is undefined or array is empty then resets all forms.
	 * @typeofeditors ["PDF"]
	 * @returns {AscPDF.CBaseField}
	 */
    CPDFDoc.prototype.HideShowForms = function(bHidden, aNames) {
        let oActionsQueue = this.GetActionsQueue();
        let oThis = this;

        if (aNames.length > 0) {
            aNames.forEach(function(name) {
                let aFields = oThis.GetFields(name);
                aFields.forEach(function(field) {
                    if (bHidden)
                        field.SetDisplay(window["AscPDF"].Api.Objects.display["hidden"]);
                    else
                        field.SetDisplay(window["AscPDF"].Api.Objects.display["visible"]);
                    
                    field.AddToRedraw();
                });
            });
        }
        else {
            this.widgets.forEach(function(field) {
                if (bHidden)
                    field.SetDisplay(window["AscPDF"].Api.Objects.display["hidden"]);
                else
                    field.SetDisplay(window["AscPDF"].Api.Objects.display["visible"]);

                field.AddToRedraw();
            });
        }

        oActionsQueue.Continue();
    };

    /**
	 * Hides/shows annots by names
	 * @memberof CPDFDoc
     * @param {boolean} bHidden
	 * @typeofeditors ["PDF"]
	 * @returns {AscPDF.CAnnotationBase}
	 */
    CPDFDoc.prototype.HideShowAnnots = function(bHidden) {
        this.annots.forEach(function(annot) {
            annot.SetDisplay(bHidden ? window["AscPDF"].Api.Objects.display["hidden"] : window["AscPDF"].Api.Objects.display["visible"]);
            annot.AddToRedraw();
        });

        this.annotsHidden   = bHidden;

        this.HideComments();
        this.mouseDownAnnot = null;
        this.Viewer.DrawingObjects.resetSelection();
    };
    CPDFDoc.prototype.IsAnnotsHidden = function() {
        return this.annotsHidden;
    };
    /**
	 * Checks the field for the field widget, if not then the field will be removed.
	 * @memberof CPDFDoc
	 * @typeofeditors ["PDF"]
	 */
    CPDFDoc.prototype.private_checkField = function(oField) {
        if (oField._kids.length == 0) {
            if (oField._parent) {
                oField._parent.RemoveKid(oField);
                this.private_checkField(oField._parent);
            }
            else if (this.rootFields.get(oField.name)) {
                this.rootFields.delete(oField.name);
            }
        }
    };

    /**
	 * Returns array with widjets fields by specified name.
	 * @memberof CPDFDoc
	 * @typeofeditors ["PDF"]
	 * @returns {boolean}
	 */
    CPDFDoc.prototype.GetFields = function(sName) {
        let aFields = [];
        for (let i = 0; i < this.widgets.length; i++) {
            if (this.widgets[i].GetFullName() == sName)
                aFields.push(this.widgets[i]);
        }

        return aFields;
    };

    /**
	 * Gets API PDF doc.
	 * @memberof CPDFDoc
	 * @typeofeditors ["PDF"]
	 * @returns {boolean}
	 */
    CPDFDoc.prototype.GetDocumentApi = function() {
        if (this.api)
            return this.api;

        return new AscPDF.ApiDocument(this);
    };

    /**
	 * Gets field by name
	 * @memberof CPDFDoc
	 * @typeofeditors ["PDF"]
	 * @returns {?CBaseField}
	 */
    CPDFDoc.prototype.GetField = function(sName) {
        let aPartNames = sName.split('.').filter(function(item) {
            if (item != "")
                return item;
        })

        let sPartName = aPartNames[0];
        for (let i = 0; i < aPartNames.length; i++) {
            for (let j = 0; j < this.widgets.length; j++) {
                if (this.widgets[j].GetFullName() == sPartName) // checks by fully name
                    return this.widgets[j];
            }
            sPartName += "." + aPartNames[i + 1];
        }

        return null;
    };
	
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Extension required for CTextBoxContent
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CPDFDoc.prototype.IsTrackRevisions = function() {
		return false;
	};
	CPDFDoc.prototype.IsDocumentEditor = function() {
		return false;
	};
	CPDFDoc.prototype.IsPresentationEditor = function() {
		return false;
	};
	CPDFDoc.prototype.IsSpreadSheetEditor = function() {
		return false;
	};
	CPDFDoc.prototype.IsPdfEditor = function() {
		return true;
	};
	CPDFDoc.prototype.Get_Styles = function() {
		return AscCommonWord.DEFAULT_STYLES;
	};
	CPDFDoc.prototype.GetStyles = function() {
		return this.Get_Styles();
	};
    /**
     * Запрашиваем настройку автозамены двух дефисов на тире
     * @returns {boolean}
     */
    CPDFDoc.prototype.IsAutoCorrectHyphensWithDash = function()
    {
        return this.AutoCorrectSettings.IsHyphensWithDash();
    };
    CPDFDoc.prototype.GetHistory = function() {
        return AscCommon.History;
    };
	CPDFDoc.prototype.Get_Numbering = function() {
		return AscWord.DEFAULT_NUMBERING;
	};
	CPDFDoc.prototype.GetNumbering = function() {
		return this.Get_Numbering();
	};
	CPDFDoc.prototype.IsDoNotExpandShiftReturn = function() {
		return false;
	};
	CPDFDoc.prototype.GetCompatibilityMode = function() {
		return AscCommon.document_compatibility_mode_Word12;
	};
	CPDFDoc.prototype.Get_PageLimits = function(pageIndex) {
		let documentRenderer = this.GetDocumentRenderer();
		return documentRenderer.Get_PageLimits(pageIndex);
	};
	CPDFDoc.prototype.Get_PageFields = function(pageIndex) {
		return this.Get_PageLimits(pageIndex);
	};
	CPDFDoc.prototype.GetApi = function() {
		return editor;
	};
	CPDFDoc.prototype.CanEdit = function() {
		return true;
	};
	CPDFDoc.prototype.IsFillingFormMode = function() {
		return false;
	};
	CPDFDoc.prototype.getDrawingObjects = function() {
		if (!this.Viewer)
			return null;
		
		return this.Viewer.DrawingObjects;
	};
	CPDFDoc.prototype.checkDefaultFieldFonts = function(callback) {
		
		if (1 === this.defaultFontsLoaded)
			return true;
		
		if (callback)
			this.fontLoaderCallbacks.push(callback);

		if (0 === this.defaultFontsLoaded)
			return false;
		
		this.defaultFontsLoaded = 0;
		let _t = this;
		this.fontLoader.LoadDocumentFonts2([{name : AscPDF.DEFAULT_FIELD_FONT}],
			Asc.c_oAscAsyncActionType.Empty,
			function()
			{
				_t.defaultFontsLoaded = 1;
				_t.fontLoaderCallbacks.forEach(function(callback) {
					callback();
				});
				
				_t.fontLoaderCallbacks = [];
			}
		);

		return 1 === this.defaultFontsLoaded;
	};
	
    function CActionQueue(oDoc) {
        this.doc            = oDoc;
        this.actions        = [];
        this.isInProgress   = false;
        this.curAction  = null;
        this.curActionIdx   = -1;
        this.callBackAfterFocus = null;
    };

    CActionQueue.prototype.AddActions = function(aActions) {
        this.actions = this.actions.concat(aActions);
    };
    CActionQueue.prototype.SetCurAction = function(oAction) {
        this.curAction = oAction;
    };
    CActionQueue.prototype.GetNextAction = function() {
        return this.actions[this.curActionIdx + 1];
    };
    CActionQueue.prototype.Clear = function() {
        this.actions = [];
        this.curActionIdx = -1;
        this.curAction = null;
        this.callBackAfterFocus = null;
    };
    CActionQueue.prototype.Stop = function() {
        this.SetInProgress(false);
    };
    CActionQueue.prototype.IsInProgress = function() {
        return this.isInProgress;
    };
    CActionQueue.prototype.SetInProgress = function(bValue) {
        this.isInProgress = bValue;
    };
    CActionQueue.prototype.SetCurActionIdx = function(nValue) {
        this.curActionIdx = nValue;
    };
    CActionQueue.prototype.Start = function() {
        if (this.IsInProgress() == false) {
            let oFirstAction = this.actions[0];
            if (oFirstAction) {
                this.SetInProgress(true);
                this.SetCurActionIdx(0);
                setTimeout(function() {
                    oFirstAction.Do();
                }, 100);
            }
        }
    };
    CActionQueue.prototype.Continue = function() {
        let oNextAction = this.GetNextAction();
        if (this.callBackAfterFocus && this.curAction.triggerType == AscPDF.FORMS_TRIGGERS_TYPES.OnFocus && (!oNextAction || oNextAction.triggerType != AscPDF.FORMS_TRIGGERS_TYPES.OnFocus))
            this.callBackAfterFocus();

        if (oNextAction && this.IsInProgress()) {
            this.curActionIdx += 1;
            oNextAction.Do();
        }
        else {
            this.Stop();
            this.doc.OnEndFormsActions();
            this.Clear();
        }
    };

    function private_createField(cName, cFieldType, nPageNum, oCoords, oPdfDoc) {
        let oField;
        switch (cFieldType) {
            case AscPDF.FIELD_TYPES.button:
                oField = new AscPDF.CPushButtonField(cName, nPageNum, oCoords, oPdfDoc);
                break;
            case AscPDF.FIELD_TYPES.checkbox:
                oField = new AscPDF.CCheckBoxField(cName, nPageNum, oCoords, oPdfDoc);
                break;
            case AscPDF.FIELD_TYPES.combobox:
                oField = new AscPDF.CComboBoxField(cName, nPageNum, oCoords, oPdfDoc);
                break;
            case AscPDF.FIELD_TYPES.listbox:
                oField = new AscPDF.CListBoxField(cName, nPageNum, oCoords, oPdfDoc);
                break;
            case AscPDF.FIELD_TYPES.radiobutton:
                oField = new AscPDF.CRadioButtonField(cName, nPageNum, oCoords, oPdfDoc);
                break;
            case AscPDF.FIELD_TYPES.signature:
                oField = new AscPDF.CSignatureField(cName, nPageNum, oCoords, oPdfDoc);;
                break;
            case AscPDF.FIELD_TYPES.text:
                oField = new AscPDF.CTextField(cName, nPageNum, oCoords, oPdfDoc);
                break;
            case AscPDF.FIELD_TYPES.unknown: 
                oField = new AscPDF.CBaseField(cName, nPageNum, oCoords, oPdfDoc);
                break;
        }

        return oField;
    }

    function CreateAnnotByProps(oProps, oPdfDoc) {
        let aRect       = oProps.rect;
        let nPageNum    = oProps.page;
        let sName       = oProps.name ? oProps.name : AscCommon.CreateGUID();
        let nAnnotType  = oProps.type;
        let sAuthor     = oProps.author ? oProps.author : AscCommon.UserInfoParser.getCurrentName();
        let sDate       = oProps.modDate ? oProps.modDate : (new Date().getTime()).toString();
        let sText       = oProps.contents;
        let isHidden    = !!oProps.hidden;
        
        let oAnnot;

        let oViewer = editor.getDocumentRenderer();
        let nScaleY = oViewer.drawingPages[nPageNum].H / oViewer.file.pages[nPageNum].H / oViewer.zoom;
        let nScaleX = oViewer.drawingPages[nPageNum].W / oViewer.file.pages[nPageNum].W / oViewer.zoom;

        let aScaledCoords = [aRect[0] * nScaleX, aRect[1] * nScaleY, aRect[2] * nScaleX, aRect[3] * nScaleY];
        switch (nAnnotType) {
            case AscPDF.ANNOTATIONS_TYPES.Text:
                oAnnot = new AscPDF.CAnnotationText(sName, nPageNum, aScaledCoords, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Ink:
                oAnnot = new AscPDF.CAnnotationInk(sName, nPageNum, aScaledCoords, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Highlight:
                oAnnot = new AscPDF.CAnnotationHighlight(sName, nPageNum, aScaledCoords, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Underline:
                oAnnot = new AscPDF.CAnnotationUnderline(sName, nPageNum, aScaledCoords, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Strikeout:
                oAnnot = new AscPDF.CAnnotationStrikeout(sName, nPageNum, aScaledCoords, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Squiggly:
                oAnnot = new AscPDF.CAnnotationSquiggly(sName, nPageNum, aScaledCoords, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Line:
                oAnnot = new AscPDF.CAnnotationLine(sName, nPageNum, aScaledCoords, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Square:
                oAnnot = new AscPDF.CAnnotationSquare(sName, nPageNum, aScaledCoords, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Circle:
                oAnnot = new AscPDF.CAnnotationCircle(sName, nPageNum, aScaledCoords, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Polygon:
                oAnnot = new AscPDF.CAnnotationPolygon(sName, nPageNum, aScaledCoords, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.PolyLine:
                oAnnot = new AscPDF.CAnnotationPolyLine(sName, nPageNum, aScaledCoords, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.FreeText:
                oAnnot = new AscPDF.CAnnotationFreeText(sName, nPageNum, aScaledCoords, oPdfDoc);
                break;
            default:
                oAnnot = new AscPDF.CAnnotationBase(sName, nAnnotType, nPageNum, aScaledCoords, oPdfDoc);
                break;
        }

        oAnnot.SetModDate(sDate);
        oAnnot.SetAuthor(sAuthor);
        oAnnot.SetDisplay(isHidden ? window["AscPDF"].Api.Objects.display["hidden"] : window["AscPDF"].Api.Objects.display["visible"]);
        oAnnot.SetContents(sText);

        oAnnot._origRect = aRect;
        oAnnot._pagePos = {
            x: aScaledCoords[0],
            y: aScaledCoords[1],
            w: (aScaledCoords[2] - aScaledCoords[0]),
            h: (aScaledCoords[3] - aScaledCoords[1])
        };

        return oAnnot;
    }

    function private_PtToMM(pt)
	{
		return 25.4 / 72.0 * pt;
	}

    function getMinRect(aPoints) {
        let xMax = aPoints[0], yMax = aPoints[1], xMin = xMax, yMin = yMax;
        for(let i = 1; i < aPoints.length; i++) {
            if (i % 2 == 0) {
                if(aPoints[i] < xMin)
                {
                    xMin = aPoints[i];
                }
                if(aPoints[i] > xMax)
                {
                    xMax = aPoints[i];
                }
            }
            else {
                if(aPoints[i] < yMin)
                {
                    yMin = aPoints[i];
                }

                if(aPoints[i] > yMax)
                {
                    yMax = aPoints[i];
                }
            }
        }

        return [xMin, yMin, xMax, yMax];
    }

    if (!window["AscPDF"])
	    window["AscPDF"] = {};

    window["AscPDF"].CPDFDoc = CPDFDoc;
    window["AscPDF"].CreateAnnotByProps = CreateAnnotByProps;

})();
