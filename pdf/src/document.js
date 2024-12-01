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

    /** @enum {number} */
    let AscLockTypeElemPDF = {
        Object:     1,
        Page:       2,
        Document:   3
    };

    function CCalculateInfo(oDoc) {
        this.ids = [];
        this.document = oDoc;
        this.isInProgress = false;
        this.sourceField = null; // поле вызвавшее calculate
    };

    CCalculateInfo.prototype.AddFieldToOrder = function(id) {
        if (this.ids.includes(id) == false)
            this.ids.push(id);
    };
    CCalculateInfo.prototype.RemoveFieldFromOrder = function(id) {
        let nIdx = this.ids.indexOf(id);
        if (nIdx != -1) {
            this.ids.splice(nIdx, 1);
        }
    };
    CCalculateInfo.prototype.SetIsInProgress = function(bValue) {
        this.isInProgress = bValue;
    };
    CCalculateInfo.prototype.IsInProgress = function() {
        return this.isInProgress;
    };
    CCalculateInfo.prototype.SetCalculateOrder = function(aIds) {
        this.ids = aIds.slice();
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
        this.drawings   = []; // из презентаций (чарты, шейпы, картинки)

        this.widgetsParents = []; // все родительские поля

        this.maxApIdx               = -1;
        this.CollaborativeEditing   = AscCommon.CollaborativeEditing;
        this.CollaborativeEditing.SetLogicDocument(this);
        this.NeedUpdateTargetForCollaboration = true;
        this.LastUpdateTargetTime = 0;

        this.MathTrackHandler       = new AscWord.CMathTrackHandler(this.GetDrawingDocument(), Asc.editor);
        this.AnnotTextPrTrackHandler= new AscPDF.CAnnotTextPrTrackHandler(this.GetDrawingDocument(), Asc.editor);
        this.TextSelectTrackHandler = new AscPDF.CTextSelectTrackHandler(this.GetDrawingDocument(), Asc.editor);
        this.SearchEngine           = new AscPDF.CPdfSearch(this);

        this.theme                  = AscFormat.GenerateDefaultTheme(this);
        this.clrSchemeMap           = AscFormat.GenerateDefaultColorMap();
        this.styles                 = AscWord.DEFAULT_STYLES.Copy();
        this.TableStylesIdMap       = {};
        this.InitDefaultTextListStyles();
        this.InitDefaultTableStyles();
        this.actionsInfo            = new CActionQueue(this);
        this.calculateInfo          = new CCalculateInfo(this);
        this.fieldsToCommit         = [];
        this.event                  = {};
        this.lastDatePickerInfo     = null;
        this.AutoCorrectSettings    = new AscCommon.CAutoCorrectSettings();

        this.pagesTransform = [];

        Object.defineProperties(this.event, {
            "change": {
                set: function(value) {
                    if (value != null && value.toString)
                        this._change = value.toString();
                },
                get: function() {
                    return this._change;
                }
            }
        });

        this._parentsMap = {}; // map при открытии форм
        this.api = this.GetDocumentApi();
		

        this.CurPosition = {X: 0, Y: 0}; // для graphic frame

        // internal
        this.activeForm         = null;
        this.activeDrawing    = null;
        this.mouseDownField     = null;
        this.mouseDownAnnot     = null;

        this._id = AscCommon.g_oIdCounter.Get_NewId();
		AscCommon.g_oTableId.Add(this, this._id);
        
		this.History = new AscPDF.History(this);
        this.History.Set_LogicDocument(this);
		this.pagesContentChanges = new AscCommon.CContentChanges();

        if (AscCommon.History)
        {
            this.History.UserSaveMode   = AscCommon.History.UserSaveMode;
            this.History.UserSavedIndex = AscCommon.History.UserSavedIndex;
        }

		this.LocalHistory   = new AscPDF.History(this);
		AscCommon.History = this.History;

		this.Spelling   = new AscWord.CDocumentSpellChecker();
        this.Viewer     = viewer;
        this.Api        = Asc.editor;

        this.annotsHidden = false;
		
		this.fontLoader             = AscCommon.g_font_loader;
		this.defaultFontsLoaded     = -1; // -1 не загружены и не грузим, 0 - грузим, 1 - загружены
        this.loadedFonts            = [];
        this.Action                 = {};
    }

	CPDFDoc.prototype.RecalculateAll = function() {
		let fontMap = {};
		
		this.widgets.forEach(function(field) {
			if (field.IsNeedDrawFromStream())
				return;
			
			let fontName = field.GetTextFontActual();
			if (fontName)
				fontMap[fontName] = true;
		});
		
		this.drawings.forEach(function(drawing) {
			drawing.GetAllFonts(fontMap);
		});
		
		this.annots.forEach(function(annot) {
			annot.GetAllFonts(fontMap);
		});
		
        if (this.checkFonts(Object.keys(fontMap), this._RecalculateAll.bind(this))) {
            this._RecalculateAll();
        }
	};
	CPDFDoc.prototype._RecalculateAll = function() {
		this.widgets.forEach(function(field) {
			if (field.IsNeedDrawFromStream())
				return;
			
			field.Recalculate();
		});
		
		this.drawings.forEach(function(drawing) {
			drawing.Recalculate();
		});
		
		this.annots.forEach(function(annot) {
			annot.Recalculate();
		});
	};
    CPDFDoc.prototype.UpdatePagesTransform = function() {
        this.pagesTransform = [];

        let oFile = this.Viewer.file;
        for (let i = 0; i < oFile.pages.length; i++) {
            let oPage = this.Viewer.drawingPages[i];
            if (!oPage) {
                continue;
            }

            let nAngle  = this.Viewer.getPageRotate(i);

            let oPageTr = new AscCommon.CMatrix();

            let xCenter = this.Viewer.width >> 1;
			if (this.Viewer.documentWidth > this.Viewer.width)
				xCenter = (this.Viewer.documentWidth >> 1) - (this.Viewer.scrollX) >> 0;

            let nPageW  = oPage.W;
            let nPageH  = oPage.H;
            let xInd    = xCenter - (oPage.W >> 1);
            let yInd    = -(this.Viewer.scrollY - this.Viewer.drawingPages[i].Y);
            
            let nScale = this.Viewer.file.pages[i].W / this.Viewer.drawingPages[i].W;

            let shx = 0, shy = 0, sx = 1, sy = 1, tx = 0, ty = 0;

            switch (nAngle) {
                case 0: {
                    tx = -xInd * nScale;
                    ty = -yInd * nScale;
                    sx = nScale;
                    sy = nScale;
                    shx = 0;
                    shy = 0;
                    break;
                }
                case 90: {
                    // Новый отступ слева после поворота
                    let newXInd = xInd + (nPageW - nPageH >> 1);
                    tx = -yInd * nScale - (0.5 / this.Viewer.zoom); // магическое число
                    ty = (nPageH + newXInd) * nScale - (0.5 / this.Viewer.zoom); // магическое число
                    sx = 0;
                    sy = 0;
                    shx = 1 * nScale;
                    shy = -1 * nScale;
                    break;
                }
                case 180: {
                    tx = (xInd + nPageW) * nScale - (1.5 / this.Viewer.zoom); // магическое число
                    ty = (yInd + nPageH) * nScale;
                    sx = -nScale;
                    sy = -nScale;
                    shx = 0;
                    shy = 0;
                    break;
                }
                case 270: {
                    // Новый отступ слева после поворота
                    let newXInd = xInd + (nPageW - nPageH >> 1);
                    tx = (nPageW + yInd) * nScale;
                    ty = -newXInd * nScale + (1.5 / this.Viewer.zoom); // магическое число;
                    sx = 0;
                    sy = 0;
                    shx = -1 * nScale;
                    shy = 1 * nScale;
                    break;
                }
            }
            
            oPageTr.shx = shx;
            oPageTr.shy = shy;
            oPageTr.sx  = sx;
            oPageTr.sy  = sy;
            oPageTr.tx  = tx;
            oPageTr.ty  = ty;
            
            this.pagesTransform.push({
                normal: oPageTr,
                invert: AscCommon.global_MatrixTransformer.Invert(oPageTr)
            });
        }
    };
    CPDFDoc.prototype.GetPageTransform = function(nPage, bForcedCalc) {
        if (!bForcedCalc) {
            return this.pagesTransform[nPage];
        }

        let oPage   = this.Viewer.drawingPages[nPage];
        let nAngle  = this.Viewer.getPageRotate(nPage);

        let oPageTr = new AscCommon.CMatrix();

        let xCenter = this.Viewer.width >> 1;
        if (this.Viewer.documentWidth > this.Viewer.width)
            xCenter = (this.Viewer.documentWidth >> 1) - (this.Viewer.scrollX) >> 0;

        let nPageW  = oPage.W;
        let nPageH  = oPage.H;
        let xInd    = xCenter - (oPage.W >> 1);
        let yInd    = -(this.Viewer.scrollY - this.Viewer.drawingPages[nPage].Y);
        
        let nScale = this.Viewer.file.pages[nPage].W / this.Viewer.drawingPages[nPage].W;

        let shx = 0, shy = 0, sx = 1, sy = 1, tx = 0, ty = 0;

        switch (nAngle) {
            case 0: {
                tx = -xInd * nScale;
                ty = -yInd * nScale;
                sx = nScale;
                sy = nScale;
                shx = 0;
                shy = 0;
                break;
            }
            case 90: {
                // Новый отступ слева после поворота
                let newXInd = xInd + (nPageW - nPageH >> 1);
                tx = -yInd * nScale - (0.5 / this.Viewer.zoom); // магическое число
                ty = (nPageH + newXInd) * nScale - (0.5 / this.Viewer.zoom); // магическое число
                sx = 0;
                sy = 0;
                shx = 1 * nScale;
                shy = -1 * nScale;
                break;
            }
            case 180: {
                tx = (xInd + nPageW) * nScale - (1.5 / this.Viewer.zoom); // магическое число
                ty = (yInd + nPageH) * nScale;
                sx = -nScale;
                sy = -nScale;
                shx = 0;
                shy = 0;
                break;
            }
            case 270: {
                // Новый отступ слева после поворота
                let newXInd = xInd + (nPageW - nPageH >> 1);
                tx = (nPageW + yInd) * nScale;
                ty = -newXInd * nScale + (1.5 / this.Viewer.zoom); // магическое число;
                sx = 0;
                sy = 0;
                shx = -1 * nScale;
                shy = 1 * nScale;
                break;
            }
        }
        
        oPageTr.shx = shx;
        oPageTr.shy = shy;
        oPageTr.sx  = sx;
        oPageTr.sy  = sy;
        oPageTr.tx  = tx;
        oPageTr.ty  = ty;
        
        return {
            normal: oPageTr,
            invert: AscCommon.global_MatrixTransformer.Invert(oPageTr)
        }
    };

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
            oViewer.paint();
            oViewer.needRedraw = false;
        }
        else {
            oViewer.paint();
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
            if (aParentsInfo[i]["i"] != null)
                oParent.SetApIdx(aParentsInfo[i]["i"]);
            if (aParentsInfo[i]["curIdxs"])
                oParent.SetApiCurIdxs(aParentsInfo[i]["curIdxs"]);
            if (aParentsInfo[i]["Opt"] && oParent instanceof AscPDF.CBaseCheckBoxField)
                oParent.SetOptions(aParentsInfo[i]["Opt"]);

            oParents[nIdx] = oParent;
            this.rootFields.set(oParent.GetPartialName(), oParent);
            this.widgetsParents.push(oParent);
        }

        for (let nParentIdx in oParents) {
            oChilds[nParentIdx].forEach(function(child) {
                oParents[nParentIdx].AddKid(child);
            });
        }
    };
    CPDFDoc.prototype.SetLocalHistory = function() {
        AscCommon.History = this.LocalHistory;
    };
    CPDFDoc.prototype.SetGlobalHistory = function() {
        this.LocalHistory.Clear();
        AscCommon.History = this.History;
    };
    CPDFDoc.prototype.OnAfterFillFormsParents = function() {
        let bInberitValue = false;
        let value;

        let aRadios = []; // обновляем состояние радиокнопок в конце

        for (let i = 0; i < this.widgets.length; i++) {
            let oField = this.widgets[i];
            if ((oField.GetPartialName() == null || oField.GetApiValue(bInberitValue) == null) && oField.GetParent()) {
                let oParent = oField.GetParent();
                if (oParent.GetType() == AscPDF.FIELD_TYPES.radiobutton && oParent.IsAllKidsWidgets())
                    aRadios.push(oParent);

                value = oParent.GetApiValue(false);
                if (value != null && value.toString) {
                    value = value.toString();
                }

                if (oParent._currentValueIndices && oParent._currentValueIndices.length != 0) {
                    oField.SetCurIdxs(oParent._currentValueIndices);
                }
                else {
                    if (oField.GetType() !== AscPDF.FIELD_TYPES.radiobutton)
                        oField.SetValue(value, true);
                }
            }
        }

        aRadios.forEach(function(field) {
            field.GetKid(0).UpdateAll();
        });
    };
    CPDFDoc.prototype.FillButtonsIconsOnOpen = function() {
		let oViewer = editor.getDocumentRenderer();
		if (window["NATIVE_EDITOR_ENJINE"]) {
			oViewer.IsOpenFormsInProgress = false;
			return;
		}
		
        let oDoc = this;

        let aIconsToLoad = [];
        let oIconsInfo = {
            "MK": [],
            "View": []
        };

        for (let i = 0; i < oViewer.pagesInfo.pages.length; i++) {
            let oPage = oViewer.drawingPages[i];

            let w = (oPage.W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
            let h = (oPage.H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

            let oFile = oViewer.file;
            let oPageIconsInfo = oFile.nativeFile["getButtonIcons"](i, w, h, undefined, true);

            if (oPageIconsInfo["View"] == null)
                continue;

            oIconsInfo["MK"] = oIconsInfo["MK"].concat(oPageIconsInfo["MK"]);
            oIconsInfo["View"] = oIconsInfo["View"].concat(oPageIconsInfo["View"]);

            // load images
            for (let nIcon = 0; nIcon < oPageIconsInfo["View"].length; nIcon++) {
                let sBase64 = oPageIconsInfo["View"][nIcon]["retValue"];

                aIconsToLoad.push({
                    Image: {
                        width: oPageIconsInfo["View"][nIcon]["w"],
                        height: oPageIconsInfo["View"][nIcon]["h"],
                    },
                    src: "data:image/png;base64," + sBase64
                });

                for (let nField = 0; nField < oPageIconsInfo["MK"].length; nField++) {
                    if (oPageIconsInfo["MK"][nField]["I"] == oPageIconsInfo["View"][nIcon]["j"]) {
                        oPageIconsInfo["MK"][nField]["I"] = aIconsToLoad[aIconsToLoad.length - 1];
                    }
                    else if (oPageIconsInfo["MK"][nField]["RI"] == oPageIconsInfo["View"][nIcon]["j"]) {
                        oPageIconsInfo["MK"][nField]["RI"] = aIconsToLoad[aIconsToLoad.length - 1];
                    }
                    else if (oPageIconsInfo["MK"][nField]["IX"] == oPageIconsInfo["View"][nIcon]["j"]) {
                        oPageIconsInfo["MK"][nField]["IX"] = aIconsToLoad[aIconsToLoad.length - 1];
                    }
                }
            }
        }

        if (aIconsToLoad.length === 0) {
            oViewer.IsOpenFormsInProgress = false;
            return;
        }

        editor.ImageLoader.LoadImagesWithCallback(aIconsToLoad.map(function(info) {
            return info.src;
        }), function() {
            // выставляем только ImageData. Форму пересчитаем и добавим картинку после того, как форма изменится, чтобы не грузить шрифты
            for (let nBtn = 0; nBtn < oIconsInfo["MK"].length; nBtn++) {
                let oBtnField = oDoc.GetFieldBySourceIdx(oIconsInfo["MK"][nBtn]["i"]);

                if (oIconsInfo["MK"][nBtn]["I"]) {
                    oBtnField.SetImageRasterId(oIconsInfo["MK"][nBtn]["I"].src, AscPDF.APPEARANCE_TYPE.normal);
                }
                if (oIconsInfo["MK"][nBtn]["RI"]) {
                    oBtnField.SetImageRasterId(oIconsInfo["MK"][nBtn]["RI"].src, AscPDF.APPEARANCE_TYPE.rollover);
                }
                if (oIconsInfo["MK"][nBtn]["IX"]) {
                    oBtnField.SetImageRasterId(oIconsInfo["MK"][nBtn]["IX"].src, AscPDF.APPEARANCE_TYPE.mouseDown);
                }
            }
            oViewer.isRepaint = true;
            oViewer.IsOpenFormsInProgress = false;
        });
    };
    
    ////////////////////////////////////

    CPDFDoc.prototype.Search = function(oProps) {
        if (true === this.SearchEngine.Compare(oProps))
		    return this.SearchEngine;
        
        this.SearchEngine.Clear();
        this.SearchEngine.Set(oProps);
        this.SearchEngine.Search();

        return this.SearchEngine;
    };
    CPDFDoc.prototype.SelectSearchElement = function(id) {
        this.BlurActiveObject();
        this.SearchEngine.Select(id);
    };
    CPDFDoc.prototype.GetSearchElementId = function(isNext) {
        let nCurPage        = this.Viewer.currentPage;
        let nCurMatchIdx    = this.SearchEngine.CurId;

        if (this.SearchEngine.Count == 0) {
            return -1;
        }
        
        if (nCurMatchIdx == -1) {
            nCurMatchIdx = 0;

            // находим индекс найденного элемента на текущей странице
            nCurMatchIdx += this.SearchEngine.PagesMatches.slice(0, nCurPage).reduce(function(accum, pageMatches) {
                return accum + pageMatches.length;
            }, 0);
        }
        else {
            if (isNext) {
                nCurMatchIdx = nCurMatchIdx + 1 < this.SearchEngine.Count ? nCurMatchIdx + 1 : 0;
            }
            else {
                nCurMatchIdx = nCurMatchIdx - 1 >= 0 ? nCurMatchIdx - 1 : this.SearchEngine.Count - 1;
            }
        }

        return nCurMatchIdx;
    };
    CPDFDoc.prototype.ClearSearch = function() {
        let isPrevSearch = this.SearchEngine.Count > 0;

        this.SearchEngine.Clear();

        if (isPrevSearch) {
            this.Api.sync_SearchEndCallback();
        }
    };

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
        // this.skipHistoryOnCommit = true;
        this.fieldsToCommit.forEach(function(field) {
            field.Commit();
        });
        
        this.ClearFieldsToCommit();
        // this.skipHistoryOnCommit = false;
    };
    CPDFDoc.prototype.ClearCacheForms = function(nPageIndex) {
        let oViewer = editor.getDocumentRenderer();

        if (oViewer.pagesInfo.pages[nPageIndex].fields != null) {
            oViewer.pagesInfo.pages[nPageIndex].fields.forEach(function(field) {
                field.ClearCache();
            });
        }

        oViewer.file.pages[nPageIndex].fieldsAPInfo = null;
    };
    CPDFDoc.prototype.ClearCacheAnnots = function(nPageIndex) {
        let oViewer = editor.getDocumentRenderer();

        if (oViewer.pagesInfo.pages[nPageIndex].annots != null) {
            oViewer.pagesInfo.pages[nPageIndex].annots.forEach(function(annot) {
                annot.ClearCache();
            });
        }
        
        oViewer.file.pages[nPageIndex].annotsAPInfo = null;
    };

    CPDFDoc.prototype.ClearCache = function(nPageIndex) {
        this.ClearCacheForms(nPageIndex);
        this.ClearCacheAnnots(nPageIndex);
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
    CPDFDoc.prototype.SelectNextForm = function() {
        let oViewer         = editor.getDocumentRenderer();
        let oDrDoc          = this.GetDrawingDocument();
        let aWidgetForms    = this.widgets;
        let oActionsQueue   = this.GetActionsQueue();
		
		if (aWidgetForms.length == 0)
            return;

        let nCurIdx = this.widgets.indexOf(this.activeForm);
        let oCurForm = this.widgets[nCurIdx];
        let oNextForm;

        for (let i = nCurIdx + 1; i <= this.widgets.length; i++) {
            if (this.widgets[i]) {
                if (this.widgets[i].IsHidden() == false) {
                    oNextForm = this.widgets[i];
                    break;
                }
            }
            else {
                if (this.widgets[0] != oCurForm)
                    oNextForm = this.widgets[0];
                else
                    return;
            }
        }

        if (!oNextForm)
            return;

        let _t = this;
		if (!this.checkFieldFont(oNextForm, function(){_t.SelectNextForm();}))
			return;

        this.BlurActiveObject();
        this.activeForm = oNextForm;
        oNextForm.Recalculate();
        oNextForm.SetDrawHighlight(false);
        
        if (oNextForm.IsNeedDrawFromStream() == true && oNextForm.GetType() != AscPDF.FIELD_TYPES.button) {
            oNextForm.SetDrawFromStream(false);
        }
        
        oNextForm.onFocus();
        if (oNextForm.GetType() != AscPDF.FIELD_TYPES.button) {
            oNextForm.AddToRedraw();
        }

        let callbackAfterFocus = function() {
            oNextForm.SetInForm(true);

            switch (oNextForm.GetType()) {
                case AscPDF.FIELD_TYPES.text:
                case AscPDF.FIELD_TYPES.combobox:
                    this.SetLocalHistory();

                    oDrDoc.UpdateTargetFromPaint = true;
                    oDrDoc.m_lCurrentPage = 0;
                    oDrDoc.m_lPagesCount = oViewer.file.pages.length;
                    oDrDoc.TargetStart(true);
                    if (oNextForm.content.IsSelectionUse())
                        oNextForm.content.RemoveSelection();
    
                    oNextForm.content.MoveCursorToStartPos();
                    oNextForm.content.RecalculateCurPos();
                    
                    break;
                default:
                    oDrDoc.TargetEnd();
                    break;
            }
        };
        
        this.NavigateToField(oNextForm);
                
        let oOnFocus = oNextForm.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.OnFocus);
        // вызываем выставление курсора после onFocus. Если уже в фокусе, тогда сразу.
        if (oOnFocus && oOnFocus.Actions.length > 0) {
            oActionsQueue.callbackAfterFocus = callbackAfterFocus.bind(this);
            oActionsQueue.AddActions(oOnFocus.Actions);
            oActionsQueue.Start();
        }
        else
            callbackAfterFocus.bind(this)();
    };
    CPDFDoc.prototype.SelectPrevForm = function() {
        let oViewer         = editor.getDocumentRenderer();
        let oDrDoc          = this.GetDrawingDocument();
        let aWidgetForms    = this.widgets;
        let oActionsQueue   = this.GetActionsQueue();
		
		if (aWidgetForms.length == 0)
            return;

        let nCurIdx = this.widgets.indexOf(this.activeForm);
        let oCurForm = this.widgets[nCurIdx];
        let oNextForm;

        for (let i = nCurIdx - 1; i >= -1; i--) {
            if (this.widgets[i]) {
                if (this.widgets[i].IsHidden() == false) {
                    oNextForm = this.widgets[i];
                    break;
                }
            }
            else {
                if (this.widgets[this.widgets.length - 1] != oCurForm)
                    oNextForm = this.widgets[this.widgets.length - 1];
                else
                    return;
            }
        }

        if (!oNextForm)
            return;

        let _t = this;
		if (!this.checkFieldFont(oNextForm, function(){_t.SelectNextForm();}))
			return;

        this.BlurActiveObject();
        this.activeForm = oNextForm;
        oNextForm.Recalculate();
        oNextForm.SetDrawHighlight(false);
        
        if (oNextForm.IsNeedDrawFromStream() == true && oNextForm.GetType() != AscPDF.FIELD_TYPES.button) {
            oNextForm.SetDrawFromStream(false);
        }
        
        oNextForm.onFocus();
        if (oNextForm.GetType() != AscPDF.FIELD_TYPES.button) {
            oNextForm.AddToRedraw();
        }

        let callbackAfterFocus = function() {
            oNextForm.SetInForm(true);

            switch (oNextForm.GetType()) {
                case AscPDF.FIELD_TYPES.text:
                case AscPDF.FIELD_TYPES.combobox:
                    this.SetLocalHistory();

                    oDrDoc.UpdateTargetFromPaint = true;
                    oDrDoc.m_lCurrentPage = 0;
                    oDrDoc.m_lPagesCount = oViewer.file.pages.length;
                    oDrDoc.TargetStart(true);
                    if (oNextForm.content.IsSelectionUse())
                        oNextForm.content.RemoveSelection();
    
                    oNextForm.content.MoveCursorToStartPos();
                    oNextForm.content.RecalculateCurPos();
                    
                    break;
                default:
                    oDrDoc.TargetEnd();
                    break;
            }
        };

        this.NavigateToField(oNextForm);
        
        let oOnFocus = oNextForm.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.OnFocus);
        // вызываем выставление курсора после onFocus. Если уже в фокусе, тогда сразу.
        if (oOnFocus && oOnFocus.Actions.length > 0) {
            oActionsQueue.callbackAfterFocus = callbackAfterFocus.bind(this);
            oActionsQueue.AddActions(oOnFocus.Actions);
            oActionsQueue.Start();
        }
        else
            callbackAfterFocus.bind(this)();
    };
    CPDFDoc.prototype.NavigateToField = function(oField) {
        let oViewer     = Asc.editor.getDocumentRenderer();
        let aOrigRect   = oField.GetOrigRect();
        let nPage       = oField.GetPage();
        
        let oViewRect = oViewer.getViewingRect2(nPage);

        let nOrigPageW = oViewer.file.pages[nPage].W;
        let nOrigPageH = oViewer.file.pages[nPage].H;
        let nPageRot = oViewer.getPageRotate(nPage);

        let nStartViewX = nOrigPageW * oViewRect.x;
        let nStartViewY = nOrigPageH * oViewRect.y;
        let nEndViewX   = nOrigPageW * oViewRect.r;
        let nEndViewY   = nOrigPageH * oViewRect.b;

        if ((aOrigRect[3] > nStartViewY && aOrigRect[1] < nEndViewY) && (aOrigRect[2] > nStartViewX && aOrigRect[0] < nEndViewX)) {
            return;
        }

        // выставляем смещения
        let yOffset;
        let xOffset;

        switch (nPageRot) {
            case 0:
                yOffset = aOrigRect[1];
                xOffset = aOrigRect[0];
                break;
            case 90:
                yOffset = aOrigRect[3];
                xOffset = aOrigRect[0];
                break;
            case 180:
                yOffset = aOrigRect[3];
                xOffset = aOrigRect[2];
                break;
            case 270:
                yOffset = aOrigRect[1];
                xOffset = aOrigRect[2];
                break;
        }

        let oTr = this.pagesTransform[nPage].invert;
        let oPos = oTr.TransformPoint(xOffset, yOffset);

        this.Viewer.scrollToXY(this.Viewer.scrollY + oPos.y, this.Viewer.scrollX + oPos.x);
    };
    CPDFDoc.prototype.CommitField = function(oField) {
        return this.DoAction(function() {
            let isValid = true;

            if ([AscPDF.FIELD_TYPES.text, AscPDF.FIELD_TYPES.combobox].includes(oField.GetType())) {
                isValid = oField.DoValidateAction(oField.GetValue(true));
            }
    
            if (isValid) {
                oField.needValidate = false; 
                oField.Commit();
                if (this.event["rc"] == true && this.IsNeedDoCalculate()) {
                    this.DoCalculateFields(oField);
                    this.AddFieldToCommit(oField);
                    this.CommitFields();
                }
    
                isValid = this.event["rc"];
            }
            else {
                oField.UndoNotAppliedChanges();
                if (oField.IsChanged() == false) {
                    oField.SetDrawFromStream(true);
                }
            }
    
            oField.SetNeedCommit(false);

            return isValid;
        }, AscDFH.historydescription_Pdf_FieldCommit, this);
    };
    CPDFDoc.prototype.EnterDownActiveField = function() {
        this.SetGlobalHistory();
        
        let oViewer = editor.getDocumentRenderer();
        let oDrDoc  = this.GetDrawingDocument();
        let oForm   = this.activeForm;

        if (!oForm)
            return;
        
        if (false == oForm.IsInForm())
            return;

        oForm.SetInForm(false);

        if ([AscPDF.FIELD_TYPES.checkbox, AscPDF.FIELD_TYPES.radiobutton].includes(oForm.GetType())) {
            oForm.onMouseUp();
        }
        else {
            oForm.SetDrawHighlight(true);
            oForm.UpdateScroll && oForm.UpdateScroll(false); // убираем скролл

            if (oForm.IsNeedRevertShiftView()) {
                oForm.RevertContentView();
            }

            if (oForm.IsNeedCommit()) {
                this.CommitField(oForm);
            }
            
            if (oForm.IsChanged() == false) {
                oForm.SetDrawFromStream(true);
            }

            if (oForm && oForm.content && oForm.content.IsSelectionUse()) {
                oForm.content.RemoveSelection();
                oViewer.onUpdateOverlay();
            }

            oDrDoc.TargetEnd(); // убираем курсор
            oForm.AddToRedraw();
        }
    };
    CPDFDoc.prototype.OnMouseDown = function(x, y, e) {
        Asc.editor.sendEvent('asc_onHidePdfFormsActions');
        Asc.editor.SetShowTextSelectPanel(true);

        let oViewer = this.Viewer;
        if (!oViewer.canInteract()) {
            return;
        }

        let oController                 = this.GetController();
        let oDrDoc                      = this.GetDrawingDocument();
        oDrDoc.UpdateTargetFromPaint    = true

        let IsOnDrawer      = this.Api.isDrawInkMode();
        let IsOnEraser      = this.Api.isEraseInkMode();
        let IsOnAddAddShape = this.Api.isStartAddShape;
        let IsPageHighlight = this.Api.IsCommentMarker();

        let oMouseDownLink      = oViewer.getPageLinkByMouse();
        let oMouseDownField     = oViewer.getPageFieldByMouse();
        let oMouseDownAnnot     = oViewer.getPageAnnotByMouse();
        let oMouseDownDrawing   = oViewer.getPageDrawingByMouse();

        // координаты клика на странице в MM
        var pageObject = oViewer.getPageByCoords2(x, y);
        if (!pageObject)
            return false;

        let X = pageObject.x;
        let Y = pageObject.y;
        
        // если ластик
        if (IsOnEraser) {
            if (oMouseDownAnnot && oMouseDownAnnot.IsInk()) {
                this.EraseInk(oMouseDownAnnot);
            }

            return;
        }
        // если добавление шейпа
        else if (IsOnAddAddShape) {
            if (!oController.isPolylineAddition()) {
                oController.startAddShape(this.Api.addShapePreset);
            }
            
            oController.OnMouseDown(e, X, Y, pageObject.index);
            return;
        }
        // если рисование
        else if (IsOnDrawer == true) {
            oController.OnMouseDown(e, X, Y, pageObject.index);
            return;
        }
        // если хайлайт (аннотация) текста на странице (селектим текст на странице, если не попали в фигуру в режиме view).
        // если попали в фигуру, то селектим в ней (т.к. это типо текст на странице)
        else if (IsPageHighlight) {
            oViewer.isMouseMoveBetweenDownUp = true;
            this.BlurActiveObject();
            
            if (null == oMouseDownDrawing) {
                oViewer.onMouseDownEpsilon(e);
                return;
            }
        }
        
        let oCurObject = this.GetMouseDownObject();
        
        // докидываем в селект
        if (e.CtrlKey && (oCurObject && oCurObject.IsDrawing() && oMouseDownDrawing && oCurObject != oMouseDownDrawing) && oMouseDownDrawing.GetPage() == oMouseDownDrawing.GetPage()) {
            oController.selectObject(oMouseDownDrawing, oMouseDownDrawing.GetPage());
            return;
        }

        // оставляем текущий объет к селекте, если кликнули по нему же
        let isDrawingSelected = oMouseDownDrawing && oController.selectedObjects.includes(oMouseDownDrawing);
        let isObjectSelected = (oCurObject && ([oMouseDownField, oMouseDownAnnot, oMouseDownDrawing, oMouseDownLink].includes(oCurObject)) || isDrawingSelected);
        if (null == oCurObject || !isObjectSelected)
            this.SetMouseDownObject(oMouseDownField || oMouseDownAnnot || oMouseDownDrawing || oMouseDownLink);

        let oMouseDownObject = this.GetMouseDownObject();
        if (oMouseDownObject) {

            // если форма, то проверяем шрифт перед кликом в неё
            if (oMouseDownObject.IsForm() && false == [AscPDF.FIELD_TYPES.signature, AscPDF.FIELD_TYPES.checkbox , AscPDF.FIELD_TYPES.radiobutton].includes(oMouseDownObject.GetType())) {
                let _t = this;
                if (!this.checkFieldFont(oMouseDownObject, function() {
                    _t.SetMouseDownObject(null);
                    _t.OnMouseDown(x, y, e)
                })) {
                    return;
                }

                if ((oMouseDownObject.IsNeedDrawFromStream() || oMouseDownObject.IsNeedRecalc()) && oMouseDownObject.GetType() != AscPDF.FIELD_TYPES.button) {
                    oMouseDownObject.Recalculate();
                    oMouseDownObject.SetNeedRecalc(true);
                }
            }

            // всегда даём кликнуть по лкм, пкм даем кликнуть один раз, при первом попадании в объект
            if (AscCommon.getMouseButton(e || {}) != 2 || (AscCommon.getMouseButton(e || {}) == 2 && oCurObject != oMouseDownObject)) {
                oMouseDownObject.onMouseDown(x, y, e, pageObject.index);
            }
            
            if ((oMouseDownObject.IsDrawing() || (oMouseDownObject.IsAnnot() && oMouseDownObject.IsFreeText())) && false == oMouseDownObject.IsInTextBox()) {
                oDrDoc.TargetEnd();
            }
        }

        if (oViewer.canSelectPageText()) {
            oViewer.isMouseMoveBetweenDownUp = true;
            oViewer.onMouseDownEpsilon(e);
        }
        else if (this.mouseDownAnnot) {
            oViewer.onUpdateOverlay();
        }
        
        // если в селекте нет drawing (аннотации или шейпа) по которой кликнули, то сбрасываем селект
        if (oMouseDownObject == null || (false == oController.selectedObjects.includes(oMouseDownObject) && oController.selection.groupSelection != oMouseDownObject)) {
            oController.resetSelection();
            oController.resetTrackState();
        }

        this.UpdateInterface();
        this.private_UpdateTargetForCollaboration();
    };
    CPDFDoc.prototype.BlurActiveObject = function() {
        let oActiveObj = this.GetActiveObject();

        if (!oActiveObj)
            return;
        
        let oDrDoc      = this.GetDrawingDocument();
        let oController = this.GetController();

        oController.resetSelection();
        oController.resetTrackState();

        let oContent;
        if (oActiveObj.IsDrawing()) {
            oContent = oActiveObj.GetDocContent();
            this.activeDrawing = null;

            oActiveObj.OnBlur();
        }
        else if (oActiveObj.IsForm()) {
            oContent = oActiveObj.GetDocContent();

            oActiveObj.SetDrawHighlight(true);
            oActiveObj.UpdateScroll && oActiveObj.UpdateScroll(false); // убираем скрол
            
            // если чекбокс то выходим сразу
            if ([AscPDF.FIELD_TYPES.checkbox, AscPDF.FIELD_TYPES.radiobutton, AscPDF.FIELD_TYPES.button].includes(oActiveObj.GetType())) {
                oActiveObj.SetPressed(false);
                oActiveObj.SetHovered(false);
                oActiveObj.AddToRedraw();
                oActiveObj.Blur();
                return;
            }
            else {
                if (oActiveObj.IsNeedCommit()) {
                    this.CommitField(oActiveObj);
                }
                else {
                    if (oActiveObj.IsChanged() == false) {
                        oActiveObj.SetDrawFromStream(true);
                    }
    
                    if (oActiveObj.IsNeedRevertShiftView()) {
                        oActiveObj.RevertContentView();
                    }
                }
            }
            
            oActiveObj.AddToRedraw();
            oActiveObj.Blur();
        }
        else if (oActiveObj.IsAnnot()) {
            if (oActiveObj.IsFreeText()) {
                oActiveObj.Blur();
            }

            this.mouseDownAnnot = null;
        }

        if (oContent) {
            oDrDoc.TargetEnd();
            if (oContent.IsSelectionUse()) {
                oContent.RemoveSelection();
            }
        }

        this.SetGlobalHistory();
        this.Viewer.onUpdateOverlay();
    };
    CPDFDoc.prototype.SetMouseDownObject = function(oObject) {
        if (!oObject) {
            this.BlurActiveObject();

            this.mouseDownField         = null;
            this.mouseDownAnnot         = null;
            this.activeDrawing          = null;
            this.mouseDownLinkObject    = null;
            return;
        }

        if (this.GetActiveObject() == oObject) {
            return;
        }

        this.Viewer.file.removeSelection();

        if (oObject.IsForm && oObject.IsForm()) {
            // если попали в другую форму, то выход из текущей
            if (this.mouseDownAnnot != this.activeForm) {
                this.BlurActiveObject();
            }

            this.mouseDownField         = oObject;
            this.mouseDownAnnot         = null;
            this.activeDrawing          = null;
            this.mouseDownLinkObject    = null;
        }
        else if (oObject.IsAnnot && oObject.IsAnnot()) {
            if (oObject != this.mouseDownAnnot) {
                this.BlurActiveObject();
            }

            this.mouseDownField         = null;
            this.mouseDownAnnot         = oObject;
            this.activeDrawing          = null;
            this.mouseDownLinkObject    = null;
        }
        else if (oObject.IsDrawing && oObject.IsDrawing()) {
            if (oObject != this.activeDrawing) {
                this.BlurActiveObject();
            }

            this.mouseDownField         = null;
            this.mouseDownAnnot         = null;
            this.activeDrawing          = oObject;
            this.mouseDownLinkObject    = null;
        }
        // значит Link object
        else {
            this.mouseDownField         = null;
            this.mouseDownAnnot         = null;
            this.activeDrawing          = null;
            this.mouseDownLinkObject    = oObject;
        }
    };
    CPDFDoc.prototype.IsSelectionUse = function() {
        let oCurObject = this.GetActiveObject();

        if (oCurObject) {
            let oContent = oCurObject.GetDocContent();
    
            if (oContent) {
                return oContent.IsSelectionUse();
            }
        }

        return false;
    };
    
    function PDFSelectedContent() {
        this.Pages = [];
        this.Drawings = [];
        this.DocContent = null;
    }

    const PDF_SEL_CONTENT_TYPES = {
        EMPTY:      0,
        DRAWINGS:   1,
        CONTENT:    2
    };

    PDFSelectedContent.prototype.copy = function() {
        let oCopy = new PDFSelectedContent();

        let oIdMap = {};
        let oCopyPr = new AscFormat.CCopyObjectProperties();
        oCopyPr.idMap = oIdMap;

        for (let i = 0; i < this.Drawings.length; i++) {
            let oCopyDrawingObj = this.Drawings[i].Drawing.copy(oCopyPr);
            oIdMap[this.Drawings[i].Drawing.GetId()] = oCopyDrawingObj.GetId();
            oCopy.Drawings.push({
                Drawing: oCopyDrawingObj,
                ExtX: this.Drawings[i].ExtX,
                X: this.Drawings[i].X,
                Y: this.Drawings[i].Y,
                base64: this.Drawings[i].base64
            });
        }

        AscFormat.fResetConnectorsIds(oCopy.Drawings.map(function(drawing) {
            return drawing.Drawing;
        }), oIdMap);

        if (this.DocContent) {
            //TODO: перенести копирование в CSelectedContent;
            oCopy.DocContent = new AscCommonWord.CSelectedContent();
            let aElements = this.DocContent.Elements;
            for (let i = 0; i < aElements.length; ++i) {
                let oSelectedElement = new AscCommonWord.CSelectedElement();
                let oElement = aElements[i];
                let oParagraph = aElements[i].Element;
                oSelectedElement.SelectedAll = oElement.SelectedAll;
    
                oSelectedElement.Element = oParagraph.Copy(oParagraph.Parent, oParagraph.DrawingDocument, {});
                oCopy.DocContent.Elements[i] = oSelectedElement;
            }
        }
        
        return oCopy;
    };
    PDFSelectedContent.prototype.getContentType = function () {
        if (this.Drawings.length > 0) {
            return PDF_SEL_CONTENT_TYPES.DRAWINGS;
        } else if (this.DocContent) {
            return PDF_SEL_CONTENT_TYPES.CONTENT;
        }
        return PDF_SEL_CONTENT_TYPES.EMPTY;
    };
    PDFSelectedContent.prototype.isDrawingsContent = function () {
        return this.getContentType() === PDF_SEL_CONTENT_TYPES.DRAWINGS;
    };
    PDFSelectedContent.prototype.isDocContent = function () {
        return this.getContentType() === PDF_SEL_CONTENT_TYPES.CONTENT;
    };

    /**Returns array of PDFSelectedContent for special paste
     * @returns {Array}
     **/
    CPDFDoc.prototype.GetSelectedContent2 = function () {
        return AscFormat.ExecuteNoHistory(function () {
            let aRet = [], oIdMap;
            let oSourceFormattingContent    = new PDFSelectedContent();
            let oEndFormattingContent       = new PDFSelectedContent();
            let oImagesSelectedContent      = new PDFSelectedContent();

            let oSelectedContent, oDocContent, oController, oTargetTextObject, oGraphicFrame, oTable, oImage, dImageWidth,
                dImageHeight, bNeedSelectAll,
                oDocContentForDraw, oParagraph, aParagraphs, dMaxWidth, oCanvas, oContext, oGraphics,
                dContentHeight, nContentIndents = 30, bOldShowParaMarks, oSelector;
            let i, j;
            
            oController         = this.GetController();
            oSelector           = oController.selection.groupSelection ? oController.selection.groupSelection : oController;
            oTargetTextObject   = AscFormat.getTargetTextObject(oController);
            bNeedSelectAll      = false;
            
            if (!oTargetTextObject) {
                if (oSelector.selection.chartSelection && oSelector.selection.chartSelection.selection.title) {
                    oDocContent = oSelector.selection.chartSelection.selection.title.getDocContent();
                    if (oDocContent) {
                        bNeedSelectAll = true;
                    }
                }
            }

            function internalResetElementsFontSize(aContent) {
                for (var j = 0; j < aContent.length; ++j) {
                    if (aContent[j].Type === para_Run) {
                        if (aContent[j].Pr && AscFormat.isRealNumber(aContent[j].Pr.FontSize)) {
                            var oPr = aContent[j].Pr.Copy();
                            oPr.FontSize = undefined;
                            aContent[j].Set_Pr(oPr);
                        }
                    } else if (aContent[j].Type === para_Hyperlink) {
                        internalResetElementsFontSize(aContent[j].Content);
                    }
                }
            }
            function collectSelectedObjects(aSpTree, aCollectArray, bRecursive, oIdMap, bSourceFormatting) {
                var oSp;
                var oPr = new AscFormat.CCopyObjectProperties();
                oPr.idMap = oIdMap;
                oPr.bSaveSourceFormatting = bSourceFormatting;
                for (let i = 0; i < aSpTree.length; ++i) {
                    oSp = aSpTree[i];
                    // if(oSp.isEmptyPlaceholder())
                    // {
                    //     continue;
                    // }
                    if (oSp.selected) {
                        var oCopy;
                        if (oSp.isGroupObject()) {
                            oCopy = oSp.copy(oPr);
                            oCopy.setParent(oSp.parent);
                        } else {
                            if (!bSourceFormatting) {
                                oCopy = oSp.copy(oPr);
                                oCopy.setParent(oSp.parent);
                                if (oSp.isPlaceholder && oSp.isPlaceholder()) {
                                    oCopy.x = oSp.x;
                                    oCopy.y = oSp.y;
                                    oCopy.extX = oSp.extX;
                                    oCopy.extY = oSp.extY;
                                    oCopy.rot = oSp.rot;
                                    AscFormat.CheckSpPrXfrm(oCopy, true);
                                }
                                oCopy.convertFromSmartArt(true);
                            } else {
                                oCopy = oSp.getCopyWithSourceFormatting();
                                oCopy.convertFromSmartArt(true);
                                oCopy.setParent(oSp.parent);
                            }
            
                        }
            
                        aCollectArray.push(new DrawingCopyObject(oCopy, oSp.x, oSp.y, oSp.extX, oSp.extY, oSp.getBase64Img()));
                        if (AscCommon.isRealObject(oIdMap)) {
                            oIdMap[oSp.Id] = oCopy.Id;
                        }
                    }
                    if (bRecursive && oSp.isGroupObject()) {
                        collectSelectedObjects(oSp.spTree, aCollectArray, bRecursive, oIdMap, bSourceFormatting);
                    }
                }
            }

            if (oTargetTextObject) {
                if (!oDocContent) {
                    oDocContent = oController.getTargetDocContent();
                }

                if (oTargetTextObject.getObjectType() === AscDFH.historyitem_type_GraphicFrame && !oDocContent) {
                    if (oTargetTextObject.graphicObject) {
                        oGraphicFrame       = oTargetTextObject.copy(undefined);
                        oSelectedContent    = new AscCommonWord.CSelectedContent();
                        
                        oTargetTextObject.graphicObject.GetSelectedContent(oSelectedContent);
                        oTable = oSelectedContent.Elements[0].Element;

                        oGraphicFrame.setGraphicObject(oTable);
                        oTable.Set_Parent(oGraphicFrame);
                        oEndFormattingContent.Drawings.push(new DrawingCopyObject(oGraphicFrame, oTargetTextObject.x, oTargetTextObject.y, oTargetTextObject.extX, oTargetTextObject.extY, oTargetTextObject.getBase64Img()));
                        
                        oGraphicFrame.parent = oTargetTextObject.parent;
                        oGraphicFrame.bDeleted = false;
                        oGraphicFrame.recalculate();

                        oSourceFormattingContent.Drawings.push(new DrawingCopyObject(oGraphicFrame.getCopyWithSourceFormatting(), oTargetTextObject.x, oTargetTextObject.y, oTargetTextObject.extX, oTargetTextObject.extY, oTargetTextObject.getBase64Img()));
                        oImage = oController.createImage(oGraphicFrame.getBase64Img(), 0, 0, oGraphicFrame.extX, oGraphicFrame.extY);
                        
                        oImagesSelectedContent.Drawings.push(new DrawingCopyObject(oImage, 0, 0, oTargetTextObject.extX, oTargetTextObject.extY, oTargetTextObject.getBase64Img()));
                        oGraphicFrame.parent = null;
                        oGraphicFrame.bDeleted = true;
                    }
                } else {
                    if (oDocContent) {
                        if (bNeedSelectAll) {
                            oDocContent.SetApplyToAll(true);
                        }
                        oSelectedContent = oDocContent.GetSelectedContent();
                        oEndFormattingContent.DocContent = oSelectedContent;
                        
                        for (let i = 0; i < oSelectedContent.Elements.length; ++i) {
                            var oElem = oSelectedContent.Elements[i].Element;
                            if (oElem.GetType() === AscCommonWord.type_Paragraph) {
                                if (oElem.Pr && oElem.Pr.DefaultRunPr && AscFormat.isRealNumber(oElem.Pr.DefaultRunPr.FontSize)) {
                                    var oPr = oElem.Pr.Copy();
                                    oPr.DefaultRunPr.FontSize = undefined;
                                    oElem.Set_Pr(oPr);
                                }
                                internalResetElementsFontSize(oElem.Content);
                            }
                        }

                        oSelectedContent = oDocContent.GetSelectedContent();
                        var aContent = [];
                        for (let i = 0; i < oSelectedContent.Elements.length; ++i) {
                            oParagraph = oSelectedContent.Elements[i].Element;
                            oParagraph.Parent = oDocContent;
                            oParagraph.private_CompileParaPr();
                            aContent.push(oParagraph);
                        }
                        
                        AscFormat.SaveContentSourceFormatting(aContent, aContent, oDocContent.Get_Theme(), oDocContent.Get_ColorMap());
                        oSourceFormattingContent.DocContent = oSelectedContent;

                        var oSelectedContent2 = oDocContent.GetSelectedContent();
                        aContent = [];
                        for (let i = 0; i < oSelectedContent2.Elements.length; ++i) {
                            oParagraph = oSelectedContent2.Elements[i].Element;
                            oParagraph.Parent = oDocContent;
                            oParagraph.private_CompileParaPr();
                            aContent.push(oParagraph);
                        }

                        AscFormat.SaveContentSourceFormatting(aContent, aContent, oDocContent.Get_Theme(), oDocContent.Get_ColorMap());

                        if (bNeedSelectAll) {
                            oDocContent.SetApplyToAll(false);
                        }
                        
                        if (oSelectedContent2.Elements.length > 0) {
                            oDocContentForDraw = new AscFormat.CDrawingDocContent(oDocContent.Parent, oDocContent.DrawingDocument, 0, 0, 20000, 20000);
                            oSelectedContent2.ReplaceContent(oDocContentForDraw);

                            var oCheckParagraph, aRuns;
                            for (i = oDocContentForDraw.Content.length - 1; i > -1; --i) {
                                oCheckParagraph = oDocContentForDraw.Content[i];
                                if (!oCheckParagraph.IsEmpty()) {
                                    aRuns = oCheckParagraph.Content;
                                    if (aRuns.length > 1) {
                                        for (j = aRuns.length - 2; j > -1; --j) {
                                            var oRun = aRuns[j];
                                            if (oRun.Type === para_Run) {
                                                for (var k = oRun.Content.length - 1; k > -1; --k) {
                                                    if (oRun.Content[k].Type === para_NewLine) {
                                                        oRun.Content.splice(k, 1);
                                                    } else {
                                                        break;
                                                    }
                                                }
                                                if (oRun.Content.length === 0) {
                                                    aRuns.splice(j, 1);
                                                } else {
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                                if (oCheckParagraph.IsEmpty()) {
                                    oDocContentForDraw.Internal_Content_Remove(i, 1, false);
                                } else {

                                    break;
                                }
                            }
                            for (let i = 0; i < oDocContentForDraw.Content.length; ++i) {
                                oCheckParagraph = oDocContentForDraw.Content[i];
                                if (!oCheckParagraph.IsEmpty()) {
                                    aRuns = oCheckParagraph.Content;
                                    if (aRuns.length > 1) {
                                        for (j = 0; j < aRuns.length - 1; ++j) {
                                            var oRun = aRuns[j];
                                            if (oRun.Type === para_Run) {
                                                for (var k = 0; k < oRun.Content.length; ++k) {
                                                    if (oRun.Content[k].Type === para_NewLine) {
                                                        oRun.Content.splice(k, 1);
                                                        k--;
                                                    } else {
                                                        break;
                                                    }
                                                }
                                                if (oRun.Content.length === 0) {
                                                    aRuns.splice(j, 1);
                                                    j--;
                                                } else {
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                                if (oCheckParagraph.IsEmpty()) {
                                    oDocContentForDraw.Internal_Content_Remove(i, 1, false);
                                    i--;
                                } else {

                                    break;
                                }
                            }
                            if (oDocContentForDraw.Content.length > 0) {

                                oDocContentForDraw.Reset(0, 0, 20000, 20000);
                                oDocContentForDraw.Recalculate_Page(0, true);
                                aParagraphs = oDocContentForDraw.Content;
                                dMaxWidth = 0;
                                for (let i = 0; i < aParagraphs.length; ++i) {
                                    oParagraph = aParagraphs[i];
                                    for (j = 0; j < oParagraph.Lines.length; ++j) {
                                        if (oParagraph.Lines[j].Ranges[0].W > dMaxWidth) {
                                            dMaxWidth = oParagraph.Lines[j].Ranges[0].W;
                                        }
                                    }
                                }
                                dMaxWidth += 1;


                                oDocContentForDraw.Reset(0, 0, dMaxWidth, 20000);
                                oDocContentForDraw.Recalculate_Page(0, true);
                                dContentHeight = oDocContentForDraw.GetSummaryHeight();

                                var oTextWarpObject = null;
                                if (oDocContentForDraw.Parent && oDocContentForDraw.Parent.parent && oDocContentForDraw.Parent.parent instanceof AscFormat.CShape) {
                                    oTextWarpObject = oDocContentForDraw.Parent.parent.checkTextWarp(oDocContentForDraw, oDocContentForDraw.Parent.parent.getBodyPr(), dMaxWidth, dContentHeight, true, false);
                                }


                                oCanvas = document.createElement('canvas');
                                dImageWidth = dMaxWidth;
                                dImageHeight = dContentHeight;
                                oCanvas.width = ((dImageWidth * AscCommon.g_dKoef_mm_to_pix) + 2 * nContentIndents + 0.5) >> 0;
                                oCanvas.height = ((dImageHeight * AscCommon.g_dKoef_mm_to_pix) + 2 * nContentIndents + 0.5) >> 0;
                                //if (AscCommon.AscBrowser.isRetina) {
                                //    oCanvas.width <<= 1;
                                //    oCanvas.height <<= 1;
                                //}

                                var sImageUrl;
                                if (!window["NATIVE_EDITOR_ENJINE"]) {
                                    oContext = oCanvas.getContext('2d');
                                    oGraphics = new AscCommon.CGraphics();

                                    oGraphics.init(oContext, oCanvas.width, oCanvas.height, dImageWidth + 2.0 * nContentIndents / AscCommon.g_dKoef_mm_to_pix, dImageHeight + 2.0 * nContentIndents / AscCommon.g_dKoef_mm_to_pix);
                                    oGraphics.m_oFontManager = AscCommon.g_fontManager;
                                    oGraphics.m_oCoordTransform.tx = nContentIndents;
                                    oGraphics.m_oCoordTransform.ty = nContentIndents;
                                    oGraphics.transform(1, 0, 0, 1, 0, 0);

                                    if (oTextWarpObject && oTextWarpObject.oTxWarpStructNoTransform) {
                                        oTextWarpObject.oTxWarpStructNoTransform.draw(oGraphics, oDocContentForDraw.Parent.parent.Get_Theme(), oDocContentForDraw.Parent.parent.Get_ColorMap());
                                    } else {

                                        bOldShowParaMarks = this.Api.ShowParaMarks;
                                        this.Api.ShowParaMarks = false;
                                        oDocContentForDraw.Draw(0, oGraphics);
                                        this.Api.ShowParaMarks = bOldShowParaMarks;
                                    }
                                    sImageUrl = oCanvas.toDataURL("image/png");
                                } else {
                                    sImageUrl = "";
                                }

                                oImage = oController.createImage(sImageUrl, 0, 0, oCanvas.width * AscCommon.g_dKoef_pix_to_mm, oCanvas.height * AscCommon.g_dKoef_pix_to_mm);
                                oImagesSelectedContent.Drawings.push(new DrawingCopyObject(oImage, 0, 0, dImageWidth, dImageHeight, sImageUrl));
                            }
                        }
                    }
                }
            } else {
                var bRecursive = isRealObject(oController.selection.groupSelection);
                var aSpTree = bRecursive ? oController.selection.groupSelection.spTree : this.Viewer.pagesInfo.pages[this.GetCurPage()].drawings;
                oIdMap = {};

                collectSelectedObjects(aSpTree, oEndFormattingContent.Drawings, bRecursive, oIdMap);
                AscFormat.fResetConnectorsIds(oEndFormattingContent.Drawings, oIdMap);
                oIdMap = {};
                collectSelectedObjects(aSpTree, oSourceFormattingContent.Drawings, bRecursive, oIdMap, true);
                AscFormat.fResetConnectorsIds(oSourceFormattingContent.Drawings, oIdMap);
                let oImageData = oController.getSelectionImageData();
                if (oImageData) {
                    let oBounds = oImageData.bounds;
                    let sImageUrl = oImageData.src;
                    oImage = oController.createImage(sImageUrl, oBounds.min_x * AscCommon.g_dKoef_pix_to_mm, oBounds.min_y * AscCommon.g_dKoef_pix_to_mm, (oImageData.width) * AscCommon.g_dKoef_pix_to_mm, (oImageData.height) * AscCommon.g_dKoef_pix_to_mm);
                    oImagesSelectedContent.Drawings.push(new DrawingCopyObject(oImage, 0, 0, (oImageData.width) * AscCommon.g_dKoef_pix_to_mm, (oImageData.height) * AscCommon.g_dKoef_pix_to_mm, sImageUrl));
                }
            }

            aRet.push(oEndFormattingContent);
            aRet.push(oSourceFormattingContent);
            aRet.push(oImagesSelectedContent);
            return aRet;
        }, this, []);
    };

    CPDFDoc.prototype.EraseInk = function(oInk) {
        this.DoAction(function() {
            this.RemoveAnnot(oInk.GetId());
        }, AscDFH.historydescription_Pdf_EraseInk, this);
    };

    CPDFDoc.prototype.OnMouseMove = function(x, y, e) {
        let oViewer = this.Viewer;
        if (!oViewer.canInteract()) {
            return;
        }

        this.Api.sync_MouseMoveStartCallback();

        let oController     = this.GetController();
        let oDrDoc          = this.GetDrawingDocument();
        
        let IsOnDrawer      = this.Api.isDrawInkMode();
        let IsOnEraser      = this.Api.isEraseInkMode();
        let IsOnAddAddShape = this.Api.isStartAddShape;
        let IsPageHighlight = this.Api.IsCommentMarker();

        let oMouseMoveLink          = oViewer.getPageLinkByMouse();
        let oMouseMoveField         = oViewer.getPageFieldByMouse();
        let oMouseMoveAnnot         = oViewer.getPageAnnotByMouse();
        let oMouseMoveDrawing       = oViewer.getPageDrawingByMouse();

        // координаты клика на странице в MM
        let pageObject = oViewer.getPageByCoords2(x, y);
        if (!pageObject)
            return false;
        let pageObjectOrig = oViewer.getPageByCoords(x, y);
        if (!pageObjectOrig)
            return false;

        let X = pageObject.x;
        let Y = pageObject.y;

        this.CollaborativeEditing.Check_ForeignCursorsLabels(X, Y, pageObject.index);
        this.CollaborativeEditing.Check_ForeignSelectedObjectsLabels(pageObjectOrig.x, pageObjectOrig.y, pageObject.index);

        // при зажатой мышке
        if (oViewer.isMouseDown)
        {
            // под ластиком стираем только ink аннотации
            if (IsOnEraser) {
                if (oMouseMoveAnnot && oMouseMoveAnnot.IsInk()) {
                    this.EraseInk(oMouseMoveAnnot);
                }

                return;
            }
            // рисуем ink линию или добавляем фигугу
            else if (IsOnDrawer || IsOnAddAddShape) {
                oController.OnMouseMove(e, X, Y, pageObject.index);
            }
            // обработка mouseMove в полях
            else if (this.activeForm) {
                // селект текста внутри формы с редаткриуемым текстом
                if ([AscPDF.FIELD_TYPES.text, AscPDF.FIELD_TYPES.combobox].includes(this.activeForm.GetType())) {
                    this.SelectionSetEnd(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y, e);
                }
                // отрисовка нажатого/отжатого состояния кнопок/чекбоксов при входе выходе мыши в форму
                else if ([AscPDF.FIELD_TYPES.button, AscPDF.FIELD_TYPES.checkbox, AscPDF.FIELD_TYPES.radiobutton].includes(this.activeForm.GetType())) {
                    if (oMouseMoveField != this.activeForm && this.activeForm.IsHovered()) {
                        this.activeForm.SetHovered(false);
                        this.activeForm.DrawUnpressed();
                    }
                    else if (oMouseMoveField == this.activeForm && this.activeForm.IsHovered() == false) {
                        this.activeForm.SetHovered(true);
                        this.activeForm.DrawPressed();
                    }
                }
            }
            else if (this.mouseDownAnnot) {
                // freetext это кастомный шейп со своими обработками взаимодействий, поэтому нужно вызывать свой preMove (не типичный шейп)
                if (this.mouseDownAnnot.IsFreeText()) {
                    if (this.mouseDownAnnot.IsInTextBox()) {
                        this.SelectionSetEnd(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y, e);
                    }
                    else if (oController.curState instanceof AscFormat.NullState) {
                        
                        this.mouseDownAnnot.onPreMove(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y, e)
                    }
                }

                oController.OnMouseMove(e, X, Y, pageObject.index);
            }
            else if (this.activeDrawing) {
                oController.OnMouseMove(e, X, Y, pageObject.index);
                // если тянем за бордер, то не обновляем оверлей, т.к. рисуется внутри oController.OnMouseMove(e, X, Y, pageObject.index);
                if (this.activeDrawing.IsGraphicFrame() && this.activeDrawing.graphicObject.Selection.Type2 === table_Selection_Border) {
                    return;
                }
            }

            oViewer.onUpdateOverlay();
        }
        else
        {
            if (IsOnAddAddShape && oController.isPolylineAddition()) {
                oController.OnMouseMove(e, X, Y, pageObject.index);
                return;
            }

            // рисование и ластик работает только при зажатой мышке
            if (IsOnDrawer || IsOnEraser || IsOnAddAddShape || IsPageHighlight)
                return;
            
            // действия mouseEnter и mouseExit у полей
            if (oMouseMoveField != this.mouseMoveField) {
                if (this.mouseMoveField) {
                    this.mouseMoveField.onMouseExit();
                }

                this.mouseMoveField = oMouseMoveField;
                if (oMouseMoveField)
                    oMouseMoveField.onMouseEnter();
            }
        }

        this.Api.sync_MouseMoveEndCallback();
        this.UpdateCursorType(x, y, e);
    };
    CPDFDoc.prototype.UpdateCursorType = function(x, y, e) {
        let oViewer         = this.Viewer;
        let oController     = this.GetController();
        let oDrDoc          = this.GetDrawingDocument();
        
        let IsOnDrawer      = this.Api.isDrawInkMode();
        let IsOnEraser      = this.Api.isEraseInkMode();
        let IsOnAddAddShape = this.Api.isStartAddShape;

        let oMouseMoveLink          = oViewer.getPageLinkByMouse();
        let oMouseMoveField         = oViewer.getPageFieldByMouse();
        let oMouseMoveAnnot         = oViewer.getPageAnnotByMouse();
        let oMouseMoveDrawing       = oViewer.getPageDrawingByMouse();

        // координаты клика на странице в MM
        var pageObject = oViewer.getPageByCoords2(x, y);
        if (!pageObject)
            return false;

        let X       = pageObject.x;
        let Y       = pageObject.y;

        let isCursorUpdated = oController.updateCursorType(pageObject.index, X, Y, e, false);
        let oCursorInfo     = oController.getGraphicInfoUnderCursor(pageObject.index, X, Y);
        let oCurObject      = this.GetActiveObject();

        // уже обновлён в oController
        if (oCurObject && oCurObject.GetId) {
            if (oCurObject.IsAnnot() && oCurObject.IsFreeText()) {
                let isUnderCursor = oCurObject.spTree.find(function(sp) {
                    return sp.GetId() == oCursorInfo.objectId;
                });
                if (isUnderCursor) {
                    return true;
                }
            }
            else if (oCurObject.GetId() == oCursorInfo.objectId) {
                return true;
            }
        }

        // курсор залочен для этих действий
        if (IsOnDrawer || IsOnEraser || IsOnAddAddShape)
            return true;

        let defaultCursor = oViewer.MouseHandObject ? AscCommon.Cursors.Grab : "default";
        let cursorType;

        if (oMouseMoveField) {
            let pageObject = oViewer.getPageByCoords(x, y);
            if (!pageObject)
                return false;

            switch (oMouseMoveField.GetType()) {
                case AscPDF.FIELD_TYPES.text: {
                    cursorType = "text";
                    
                    if (oMouseMoveField.IsDateFormat() && oMouseMoveField.IsInForm()) {
                        // попадание в mark поля с датой
                        if (pageObject.x >= oMouseMoveField._markRect.x1 && pageObject.x <= oMouseMoveField._markRect.x2 && pageObject.y >= oMouseMoveField._markRect.y1 && pageObject.y <= oMouseMoveField._markRect.y2) {
                            cursorType = "pointer";
                        }
                    }
                    break;
                }
                case AscPDF.FIELD_TYPES.combobox: {
                    cursorType = "text";

                    // попадание в mark выбора элементов списка
                    if (pageObject.x >= oMouseMoveField._markRect.x1 && pageObject.x <= oMouseMoveField._markRect.x2 && pageObject.y >= oMouseMoveField._markRect.y1 && pageObject.y <= oMouseMoveField._markRect.y2 && oMouseMoveField._options.length != 0) {
                        cursorType = "pointer";
                    }
                    break;
                }
                default:
                    cursorType = "pointer";
            }
        }
        else if (oMouseMoveAnnot) {

            if (oMouseMoveAnnot.IsComment()) {
                cursorType = "move";
            }
            else if (oMouseMoveAnnot.IsTextMarkup()) {
                cursorType = "default";
            }
            else if (oMouseMoveAnnot.IsFreeText() && oMouseMoveAnnot.IsInTextBox()) {
                cursorType = "text";
            }
        }
        else if (oMouseMoveLink) {
            cursorType = "pointer";
        }

        // если не обновлен по drawing объектам и не задан по объектам из pdf то выставляем дефолтный
        if (cursorType == undefined) {
            if (isCursorUpdated == false) {
                cursorType = defaultCursor;
                oViewer.setCursorType(cursorType);
            }
        }
        else {
            oViewer.setCursorType(cursorType);
        }

        return true;
    };
    CPDFDoc.prototype.OnMouseUp = function(x, y, e) {
        let oViewer = this.Viewer;
        if (!oViewer.canInteract()) {
            return;
        }
        
        let oController     = this.GetController();
        let oDrDoc          = this.GetDrawingDocument();
        
        let IsOnDrawer      = this.Api.isDrawInkMode();
        let IsOnEraser      = this.Api.isEraseInkMode();
        let IsOnAddAddShape = this.Api.isStartAddShape;

        let oMouseUpLink        = oViewer.getPageLinkByMouse();
        let oMouseUpField       = oViewer.getPageFieldByMouse();
        let oMouseUpAnnot       = oViewer.getPageAnnotByMouse();
        let oMouseUpDrawing     = oViewer.getPageDrawingByMouse();

        // координаты клика на странице в MM
        var pageObject = oViewer.getPageByCoords2(x, y);
        if (!pageObject)
            return false;

        let X = pageObject.x;
        let Y = pageObject.y;

        // ластик работает на mousedown
        if (IsOnEraser) {
            return;
        }
        // если рисование или добавление шейпа то просто заканчиваем его
        else if (IsOnDrawer || IsOnAddAddShape) {
            oController.OnMouseUp(e, X, Y, pageObject.index);
            return;
        }

        if (this.mouseDownField && oMouseUpField == this.mouseDownField) {
            this.OnMouseUpField(oMouseUpField, e);
        }
        else if (this.mouseDownAnnot) {
            oController.OnMouseUp(e, X, Y, pageObject.index);
            if (this.mouseDownAnnot == oMouseUpAnnot)
                oMouseUpAnnot.onMouseUp(x, y, e);
        }
        else if (this.activeDrawing) {
            oController.OnMouseUp(e, X, Y, pageObject.index);
            if (this.Api.isMarkerFormat && !this.Api.IsCommentMarker() && this.HighlightColor && this.activeDrawing.IsInTextBox()) {
                this.SetHighlight(this.HighlightColor.r, this.HighlightColor.g, this.HighlightColor.b);
            }

            oController.updateCursorType(pageObject.index, X, Y, e, false);
            oDrDoc.UnlockCursorType();
        }
        else if (this.mouseDownLinkObject && this.mouseDownLinkObject == oMouseUpLink) {
            oViewer.navigateToLink(oMouseUpLink);
        }
        
        this.UpdateInterface();
        oViewer.onUpdateOverlay();
        oViewer.file.onUpdateSelection();
    };

    CPDFDoc.prototype.OnMouseUpField = function(oField) {
        oField.onMouseUp();
        
        if ([AscPDF.FIELD_TYPES.checkbox, AscPDF.FIELD_TYPES.radiobutton].includes(oField.GetType())) {
            if (oField.IsNeedCommit() && this.IsNeedDoCalculate()) {
                this.DoCalculateFields();
                this.CommitFields();
            }
        }
    };
    CPDFDoc.prototype.DoUndo = function() {
        let oDrDoc = this.GetDrawingDocument();
        oDrDoc.UpdateTargetFromPaint = true;

        if (this.History.Can_Undo() && !this.LocalHistory.Can_Undo())
            this.SetGlobalHistory();
        
        if (AscCommon.History == this.History && true !== this.History.Can_Undo() && this.Api && this.CollaborativeEditing && true === this.CollaborativeEditing.Is_Fast() && true !== this.CollaborativeEditing.Is_SingleUser()) {
            if (this.CollaborativeEditing.CanUndo() && true === this.Api.canSave) {
                this.CollaborativeEditing.Set_GlobalLock(true);
                this.Api.forceSaveUndoRequest = true;
            }
        }

        if (AscCommon.History.Can_Undo())
        {
            let oActive = this.GetActiveObject();
            if (oActive) {
                let oContent = oActive.GetDocContent();
                if (oContent) {
                    oContent.RemoveSelection();
                }
            }

            this.ClearSearch();
            this.currInkInDrawingProcess = null;

            let nCurPoindIdx = AscCommon.History.Index;
            let oCurPoint = AscCommon.History.Points[nCurPoindIdx];

            let arrChanges = AscCommon.History.Undo();
            let recalcData = AscCommon.History.Get_RecalcData(null, arrChanges);
            
            AscCommon.History.UndoRedoInProgress = true;

            let aSourceObjects  = oCurPoint.Additional.Pdf || [];
               
            for (let i = 0; i < aSourceObjects.length; i++) {
                let oSourceObj = aSourceObjects[i];

                if (oSourceObj.IsForm()) {
                    // в глобальной истории должен срабатывать commit
                    if (AscCommon.History == this.History) {
                        oDrDoc.TargetEnd(); // убираем курсор
                        
                        if (this.activeForm) {
                            this.activeForm.UpdateScroll && this.activeForm.UpdateScroll(false);
                            this.activeForm.SetDrawHighlight(true);
                            this.activeForm = null;
                        }
                    }
    
                    oSourceObj.SetNeedRecalc(true);
                }
            }
            
            AscCommon.History.UndoRedoInProgress = false;
            this.UpdateInterface();
        }
    };
    CPDFDoc.prototype.DoRedo = function() {
        let oDrDoc = this.GetDrawingDocument();
        oDrDoc.UpdateTargetFromPaint = true;

        if (this.History.Can_Redo() && !this.LocalHistory.Can_Redo())
            this.SetGlobalHistory();

        if (AscCommon.History.Can_Redo())
        {
            let oActive = this.GetActiveObject();
            if (oActive) {
                let oContent = oActive.GetDocContent();
                if (oContent) {
                    oContent.RemoveSelection();
                }
            }

            this.ClearSearch();
            this.currInkInDrawingProcess = null;

            let arrChanges = AscCommon.History.Redo();
            let recalcData = AscCommon.History.Get_RecalcData(null, arrChanges);

            AscCommon.History.UndoRedoInProgress = true;

            let nCurPoindIdx = AscCommon.History.Index;
            let oCurPoint = AscCommon.History.Points[nCurPoindIdx];

            let aSourceObjects  = oCurPoint.Additional.Pdf || [];

            for (let i = 0; i < aSourceObjects.length; i++) {
                let oSourceObj = aSourceObjects[i];

                if (oSourceObj.IsForm()) {
                    if (AscCommon.History == this.History) {
                        oDrDoc.TargetEnd(); // убираем курсор
                            
                        if (this.activeForm) {
                            this.activeForm.UpdateScroll && this.activeForm.UpdateScroll(false);
                            this.activeForm.SetDrawHighlight(true);
                            this.activeForm = null;
                        }
                    }
    
                    oSourceObj.SetNeedRecalc(true);
                }
                else if (oSourceObj.IsAnnot() || oSourceObj.IsDrawing()) {
                    oSourceObj.SetNeedRecalc(true);
                }
            }

            AscCommon.History.UndoRedoInProgress = false;
            this.UpdateInterface();
        }
    };
    CPDFDoc.prototype.SetNeedUpdateTarget = function(bUpdate) {
        this.NeedUpdateTarget = bUpdate;
    };
    
    /**
	 * Получает активный объект
	 */
    CPDFDoc.prototype.GetActiveObject = function() {
        return this.activeForm || this.mouseDownAnnot || this.activeDrawing;
    };
    /**
	 * Разница от предыдущего метода в том, что для полей будет полочено поле, в которое был клик, а не активное
     * так как после клика в поле, оно может перестать быть активный после выполнения каких либо actions
	 */
    CPDFDoc.prototype.GetMouseDownObject = function() {
        return this.mouseDownField || this.mouseDownAnnot || this.activeDrawing;
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
        this.calculateInfo.ids.forEach(function(id) {
            let oField = oThis.GetFieldBySourceIdx(id);
            if (!oField)
                return;
            
            let oCalcTrigget = oField.GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Calculate);
            if (oCalcTrigget == null && oField._kids[0]) {
                // to do: action действие должно быть в родителе одинаковых виджетов
                oCalcTrigget = oField._kids[0].GetTrigger(AscPDF.FORMS_TRIGGERS_TYPES.Calculate);
                if (oCalcTrigget)
                    oField = oField._kids[0];
            }
            let oActionRunScript = oCalcTrigget ? oCalcTrigget.GetActions()[0] : null;

            if (oActionRunScript) {
                oThis.StartNoHistoryMode();
                oActionRunScript.RunScript();
                oThis.EndNoHistoryMode();
                if (oField.IsNeedCommit()) {
                    oField.SetNeedRecalc(true);
                    oThis.fieldsToCommit.push(oField);
                }
            }
        });
        this.calculateInfo.SetIsInProgress(false);
        this.calculateInfo.SetSourceField(null);
    };
    CPDFDoc.prototype.IsNeedDoCalculate = function() {
        if (this.calculateInfo.ids.length > 0 && false == AscCommon.History.UndoRedoInProgress)
            return true;

        return false;
    }

    CPDFDoc.prototype.GetCalculateInfo = function() {
        return this.calculateInfo;
    };

    CPDFDoc.prototype.GetActionsQueue = function() {
        return this.actionsInfo;
    };
    
    CPDFDoc.prototype.EscapeForm = function() {
        this.SetGlobalHistory();
        if (this.activeForm && this.activeForm.IsNeedDrawHighlight() == false) {
            this.activeForm.UndoNotAppliedChanges();

            if (this.activeForm.IsChanged() == false)
                this.activeForm.SetDrawFromStream(true);

            this.activeForm.AddToRedraw();
            this.activeForm.SetDrawHighlight(true);
            this.GetDrawingDocument().TargetEnd();
        }
    };

    /**
	 * Adds a new page to the active document.
	 * @memberof CPDFDoc
	 * @typeofeditors ["PDF"]
     * @param {number} [nPos] - (optional) The page after which to add the new page in a 1-based page numbering
	 * @returns {boolean}
	 */
    CPDFDoc.prototype.AddPage = function(nPos, oPage) {
        let oViewer     = editor.getDocumentRenderer();
        let oFile       = oViewer.file;
        let oController = this.GetController();

        if (!oPage) {
            oPage = {
                Dpi: 72,
                fonts: [],
                Rotate: 0
            }
        }

        if (nPos === undefined || -1 === nPos)
            nPos = oFile.pages.length;
        if (oPage.W === undefined)
            oPage.W = oFile.pages[Math.max(nPos - 1, 0)].W;
        if (oPage.H === undefined)
            oPage.H = oFile.pages[Math.max(nPos - 1, 0)].H;

        oFile.addPage(nPos, oPage);
	
		oViewer.drawingPages.splice(nPos, 0, {
			X : 0,
			Y : 0,
			W : (oFile.pages[nPos].W * 96 / oFile.pages[nPos].Dpi) >> 0,
			H : (oFile.pages[nPos].H * 96 / oFile.pages[nPos].Dpi) >> 0,
			Image : undefined
		});

        let oPageInfo;
        // from history
        if (oPage.Id) {
            oPageInfo = AscCommon.g_oTableId.GetById(oPage.Id);
        }
        // to history
        else {
            oPageInfo = new AscPDF.CPageInfo();
            oPage.Id = oPageInfo.Id;
        }

        if (oViewer.pagesInfo.pages.length == 0)
            oViewer.pagesInfo.setCount(1);
        else
            oViewer.pagesInfo.pages.splice(nPos, 0, oPageInfo);

        // can be uninitialized on Apply_Changes
        if (oViewer.thumbnails) {
            oViewer.thumbnails._addPage(nPos);
            oViewer.thumbnails.resetSelection();
        }

        oViewer.resize(true);

        for (let i = 0; i < oViewer.file.pages.length; i++) {
            oController.mergeDrawings(i);
        }
        this.GetDrawingDocument().m_lPagesCount = oViewer.file.pages.length;

        oViewer.sendEvent("onPagesCount", oFile.pages.length);

        this.History.Add(new CChangesPDFDocumentPagesContent(this, nPos, [oPage], true));

        this.History.StartNoHistoryMode();
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
            if (oViewer.pagesInfo.pages[nPage].drawings) {
                oViewer.pagesInfo.pages[nPage].drawings.forEach(function(drawing) {
                    drawing.SetPage(nPage);
                });
            }
        }
        this.History.EndNoHistoryMode();

        oViewer.paint();
    };

    /**
	 * Removes a page from document.
	 * @memberof CPDFDoc
	 * @typeofeditors ["PDF"]
     * @param {number} [nPos = 0] 
	 * @returns {boolean}
	 */
    CPDFDoc.prototype.RemovePage = function(nPos) {
        let oThis       = this;
        let oViewer     = editor.getDocumentRenderer();
        let oFile       = oViewer.file;
        let oController = this.GetController();
        
        if (oFile.pages.length == 1)
            return false;
        if (!AscCommon.isNumber(nPos) || nPos < 0)
            nPos = 0;

        oFile.removeSelection();
        
        // сначала удаляем все объекты со страницы
        if (oViewer.pagesInfo.pages[nPos].fields) {
            oViewer.pagesInfo.pages[nPos].fields.slice().forEach(function(field) {
                oThis.RemoveForm(field);
            });
        }
        if (oViewer.pagesInfo.pages[nPos].annots) {
            oViewer.pagesInfo.pages[nPos].annots.slice().forEach(function(annot) {
                oThis.RemoveAnnot(annot.GetId());
            });
        }
        if (oViewer.pagesInfo.pages[nPos].drawings) {
            oViewer.pagesInfo.pages[nPos].drawings.slice().forEach(function(drawing) {
                oThis.RemoveDrawing(drawing.GetId());
            });
        }

        // убираем информацию о странице
        let aPages = oFile.removePage(nPos);
		oViewer.drawingPages.splice(nPos, 1);
        let aPagesInfo = oViewer.pagesInfo.pages.splice(nPos, 1);
        
        // to history
        aPages[0].Id = aPagesInfo[0].Id;

        // can be uninitialized on Apply_Changes
        if (oViewer.thumbnails) {
            oViewer.thumbnails._deletePage(nPos);
            oViewer.thumbnails.resetSelection();
        }

        oViewer.checkVisiblePages();
        oViewer.resize(true);
        for (let i = 0; i < oViewer.file.pages.length; i++) {
            oController.mergeDrawings(i);
        }
        this.GetDrawingDocument().m_lPagesCount = oViewer.file.pages.length;
        oViewer.sendEvent("onPagesCount", oFile.pages.length);

        this.History.Add(new CChangesPDFDocumentPagesContent(this, nPos, aPages, false));

        this.History.StartNoHistoryMode();
        // проставляем новые номера страниц объектам на остальных страницах
        for (let nPage = nPos; nPage < oViewer.pagesInfo.pages.length; nPage++) {
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
            if (oViewer.pagesInfo.pages[nPage].drawings) {
                oViewer.pagesInfo.pages[nPage].drawings.forEach(function(drawing) {
                    drawing.SetPage(nPage);
                });
            }
        }
        this.History.EndNoHistoryMode();

        oViewer.paint();
    };
    CPDFDoc.prototype.RemovePages = function(aIndexes) {
        aIndexes.sort(function(a, b) {
            return a - b;
        });
    
        for (var i = 0; i < aIndexes.length; i++) {
            this.RemovePage(aIndexes[i] - i);
        }
    };
    CPDFDoc.prototype.SetPageRotate = function(nPage, nAngle) {
		let oViewer     = this.Viewer;
        let oPageInfo   = this.GetPageInfo(nPage);

        oPageInfo.SetRotate(nAngle);

        // sticky note всегда неповернуты
        oViewer.pagesInfo.pages[nPage].annots.forEach(function(annot) {
            if (annot.IsComment()) {
                annot.AddToRedraw();
            }
        });
        
		oViewer.resize(true);
        oViewer.paint();
    };
    CPDFDoc.prototype.RotatePages = function(aIndexes, nAngle) {
        let oViewer = this.Viewer;

        for (let i = 0; i < aIndexes.length; i++) {
            let nNewPageAngle = (oViewer.getPageRotate(aIndexes[i]) + nAngle) % 360;
            if (nNewPageAngle < 0) {
                nNewPageAngle += 360;
            }

            this.SetPageRotate(aIndexes[i], nNewPageAngle);
        }
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

        let oPagesInfo = oViewer.pagesInfo;
        if (!oPagesInfo.pages[nPageNum])
            return null;
        
        let oField = private_createField(cName, cFieldType, nPageNum, aCoords, this);
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
    CPDFDoc.prototype.AddAnnotByProps = function(oProps) {
        let nPageNum = oProps.page;

        let oPageInfo = this.GetPageInfo(nPageNum);
        if (!oPageInfo)
            return null;
        
        AscCommon.g_oTableId.TurnOn();
        let oAnnot = AscPDF.CreateAnnotByProps(oProps, this);
        if (oProps.apIdx != undefined) {
            oAnnot.SetApIdx(oProps.apIdx);
        }

        oAnnot.SetNeedRecalc(true);
        oAnnot.SetDisplay(this.IsAnnotsHidden() ? window["AscPDF"].Api.Objects.display["hidden"] : window["AscPDF"].Api.Objects.display["visible"]);
        
        this.annots.push(oAnnot);
        oPageInfo.AddAnnot(oAnnot, nPageNum);
        
        
        return oAnnot;
    };
    CPDFDoc.prototype.AddComment = function(AscCommentData) {
        let oCurHistory = AscCommon.History;
        AscCommon.History = this.History;

        let oActiveAnnot = this.mouseDownAnnot;

        let res = this.DoAction(function() {
            let oViewer     = editor.getDocumentRenderer();
            let pageObject  = oViewer.getPageByCoords(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
            let nPage       = pageObject ? pageObject.index : this.GetCurPage();
            let posToAdd    = this.anchorPositionToAdd ? this.anchorPositionToAdd : {x: 10, y: 10};
            
            let X2 = posToAdd.x + 20;
            let Y2 = posToAdd.y + 20;

            let oProps = {
                rect:           [posToAdd.x, posToAdd.y, X2, Y2],
                page:           nPage,
                name:           AscCommon.CreateGUID(),
                type:           AscPDF.ANNOTATIONS_TYPES.Text,
                author:         AscCommentData.m_sUserName,
                modDate:        AscCommentData.m_sOOTime,
                creationDate:   AscCommentData.m_sOOTime,
                contents:       AscCommentData.m_sText,
                hidden:         false
            }

            this.anchorPositionToAdd = null;

            let oStickyComm;
            if (this.mouseDownAnnot) {
                if (this.mouseDownAnnot.IsUseContentAsComment() && !this.mouseDownAnnot.GetContents()) {
                    // If the annotation uses content as comment and there's no content, set the content
                    this.mouseDownAnnot.SetContents(AscCommentData.m_sText);
                }
                else {
                    let oDataForEdit;

                    // For all other cases, add a reply to the comment
                    let newCommentData = new AscCommon.CCommentData();
                    newCommentData.Read_FromAscCommentData(AscCommentData);
                
                    // if freetext or line with cap
                    if (!this.mouseDownAnnot.IsUseContentAsComment() && this.mouseDownAnnot.GetReply(0) == null) {
                        oDataForEdit = newCommentData;
                    }
                    else {
                        let curCommentData = new AscCommon.CCommentData();
                        curCommentData.Read_FromAscCommentData(this.mouseDownAnnot.GetAscCommentData());
                        curCommentData.Add_Reply(newCommentData);
                        oDataForEdit = curCommentData;
                    }
                
                    this.EditComment(this.mouseDownAnnot.GetId(), oDataForEdit);
                }
            }
            else {
                oStickyComm = this.AddAnnotByProps(oProps);
                AscCommentData.m_sUserData = oStickyComm.GetId();
                AscCommentData.m_sQuoteText = "";
                this.CheckComment(oStickyComm);
            }
            
            if (!oStickyComm)
                this.UpdateUndoRedo();
            
            return oStickyComm;
        }, AscDFH.historydescription_Pdf_AddComment, this, oActiveAnnot && oActiveAnnot.GetId());

        AscCommon.History = oCurHistory;
        return res;
    };
    CPDFDoc.prototype.convertPixToMM = function(pix) {
        return this.GetDrawingDocument().GetMMPerDot(pix);
    };
    /**
	 * Обновляет позицию всплывающего окна комментария
	 * @memberof CPDFDoc
	 * @typeofeditors ["PDF"]
	 */
    CPDFDoc.prototype.UpdateCommentPos = function() {
        if (this.showedCommentId) {
            let oAnnot = this.GetAnnotById(this.showedCommentId);

            if (!oAnnot) {
                this.showedCommentId = undefined;
                return;
            }

            let oPos;
            let nPage       = oAnnot.GetPage();
            let nPageRotate = this.Viewer.getPageRotate(nPage);
            let aOrigRect   = oAnnot.GetOrigRect();
            let oTr         = this.pagesTransform[nPage].invert;
            
            let w = aOrigRect[2] - aOrigRect[0];
            let h = aOrigRect[3] - aOrigRect[1];

            let X, Y;
            switch (nPageRotate) {
                case 0: {
                    X = aOrigRect[0] + (oAnnot.IsComment() ? w / this.Viewer.zoom : w);
                    Y = aOrigRect[1] + (oAnnot.IsComment() ? h / 2 / this.Viewer.zoom : h / 2);
                    break;
                }
                case 90: {
                    X = aOrigRect[0] + (oAnnot.IsComment() ? w / this.Viewer.zoom / 2 : w / 2);
                    Y = aOrigRect[1];
                    break;
                }
                case 180: {
                    X = aOrigRect[0];
                    Y = aOrigRect[1] + (oAnnot.IsComment() ? h / 2 / this.Viewer.zoom : h / 2);
                    break;
                }
                case 270: {
                    X = aOrigRect[0] + (oAnnot.IsComment() ? w / 2 / this.Viewer.zoom : w / 2);
                    Y = aOrigRect[3];
                    break;
                }
            }

            oPos = oTr.TransformPoint(X, Y);

            editor.sync_UpdateCommentPosition(oAnnot.GetId(), oPos.x, oPos.y);
        }
    };
    CPDFDoc.prototype.UpdateMathTrackPos = function() {
        this.MathTrackHandler.OnChangePosition();
    };
    CPDFDoc.prototype.UpdateAnnotTrackPos = function() {
        this.AnnotTextPrTrackHandler.OnChangePosition();
    };
    CPDFDoc.prototype.UpdateSelectionTrackPos = function() {
        this.TextSelectTrackHandler.OnChangePosition();
    };
    CPDFDoc.prototype.ConvertMathView = function (isToLinear, isAll) {
        let oController = this.GetController();
        oController.convertMathView(isToLinear, isAll);
    };
    CPDFDoc.prototype.Set_MathProps = function (oMathProps) {
        let oController = this.GetController();
        oController.setMathProps(oMathProps);
    };

    CPDFDoc.prototype.CreateNewHistoryPoint = function(oAdditional) {
        if (this.IsNeedSkipHistory() || this.Viewer.IsOpenFormsInProgress || this.Viewer.IsOpenAnnotsInProgress || AscCommon.History.UndoRedoInProgress)
            return;

        if (!AscCommon.History.IsOn()) {
            AscCommon.History.TurnOn();
        }
        if (AscCommon.History.Is_LastPointEmpty()) {
            AscCommon.History.Remove_LastPoint();
        }
        
        AscCommon.History.Create_NewPoint(oAdditional ? oAdditional.description : undefined);

        if (oAdditional) {
            if (oAdditional.objects) {
                AscCommon.History.SetSourceObjectsToPointPdf(oAdditional.objects);
            }
        }
    };
	
	CPDFDoc.prototype.DoAction = function(fAction, nDescription, oThis, Additional) {
		let nChangesType;
		
		switch (nDescription) {
			case AscDFH.historydescription_Document_BackSpaceButton:
			case AscDFH.historydescription_Document_DeleteButton:
				nChangesType = AscCommon.changestype_Delete;
				break;
			case AscDFH.historydescription_Pdf_FreeTextGeom:
			case AscDFH.historydescription_CommonDrawings_EndTrack:
			case AscDFH.historydescription_Pdf_FreeTextFitTextBox:
				nChangesType = AscCommon.changestype_Drawing_Props;
				break;
			case AscDFH.historydescription_Document_ChangeComment:
			case AscDFH.historydescription_Pdf_AddComment:
			case AscDFH.historydescription_Pdf_RemoveComment:
				nChangesType = AscCommon.changestype_2_Comment;
				break;
		}
		
		if (this.IsSelectionLocked(nChangesType, Additional)) {
			return;
		}
		
		this.StartAction(nDescription);
		let result = fAction.call(oThis);
		this.FinalizeAction(true);
		return result;
	};
    /**
     * Начинаем новое действие, связанное с изменением документа
     * @param {number} nDescription - тип изменения, ex.: AscDFH.historydescription_Pdf_FieldCommit
     * @param {object} [oSelectionState=null] - начальное состояние селекта, до начала действия
     */
    CPDFDoc.prototype.StartAction = function(nDescription, oSelectionState) {
        if (this.IsNeedSkipHistory() || this.Viewer.IsOpenFormsInProgress || this.Viewer.IsOpenAnnotsInProgress || AscCommon.History.UndoRedoInProgress)
            return;
        
        let oAdditional = {};

        switch (nDescription) {
            case AscDFH.historydescription_Pdf_FieldCommit:
            case AscDFH.historydescription_Pdf_UpdateAnnotRC:
                this.SetGlobalHistory();
                break;
        }

        if (AscCommon.History.Is_LastPointEmpty()) {
            AscCommon.History.Remove_LastPoint();
        }

        var isNewPoint = AscCommon.History.Create_NewPoint(nDescription, oSelectionState);
        
        if (oAdditional.objects) {
            AscCommon.History.SetSourceObjectsToPointPdf(oAdditional.objects);
        }
        
        if (true === this.Action.Start) {
            this.Action.Depth++;

            if (isNewPoint) {
                this.Action.PointsCount++;
            }
        }
        else {
            this.Action.Start           = true;
            this.Action.Depth           = 0;
            this.Action.PointsCount     = isNewPoint ? 1 : 0;
            this.Action.Recalculate     = false;
            this.Action.Description     = nDescription;
            this.Action.UpdateSelection = false;
            this.Action.UpdateInterface = false;
            this.Action.UpdateRulers    = false;
            this.Action.UpdateUndoRedo  = false;
            this.Action.UpdateTracks    = false;
            // this.Action.Redraw.Start    = undefined;
            // this.Action.Redraw.End      = undefined;
            this.Action.Additional      = {};
        }
    };

    CPDFDoc.prototype.FinalizeAction = function(checkEmptyAction) {
		
        if (checkEmptyAction && AscCommon.History.Is_LastPointEmpty()) {
            AscCommon.History.Remove_LastPoint();
            this.UpdateInterface();
            return;
        }

		AscCommon.History.Get_RecalcData();
		AscCommon.History.Reset_RecalcIndex();
        Asc.editor.checkLastWork();
        
        if (false /* тут проверяем локи */) {
            AscCommon.History.Undo();
            AscCommon.History.Clear_Redo();
        }

        this.UpdateInterface();
    };

    CPDFDoc.prototype.Refresh_RecalcData = function(){};

    CPDFDoc.prototype.EditComment = function(Id, CommentData) {
        let oAnnotToEdit = this.annots.find(function(annot) {
            return annot.GetId() === Id;
        });

        oAnnotToEdit.EditCommentData(CommentData);
        editor.sync_ChangeCommentData(Id, CommentData);
    };
    CPDFDoc.prototype.CheckComment = function(oAnnot) {
        let bUseContentsAsComment = oAnnot.IsUseContentAsComment();
        
        if (oAnnot.IsUseInDocument()) {
            if ((bUseContentsAsComment && oAnnot.GetContents()) || (bUseContentsAsComment == false && oAnnot.GetReply(0) instanceof AscPDF.CAnnotationText)) {
                editor.sendEvent("asc_onAddComment", oAnnot.GetId(), oAnnot.GetAscCommentData());
            }
        }
    };
    CPDFDoc.prototype.TurnOffHistory = function() {
        if (AscCommon.History.IsOn() == true)
            AscCommon.History.TurnOff();
    };
    CPDFDoc.prototype.TurnOnHistory = function() {
        if (AscCommon.History.IsOn() == false)
            AscCommon.History.TurnOn();
    };
    CPDFDoc.prototype.StartNoHistoryMode = function() {
        this.TurnOffHistory();
    };
    CPDFDoc.prototype.EndNoHistoryMode = function() {
        this.TurnOnHistory();
    };

    CPDFDoc.prototype.ShowComment = function(arrId) {
        let oAnnot;
        for (let nIndex = 0, nCount = arrId.length; nIndex < nCount; ++nIndex) {
            oAnnot = this.GetAnnotById(arrId[nIndex]);

            if (oAnnot) {
                break;
            }
        }

        if (oAnnot) {
            this.showedCommentId = oAnnot.GetId();

            let oPos;
            let nPage       = oAnnot.GetPage();
            let nPageRotate = this.Viewer.getPageRotate(nPage);
            let aOrigRect   = oAnnot.GetOrigRect();
            let oTr         = this.pagesTransform[nPage].invert;
            
            let w = aOrigRect[2] - aOrigRect[0];
            let h = aOrigRect[3] - aOrigRect[1];

            let X, Y;
            switch (nPageRotate) {
                case 0: {
                    X = aOrigRect[0] + (oAnnot.IsComment() ? w / this.Viewer.zoom : w);
                    Y = aOrigRect[1] + (oAnnot.IsComment() ? h / 2 / this.Viewer.zoom : h / 2);
                    break;
                }
                case 90: {
                    X = aOrigRect[0] + (oAnnot.IsComment() ? w / this.Viewer.zoom / 2 : w / 2);
                    Y = aOrigRect[1];
                    break;
                }
                case 180: {
                    X = aOrigRect[0];
                    Y = aOrigRect[1] + (oAnnot.IsComment() ? h / 2 / this.Viewer.zoom : h / 2);
                    break;
                }
                case 270: {
                    X = aOrigRect[0] + (oAnnot.IsComment() ? w / 2 / this.Viewer.zoom : w / 2);
                    Y = aOrigRect[3];
                    break;
                }
            }

            oPos = oTr.TransformPoint(X, Y);
            Asc.editor.sync_ShowComment([this.showedCommentId], oPos.x, oPos.y);
        }
        else {
            Asc.editor.sync_HideComment();
            this.showedCommentId = undefined;
        }
    };
    
    CPDFDoc.prototype.Remove = function(nDirection, isCtrlKey) {
        let oDrDoc = this.GetDrawingDocument();
        oDrDoc.UpdateTargetFromPaint = true;

        let oForm       = this.activeForm;
        let oAnnot      = this.mouseDownAnnot;
        let oFreeText   = this.mouseDownAnnot && this.mouseDownAnnot.IsFreeText() ? this.mouseDownAnnot : null;
        let oDrawing    = this.activeDrawing;

        let oContent;
        if (oForm && oForm.IsCanEditText()) {
            oForm.Remove(nDirection, isCtrlKey);
            oContent = oForm.GetDocContent();
        }
        else if (oFreeText) {
            if (oFreeText.IsInTextBox()) {
                oFreeText.Remove(nDirection, isCtrlKey);
                oContent = oFreeText.GetDocContent();
            }
            else {
                this.RemoveAnnot(oFreeText.GetId());
            }
        }
        else if (oDrawing && !Asc.editor.isRestrictionView()) {
            if (oDrawing.IsInTextBox()) {
                oDrawing.Remove(nDirection, isCtrlKey);
                oContent = oDrawing.GetDocContent();
            }
            else {
                let oThis = this;
                let oController = this.GetController();
                let aDrawings = oController.getSelectedObjects().slice();
                aDrawings.forEach(function(drawing) {
                    oThis.RemoveDrawing(drawing.GetId());
                });
            }
        }
        else if (oAnnot && this.Viewer.isMouseDown == false) {
            this.RemoveAnnot(oAnnot.GetId());
        }

        if (oContent) {
            oDrDoc.TargetStart(true);
        }
        else {
            oDrDoc.TargetEnd();
        }
    };
    CPDFDoc.prototype.EnterDown = function(isShiftKey) {
        let oDrDoc      = this.GetDrawingDocument();
        
        let oForm       = this.activeForm;
        let oFreeText   = this.mouseDownAnnot && this.mouseDownAnnot.IsFreeText() ? this.mouseDownAnnot : null;
        let oDrawing  = this.activeDrawing;

        let oContent;
        if (oForm) {
            if (oForm.GetType() == AscPDF.FIELD_TYPES.text && oForm.IsCanEditText() && oForm.IsMultiline()) {
                Asc.editor.asc_enterText([13]);
                oContent = oForm.GetDocContent();
            }
            else {
                this.EnterDownActiveField();
            }
        }
        else if (oFreeText && oFreeText.IsInTextBox()) {
            this.AddNewParagraph();
            oContent = oFreeText.GetDocContent();
        }
        else if (oDrawing) {
            this.AddNewParagraph();
            oContent = oDrawing.GetDocContent();
        }

        if (oContent) {
            oDrDoc.TargetStart(true);
        }
    };
    
    CPDFDoc.prototype.RemoveComment = function(Id) {
        let oAnnot = this.annots.find(function(annot) {
            return annot.GetId() === Id;
        });

        if (!oAnnot)
            return;

        this.DoAction(function() {
            Asc.editor.sync_HideComment();
            if (oAnnot.IsComment()) {
                this.RemoveAnnot(oAnnot.GetId());
            }
            else {
                oAnnot.RemoveComment();
            }
        }, AscDFH.historydescription_Pdf_RemoveComment, this);
    };
    CPDFDoc.prototype.RemoveAnnot = function(Id) {
        let oController = this.GetController();

        let oAnnot = this.annots.find(function(annot) {
            return annot.GetId() === Id;
        });

        if (!oAnnot)
            return;

        let nPos = this.annots.indexOf(oAnnot);
        this.annots.splice(nPos, 1);
        
        let oPage = oAnnot.GetParentPage();
        oPage.RemoveAnnot(Id);
        
        if (this.mouseDownAnnot == oAnnot)
            this.mouseDownAnnot = null;

        Asc.editor.sync_HideComment();
        Asc.editor.sync_RemoveComment(Id);
        oController.resetSelection();
        oController.resetTrackState();
        this.private_UpdateTargetForCollaboration(true);
    };

    CPDFDoc.prototype.RemoveDrawing = function(Id) {
        let oController = this.GetController();
        let oDrawing = this.drawings.find(function(drawing) {
            return drawing.GetId() === Id;
        });

        if (!oDrawing)
            return;

        let nPos = this.drawings.indexOf(oDrawing);
        this.drawings.splice(nPos, 1);

        let oPage = oDrawing.GetParentPage();
        oPage.RemoveDrawing(Id);

        oController.resetSelection(true);
        oController.resetTrackState();

        if (this.activeDrawing == oDrawing) {
            this.activeDrawing = null;
        }

        this.ClearSearch();
        this.private_UpdateTargetForCollaboration(true);
    };

    CPDFDoc.prototype.RemoveForm = function(oForm) {
        let oController = this.GetController();
        this.BlurActiveObject();

        if (!oForm.IsWidget())
            return;

        // удаляем поле из виджетов и со страницы
        let nPos = this.widgets.indexOf(oForm);
        this.widgets.splice(nPos, 1);

        let oPage = oForm.GetParentPage();
        oPage.RemoveField(oForm.GetId());
        
        oController.resetSelection();
        oController.resetTrackState();

        this.ClearSearch();
    };
    
    /**
	 * Checks the parent form and deletes if necessary
	 * @memberof CPDFDoc
     * @param {Object} oForm - parent form
	 * @typeofeditors ["PDF"]
	 */
    CPDFDoc.prototype.CheckParentForm = function(oForm) {
        let aKids   = oForm.GetKids();
        let oParent = oForm.GetParent();

        if (aKids == 0) {
            // удаляем поле из массива родительских полей
            let nIdx = this.widgetsParents.indexOf(oForm);
            if (nIdx != -1) {
                this.widgetsParents.splice(nIdx, oForm);
                this.History.Add(new CChangesPDFDocumentFieldsContent(this, -1, [oForm], false))
            }

            // проверяем родителя этого родителя
            if (oParent) {
                oParent.RemoveKid(oForm);
                this.CheckParentForm(oParent);
            }
        }
    };

    /**
	 * Move page to annot (if annot is't visible)
	 * @memberof CPDFDoc
	 * @typeofeditors ["PDF"]
     * @param {string} sId - id of annot.
     * @param {boolean} bForceMove - move to annot even it's visible.
     * @returns {object}
	 */
    CPDFDoc.prototype.GoToAnnot = function(sId, bForceMove) {
        let oAnnot = this.GetAnnotById(sId);
        if (!oAnnot)
            return;

        let nPage = oAnnot.GetPage();
        let aRect = oAnnot.GetOrigRect();

        let isVisible = false;
        let oViewRect = this.Viewer.getViewingRect2(nPage);

        let nOrigPageW = this.Viewer.file.pages[nPage].W;
        let nOrigPageH = this.Viewer.file.pages[nPage].H;
        let nPageRot = this.Viewer.getPageRotate(nPage);

        let nStartViewX = nOrigPageW * oViewRect.x;
        let nStartViewY = nOrigPageH * oViewRect.y;
        let nEndViewX   = nOrigPageW * oViewRect.r;
        let nEndViewY   = nOrigPageH * oViewRect.b;

        if ((aRect[3] > nStartViewY && aRect[1] < nEndViewY) && (aRect[2] > nStartViewX && aRect[0] < nEndViewX)) {
            isVisible = true;
        }
        
        if (isVisible == true && bForceMove != true)
            return;
        
        // выставляем смещения
        let yOffset;
        let xOffset;

        switch (nPageRot) {
            case 0:
                yOffset = aRect[1];
                xOffset = aRect[0];
                break;
            case 90:
                yOffset = aRect[3];
                xOffset = aRect[0];
                break;
            case 180:
                yOffset = aRect[3];
                xOffset = aRect[2];
                break;
            case 270:
                yOffset = aRect[1];
                xOffset = aRect[2];
                break;
        }

        let oTr = this.pagesTransform[nPage].invert;
        let oPos = oTr.TransformPoint(xOffset, yOffset);

        if (yOffset != undefined && xOffset != undefined || this.Viewer.currentPage != nPage) {
            this.Viewer.disabledPaintOnScroll = true; // вырубаем отрисовку на скроле
            this.Viewer.scrollToXY(this.Viewer.scrollY + oPos.y, this.Viewer.scrollX + oPos.x);
            this.Viewer.disabledPaintOnScroll = false;
            this.Viewer.paint();
        }
    };
    CPDFDoc.prototype.HideComments = function() {
        editor.sync_HideComment();
        this.showedCommentId = undefined;
    };

    CPDFDoc.prototype.GetFieldBySourceIdx = function(nIdx) {
        for (let i = 0; i < this.widgets.length; i++) {
            if (this.widgets[i].GetApIdx() == nIdx) {
                return this.widgets[i];
            }
        }
        for (let i = 0; i < this.widgetsParents.length; i++) {
            if (this.widgetsParents[i].GetApIdx() == nIdx) {
                return this.widgetsParents[i];
            }
        }
    };
    CPDFDoc.prototype.GetAnnotById = function(sId) {
        return this.annots.find(function(annot) {
            return annot.GetId() == sId;
        });
    };
    CPDFDoc.prototype.GetShapeBasedAnnotById = function(sId) {
        if (!sId) {
            return null;
        }

        function findInAnnot(annot) {
            let isFound = false;
            if (annot.GetId() == sId) {
                isFound = true;
                return isFound;
            }

            if (annot.IsAnnot() && annot.IsFreeText()) {
                isFound = annot.spTree.find(findInAnnot)
            }

            return isFound;
        };

        let oDrawing = this.annots.find(function(annot) {
            return findInAnnot(annot);
        });

        return oDrawing;
    };
    CPDFDoc.prototype.GetDrawingById = function(sId) {
        if (!sId) {
            return null;
        }

        function findInDrawing(drawing) {
            let isFound = false;
            if (drawing.GetId() == sId) {
                isFound = true;
                return isFound;
            }

            if (drawing.isGroup()) {
                isFound = drawing.spTree.find(findInDrawing)
            }

            return isFound;
        };

        let oDrawing = this.drawings.find(function(drawing) {
            return findInDrawing(drawing);
        });

        return oDrawing;
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

        this.CheckParentForm(oFieldParent);
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
	 * Resets the forms values.
     * Note: This method used by forms actions.
	 * @memberof CPDFDoc
     * @param {CBaseField[]} aNames - array with forms names to reset. If param is undefined or array is empty then resets all forms.
     * @param {boolean} bAllExcept - reset all fields except aNames
	 * @typeofeditors ["PDF"]
	 */
    CPDFDoc.prototype.ResetForms = function(aNames, bAllExcept) {
        let oActionsQueue = this.GetActionsQueue();
        let oThis = this;

        let aReseted = [];

        if (aNames.length > 0) {
            if (bAllExcept) {
                for (let nField = 0; nField < this.widgets.length; nField++) {
                    let oField = this.widgets[nField];
                    let sFieldName = oField.GetFullName();

                    if (aNames.includes(sFieldName) == false && false == aReseted.includes(sFieldName)) {
                        oField.Reset();
                        aReseted.push(sFieldName);
                        this.AddFieldToCommit(oField);
                    }
                }
            }
            else {
                aNames.forEach(function(name) {
                    let oField = oThis.GetField(name);
                    let sFieldName = oField.GetFullName();

                    if (aNames.includes(sFieldName) == true && false == aReseted.includes(sFieldName)) {
                        oField.Reset();
                        aReseted.push(sFieldName);
                        this.AddFieldToCommit(oField);
                    }
                });
            }
        }
        else {
            this.widgets.forEach(function(field) {
                let sFieldName = field.GetFullName();

                if (false == aReseted.includes(sFieldName)) {
                    field.Reset();
                    aReseted.push(sFieldName);
                    this.AddFieldToCommit(field);
                }
            });
            if (this.widgets.length > 0)
                AscCommon.History.Clear()
        }

        this.CommitFields();
        this.DoCalculateFields();
        
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
                let aFields = oThis.GetAllWidgets(name);
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
        let oController = this.GetController();

        this.annots.forEach(function(annot) {
            annot.SetDisplay(bHidden ? window["AscPDF"].Api.Objects.display["hidden"] : window["AscPDF"].Api.Objects.display["visible"]);
            annot.AddToRedraw();
        });

        this.annotsHidden = bHidden;

        this.HideComments();
        this.mouseDownAnnot = null;
        oController.resetSelection();
        oController.resetTrackState();
    };
    CPDFDoc.prototype.IsAnnotsHidden = function() {
        return this.annotsHidden;
    };
    
    /**
	 * Returns array with widgets fields by specified name.
	 * @memberof CPDFDoc
	 * @typeofeditors ["PDF"]
	 * @returns {boolean}
	 */
    CPDFDoc.prototype.GetAllWidgets = function(sName) {
        let aFields = [];
        for (let i = 0; i < this.widgets.length; i++) {
            if (this.widgets[i].GetFullName() == sName)
                aFields.push(this.widgets[i]);
        }

        if (aFields.length == 0) {
            for (let i = 0; i < this.widgetsParents.length; i++) {
                if (this.widgetsParents[i].GetFullName() == sName) {
                    aFields = aFields.concat(this.widgetsParents[i].GetAllWidgets());
                    break;
                }
            }
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
        for (let i = 0; i < this.widgetsParents.length; i++) {
            if (this.widgetsParents[i].GetFullName() == sName) {
                return this.widgetsParents[i];
            }
        }

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
	// Work with interface
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CPDFDoc.prototype.UpdateInterface = function() {
        this.Api.sync_BeginCatchSelectedElements();

        let oDrDoc      = this.GetDrawingDocument();
        let oController = this.GetController();
        let oDrawingPr  = oController.getDrawingProps();
        let oCurObject  = this.GetActiveObject();

        if (oCurObject) {
            if (oCurObject.IsDrawing()) {
                let oImgPr      = oDrawingPr.imageProps;
                let oSpPr       = oDrawingPr.shapeProps;
                let oChartPr    = oDrawingPr.chartProps;
                let oTblPr      = oDrawingPr.tableProps;

                if (oImgPr) {
                    oImgPr.Width = oImgPr.w;
                    oImgPr.Height = oImgPr.h;
                    oImgPr.Position = {X: oImgPr.x, Y: oImgPr.y};
                    if (AscFormat.isRealBool(oImgPr.locked) && oImgPr.locked) {
                        oImgPr.Locked = true;
                    }
                    this.Api.sync_ImgPropCallback(oImgPr);
                }
                if (oSpPr) {
                    oSpPr.Position = new Asc.CPosition({X: oSpPr.x, Y: oSpPr.y});
                    this.Api.sync_shapePropCallback(oSpPr);
                    this.Api.sync_VerticalTextAlign(oSpPr.verticalTextAlign);
                    this.Api.sync_Vert(oSpPr.vert);
                }
                if (oTblPr) {
                    oDrDoc.CheckTableStyles(oTblPr.TableLook);
                    this.Api.sync_TblPropCallback(oTblPr);
                    if (!oSpPr) {
                        if (oTblPr.CellsVAlign === vertalignjc_Bottom) {
                            this.Api.sync_VerticalTextAlign(AscFormat.VERTICAL_ANCHOR_TYPE_BOTTOM);
                        } else if (oTblPr.CellsVAlign === vertalignjc_Center) {
                            this.Api.sync_VerticalTextAlign(AscFormat.VERTICAL_ANCHOR_TYPE_CENTER);
                        } else {
                            this.Api.sync_VerticalTextAlign(AscFormat.VERTICAL_ANCHOR_TYPE_TOP);
                        }
                    }
                }
            }
            else if (oCurObject.IsAnnot()) {
                this.Api.sync_annotPropCallback(oCurObject);
            }
        }
        
        let oTargetDocContent = oController.getTargetDocContent(undefined, true);
        let oTargetTextObject = AscFormat.getTargetTextObject(oController);

        this.UpdateUndoRedo();
        this.UpdateCommentPos();
        this.UpdateMathTrackPos();
        this.UpdateAnnotTrackPos();
        this.UpdateSelectionTrackPos();
        this.UpdateCopyCutState();
        this.UpdateParagraphProps();
        this.UpdateTextProps();
        this.UpdateCanAddHyperlinkState();
        if (oTargetTextObject && (!oTargetTextObject.group || !oTargetTextObject.group.IsAnnot())) {
            oTargetDocContent && oTargetDocContent.Document_UpdateInterfaceState();
        }
        this.Api.sync_EndCatchSelectedElements();
        
        Asc.editor.CheckChangedDocument();
    };

    //-----------------------------------------------------------------------------------
    // Функции для работы с гиперссылками
    //-----------------------------------------------------------------------------------
    CPDFDoc.prototype.AddHyperlink = function (HyperProps) {
        let oController = this.GetController();
        oController.checkSelectedObjectsAndCallback(oController.hyperlinkAdd, [HyperProps], false, AscDFH.historydescription_Presentation_HyperlinkAdd);
    };
    CPDFDoc.prototype.ModifyHyperlink = function (HyperProps) {
        let oController = this.GetController();
        oController.checkSelectedObjectsAndCallback(oController.hyperlinkModify, [HyperProps], false, AscDFH.historydescription_Presentation_HyperlinkModify);
    };
    CPDFDoc.prototype.RemoveHyperlink = function () {
        let oController = this.GetController();
        oController.checkSelectedObjectsAndCallback(oController.hyperlinkRemove, [], false, AscDFH.historydescription_Presentation_HyperlinkRemove);
    };
    CPDFDoc.prototype.UpdateCanAddHyperlinkState = function() {
        this.Api.sync_CanAddHyperlinkCallback(this.CanAddHyperlink(false));
    };
    CPDFDoc.prototype.CanAddHyperlink = function(bCheckInHyperlink) {
        let oController = this.GetController();
        return oController.hyperlinkCanAdd(bCheckInHyperlink);
    };
    //-----------------------------------------------------------------------------------

    CPDFDoc.prototype.UpdateUndoRedo = function() {
        let bCanUndo = this.History.Can_Undo() || this.LocalHistory.Can_Undo();
        let bCanRedo = this.History.Can_Redo() || this.LocalHistory.Can_Redo();

        if (true !== bCanUndo && this.Api && this.CollaborativeEditing && true === this.CollaborativeEditing.Is_Fast() && true !== this.CollaborativeEditing.Is_SingleUser())
			bCanUndo = this.CollaborativeEditing.CanUndo();
        
		Asc.editor.sync_CanUndoCallback(bCanUndo);
		Asc.editor.sync_CanRedoCallback(bCanRedo);
    };
    CPDFDoc.prototype.UpdateCopyCutState = function() {
        let oCanCopyCut = this.CanCopyCut();
        editor.sync_CanCopyCutCallback(oCanCopyCut.copy, oCanCopyCut.cut);
    };
    CPDFDoc.prototype.CanCopyCut = function() {
        let oViewer         = editor.getDocumentRenderer();
        let oActiveForm     = this.activeForm;
        let oActiveAnnot    = this.mouseDownAnnot;
        let oActiveDrawing  = this.activeDrawing;

        let isCanCopy = false;
        let isCanCut = false;

        let oSelection = oViewer.file.Selection;
        if (oSelection.Glyph1 != oSelection.Glyph2 || oSelection.Line1 != oSelection.Line2 || oSelection.Page1 != oSelection.Page2) {
            isCanCopy = true;
        }
        
        let oContent;
        if (oActiveForm) {
            oContent = oActiveForm.GetDocContent();
        }
        else if (oActiveAnnot && oActiveAnnot.IsFreeText() && oActiveAnnot.IsInTextBox()) {
            oContent = oActiveAnnot.GetDocContent();
        }
        else if (oActiveDrawing) {
            oContent = oActiveDrawing.GetDocContent();
            if (!oActiveDrawing.IsInTextBox()) {
                isCanCopy = true;
                isCanCut = true;
            }
        }

        if (oContent && oContent.IsSelectionUse() && !oContent.IsSelectionEmpty()) {
            isCanCopy = true;
            isCanCut = true;
        }

        return {
            copy: isCanCopy,
            cut: isCanCut
        }
    };
    CPDFDoc.prototype.UpdateParagraphProps = function() {
        let oParaPr = this.GetCalculatedParaPr();

        if (oParaPr) {
            let isCanIncreaseInd = this.CanIncreaseParagraphLevel(true);
            let isCanDecreaseInd = this.CanIncreaseParagraphLevel(false);
            Asc.editor.sendEvent("asc_canIncreaseIndent", isCanIncreaseInd);
            Asc.editor.sendEvent("asc_canDecreaseIndent", isCanDecreaseInd);
            
            Asc.editor.UpdateParagraphProp(oParaPr);
            Asc.editor.sync_PrPropCallback(oParaPr);
        }
    };
    CPDFDoc.prototype.UpdateTextProps = function() {
        let oTextPr = this.GetCalculatedTextPr();
        if (oTextPr) {
            Asc.editor.UpdateTextPr(oTextPr);
        }
    };
    CPDFDoc.prototype.CanIncreaseParagraphLevel = function(bIncrease) {
        let oController = this.GetController();
        return oController.canIncreaseParagraphLevel(bIncrease);
    };

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Work with text
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    CPDFDoc.prototype.MoveCursorLeft = function(isShiftKey, isCtrlKey) {
        let oDrDoc = this.GetDrawingDocument();

        let oForm       = this.activeForm;
        let oFreeText   = this.mouseDownAnnot && this.mouseDownAnnot.IsFreeText() ? this.mouseDownAnnot : null;
        let oDrawing    = this.activeDrawing;
        let oController = this.GetController();

        oDrDoc.UpdateTargetFromPaint = true;

        let oContent;
        if (oForm && oForm.IsInForm() && [AscPDF.FIELD_TYPES.text, AscPDF.FIELD_TYPES.combobox].includes(oForm.GetType())) {
            oForm.MoveCursorLeft(isShiftKey, isCtrlKey);
            oContent = oForm.GetDocContent();
        }
        else if (oFreeText && oFreeText.IsInTextBox()) {
            oFreeText.MoveCursorLeft(isShiftKey, isCtrlKey);
            oContent = oFreeText.GetDocContent();
        }
        else if (oDrawing && oDrawing.IsInTextBox()) {
            oController.cursorMoveLeft(isShiftKey, isCtrlKey);
        }

        if (oContent) {
            oDrDoc.TargetStart(true);
            
            if (oContent.IsSelectionUse() && false == oContent.IsSelectionEmpty())
                oDrDoc.TargetEnd();

            this.Viewer.onUpdateOverlay();
        }
    };
    CPDFDoc.prototype.MoveCursorUp = function(isShiftKey, isCtrlKey) {
        let oDrDoc = this.GetDrawingDocument();

        let oForm       = this.activeForm;
        let oFreeText   = this.mouseDownAnnot && this.mouseDownAnnot.IsFreeText() ? this.mouseDownAnnot : null;
        let oDrawing    = this.activeDrawing;
        let oController = this.GetController();

        oDrDoc.UpdateTargetFromPaint = true;

        let oContent;
        if (oForm && !oForm.IsNeedDrawHighlight())
        {
            switch (oForm.GetType())
            {
                case AscPDF.FIELD_TYPES.listbox:
                    oForm.MoveSelectUp();
                    break;
                case AscPDF.FIELD_TYPES.text: {
                    if (oForm.IsInForm()) {
                        oForm.MoveCursorUp(isShiftKey, isCtrlKey);
                        oContent = oForm.GetDocContent();
                    }
                    break;
                }
            }
        }
        else if (oFreeText && oFreeText.IsInTextBox()) {
            oFreeText.MoveCursorUp(isShiftKey, isCtrlKey);
            oContent = oFreeText.GetDocContent();
        }
        else if (oDrawing && oDrawing.IsInTextBox()) {
            oController.cursorMoveUp(isShiftKey, isCtrlKey);
        }

        if (oContent) {
            oDrDoc.TargetStart(true);

            if (oContent.IsSelectionUse() && false == oContent.IsSelectionEmpty())
                oDrDoc.TargetEnd();

            this.Viewer.onUpdateOverlay();
        }
    };
    CPDFDoc.prototype.MoveCursorRight = function(isShiftKey, isCtrlKey) {
        let oDrDoc = this.GetDrawingDocument();

        let oForm       = this.activeForm;
        let oFreeText   = this.mouseDownAnnot && this.mouseDownAnnot.IsFreeText() ? this.mouseDownAnnot : null;
        let oDrawing    = this.activeDrawing;
        let oController = this.GetController();

        oDrDoc.UpdateTargetFromPaint = true;

        let oContent;
        if (oForm && oForm.IsInForm() && [AscPDF.FIELD_TYPES.text, AscPDF.FIELD_TYPES.combobox].includes(oForm.GetType())) {
            oForm.MoveCursorRight(isShiftKey, isCtrlKey);
            oContent = oForm.GetDocContent();
        }
        else if (oFreeText && oFreeText.IsInTextBox()) {
            oFreeText.MoveCursorRight(isShiftKey, isCtrlKey);
            oContent = oFreeText.GetDocContent();
        }
        else if (oDrawing && oDrawing.IsInTextBox()) {
            oController.cursorMoveRight(isShiftKey, isCtrlKey);
        }

        if (oContent) {
            oDrDoc.TargetStart(true);

            if (oContent.IsSelectionUse() && false == oContent.IsSelectionEmpty())
                oDrDoc.TargetEnd();

            this.Viewer.onUpdateOverlay();
        }
    };
    CPDFDoc.prototype.MoveCursorDown = function(isShiftKey, isCtrlKey) {
        let oDrDoc = this.GetDrawingDocument();

        let oForm           = this.activeForm;
        let oFreeText       = this.mouseDownAnnot && this.mouseDownAnnot.IsFreeText() ? this.mouseDownAnnot : null;
        let oDrawing        = this.activeDrawing;
        let oController     = this.GetController();

        oDrDoc.UpdateTargetFromPaint = true;

        let oContent;
        if (oForm && !oForm.IsNeedDrawHighlight())
        {
            switch (oForm.GetType())
            {
                case AscPDF.FIELD_TYPES.listbox:
                    oForm.MoveSelectDown();
                    break;
                case AscPDF.FIELD_TYPES.text: {
                    if (oForm.IsInForm()) {
                        oForm.MoveCursorDown(isShiftKey, isCtrlKey);
                        oContent = oForm.GetDocContent();
                    }
                    
                    break;
                }
            }
            
        }
        else if (oFreeText && oFreeText.IsInTextBox()) {
            oFreeText.MoveCursorDown(isShiftKey, isCtrlKey);
            oContent = oFreeText.GetDocContent();
        }
        else if (oDrawing && oDrawing.IsInTextBox()) {
            oController.cursorMoveDown(isShiftKey, isCtrlKey);
        }

        if (oContent) {
            oDrDoc.TargetStart(true);

            if (oContent.IsSelectionUse() && false == oContent.IsSelectionEmpty())
                oDrDoc.TargetEnd();

            this.Viewer.onUpdateOverlay();
        }
    };
    CPDFDoc.prototype.SelectAll = function() {
        let oDrDoc      = this.GetDrawingDocument();
        let oController = this.GetController();

        let oForm       = this.activeForm;
        let oFreeText   = this.mouseDownAnnot && this.mouseDownAnnot.IsFreeText() ? this.mouseDownAnnot : null;
        let oDrawing  = this.activeDrawing;

        let oContent;
        if (oForm && oForm.IsInForm() && [AscPDF.FIELD_TYPES.text, AscPDF.FIELD_TYPES.combobox].includes(oForm.GetType())) {
            oForm.SelectAllText();
            oContent = oForm.GetDocContent();
        }
        else if (oFreeText && oFreeText.IsInTextBox()) {
            oFreeText.SelectAllText();
            oContent = oFreeText.GetDocContent();
        }
        else if (oDrawing) {
            oContent = oDrawing.GetDocContent();
            oController.selectAll();

            if (oController.getSelectedArray().length != 0) {
                this.Viewer.onUpdateOverlay();
                return;
            }
        }

        if (oContent) {
            if (false == oContent.IsEmpty()) {
                oDrDoc.TargetEnd();
                this.Viewer.onUpdateOverlay();
            }
            else {
                oContent.RemoveSelection();
            }
        }
        else {
            if (!this.Viewer.isFullTextMessage) {
                if (!this.Viewer.isFullText) {
                    this.Viewer.fullTextMessageCallbackArgs = [];
                    this.Viewer.fullTextMessageCallback = function() {
                        this.selectAll();
                    };
                    this.Viewer.showTextMessage();
                }
                else {
                    this.Viewer.file.selectAll();
                }
            }
        }
    };
    CPDFDoc.prototype.SelectionSetStart = function(x, y, e) {
        let oDrDoc      = this.GetDrawingDocument();
        let oForm       = this.activeForm;
        let oFreeText   = this.mouseDownAnnot && this.mouseDownAnnot.IsFreeText() ? this.mouseDownAnnot : null;
        let oDrawing  = this.activeDrawing;

        // координаты клика на странице в MM
        var pageObject = this.Viewer.getPageByCoords2(x, y);
        if (!pageObject)
            return false;

        let X = pageObject.x;
        let Y = pageObject.y;

        if (oForm && oForm.IsInForm() && [AscPDF.FIELD_TYPES.text, AscPDF.FIELD_TYPES.combobox].includes(oForm.GetType())) {
            oForm.SelectionSetStart(X, Y, e);
            if (false == this.Viewer.isMouseDown) {
                oForm.content.RemoveSelection();
            }
        }
        else if (oFreeText && oFreeText.IsInTextBox()) {
            oFreeText.SelectionSetStart(X, Y, e);
        }
        else if (oDrawing) {
            oDrawing.SelectionSetStart(X, Y, e);
        }
        
        oDrDoc.UpdateTargetFromPaint = true;
        oDrDoc.TargetStart(true);
    };
    CPDFDoc.prototype.SelectionSetEnd = function(x, y, e) {
        let oDrDoc      = this.GetDrawingDocument();
        let oForm       = this.activeForm;
        let oFreeText   = this.mouseDownAnnot && this.mouseDownAnnot.IsFreeText() ? this.mouseDownAnnot : null;
        let oDrawing  = this.activeDrawing;

        // координаты клика на странице в MM
        var pageObject = this.Viewer.getPageByCoords2(x, y);
        if (!pageObject)
            return false;

        let X = pageObject.x;
        let Y = pageObject.y;

        let oContent;
        if (oForm && oForm.IsInForm() && [AscPDF.FIELD_TYPES.text, AscPDF.FIELD_TYPES.combobox].includes(oForm.GetType())) {
            oForm.SelectionSetEnd(X, Y, e);
            oContent = oForm.GetDocContent();
        }
        else if (oFreeText && oFreeText.IsInTextBox()) {
            oFreeText.SelectionSetEnd(X, Y, e);
            oContent = oFreeText.GetDocContent();
        }
        else if (oDrawing) {
            oDrawing.SelectionSetEnd(X, Y, e);
            oContent = oDrawing.GetDocContent();
        }

        if (oContent) {
            if (oContent.IsSelectionEmpty() == false) {
                oDrDoc.TargetEnd();
            }
            else {
                oDrDoc.TargetStart(true);
            }
        }
    };
    CPDFDoc.prototype.GetCalculatedParaPr = function(bInitIfNull) {
        let oController = this.GetController();
        let oParaPr     = oController.getParagraphParaPr();

        if (oParaPr) {
            return oParaPr;
        }

        if (bInitIfNull) {
            return new AscWord.CParaPr();
        }
    };
    CPDFDoc.prototype.GetCalculatedTextPr = function (bInitIfNull) {
        let oController = this.GetController();
        let oTextPr     = oController.getParagraphTextPr();

        if (oTextPr) {
            let oTheme = oController.getTheme();
            if (oTheme) {
                oTextPr.ReplaceThemeFonts(oTheme.themeElements.fontScheme);
            }
            return oTextPr;
        }

        if (bInitIfNull) {
            return new AscWord.CTextPr();
        }
    };
    CPDFDoc.prototype.Get_GraphicObjectsProps = function () {
        return this.GetController().getDrawingProps();
    };
    CPDFDoc.prototype.GetDirectTextPr = function() {
        let oController = this.GetController();
        return oController.getParagraphTextPr();
    };
    CPDFDoc.prototype.AddToParagraph = function(oParaItem) {
        this.DoAction(function() {
            let oController = this.GetController();
            let oMathShape  = null;

            let nCurPage = this.Viewer.currentPage;

            let oActiveObj = this.GetActiveObject();
            if (oParaItem.Type === para_Math) {
                if (!oActiveObj || oActiveObj.IsAnnot() || !(oController.selection.textSelection || (oController.selection.groupSelection && oController.selection.groupSelection.selection.textSelection))) {
                    oController.resetSelection();
                    oController.resetTrackState();

                    oMathShape = oController.createTextArt(0, false, null, "");
                    oMathShape.SetDocument(this);
                    oMathShape.SetPage(nCurPage);
                    oMathShape.Recalculate();

                    let oXfrm       = oMathShape.getXfrm();
                    let nRotAngle   = this.Viewer.getPageRotate(nCurPage);

                    let nExtX   = oXfrm.extX;
                    let nExtY   = oXfrm.extY;
                    let oPos    = private_computeDrawingAddingPos(nCurPage, nExtX, nExtY);
                    
                    if (nRotAngle != 0) {
                        oXfrm.setRot(-nRotAngle * Math.PI / 180);
                    }

                    oXfrm.setOffX(oPos.x);
                    oXfrm.setOffY(oPos.y);

                    this.AddDrawing(oMathShape, nCurPage);
                    oMathShape.SetNeedRecalc(true);
                    oMathShape.select(oController, nCurPage);
                    this.SetMouseDownObject(oMathShape);
                    oController.selection.textSelection = oMathShape;
                }
            }

            oController.paragraphAdd(oParaItem, false);
        }, AscDFH.historydescription_Presentation_ParagraphAdd, this);
    };
    CPDFDoc.prototype.AddNewParagraph = function() {
        let oController = this.GetController();
        let oActiveObj  = this.GetActiveObject();
        
        oController.checkSelectedObjectsAndCallback(oController.addNewParagraph, [], false, AscDFH.historydescription_Presentation_AddNewParagraph);

        if (oActiveObj.IsAnnot() && oActiveObj.IsFreeText()) {
            oActiveObj.FitTextBox();
        }

        this.FinalizeAction();
    };
    CPDFDoc.prototype.GetSelectedText = function(bClearText, oPr) {
        let oForm       = this.activeForm;
        let oController = this.GetController();

        if (oForm) {
            let oContent = oForm.GetDocContent();
            if (oContent) {
                return oContent.GetSelectedText(bClearText, oPr);
            }
        }
        else {
            return oController.GetSelectedText(bClearText, oPr);
        }
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
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Text/Para Pr
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    CPDFDoc.prototype.SetHighlight = function(r, g, b, opacity) {
        this.HighlightColor = {
            r: r,
            g: g,
            b: b,
            a: opacity
        };

        let oDrawing        = this.activeDrawing;
        let oViewer         = editor.getDocumentRenderer();
        let oFile           = oViewer.file;
        let aSelQuads       = null == oDrawing ? oFile.getSelectionQuads() : oDrawing.GetSelectionQuads();

        if (oDrawing && false == this.Api.IsCommentMarker()) {
            this.SetParagraphHighlight(AscCommon.isNumber(r) && AscCommon.isNumber(g) && AscCommon.isNumber(b), r, g, b);
            return;
        }
        else {
            if (aSelQuads.length == 0) {
                return;
            }

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
                    rect:           [MinX - 3, MinY - 1, MaxX + 3, MaxY + 1],
                    page:           nPage,
                    name:           AscCommon.CreateGUID(),
                    type:           AscPDF.ANNOTATIONS_TYPES.Highlight,
                    creationDate:   (new Date().getTime()).toString(),
                    modDate:        (new Date().getTime()).toString(),
                    hidden:         false
                }

                let oAnnot = this.AddAnnotByProps(oProps);

                oAnnot.SetQuads(aQuads);
                oAnnot.SetStrokeColor([private_correctRGBColorComponent(r)/255, private_correctRGBColorComponent(g)/255, private_correctRGBColorComponent(b)/255]);
                oAnnot.SetOpacity(typeof(opacity) != "boolean" ? opacity / 100 : 1);
            }
        }

        if (this.bOffMarkerAfterUsing) {
            editor.sendEvent("asc_onMarkerFormatChanged", AscPDF.ANNOTATIONS_TYPES.Highlight, false);
            editor.SetMarkerFormat(AscPDF.ANNOTATIONS_TYPES.Highlight, false);
        }
    };
    CPDFDoc.prototype.SetParagraphHighlight = function(IsColor, r, g, b) {
        let oController = this.GetController();
        
        let oDoc = this;
        let oTargetContent = oController.getTargetDocContent();
		if (!oTargetContent || oTargetContent.IsSelectionUse() && !oTargetContent.IsSelectionEmpty()) {
			oController.checkSelectedObjectsAndCallback(function () {
				if (false === IsColor) {
					oDoc.AddToParagraph(new ParaTextPr({HighlightColor: null}));
				} else {
					oDoc.AddToParagraph(new ParaTextPr({HighlightColor: AscFormat.CreateUniColorRGB(r, g, b)}));
				}
			}, [], false, AscDFH.historydescription_Document_SetTextHighlight);
		}
    };
    CPDFDoc.prototype.SetUnderline = function(r, g, b, opacity) {
        this.UnderlineColor = {
            r: r,
            g: g,
            b: b,
            a: opacity
        };

        let oDrawing        = this.activeDrawing;
        let oViewer         = editor.getDocumentRenderer();
        let oFile           = oViewer.file;
        let aSelQuads       = null == oDrawing ? oFile.getSelectionQuads() : oDrawing.GetSelectionQuads();
        
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
                rect:           [MinX - 3, MinY - 1, MaxX + 3, MaxY + 1],
                page:           nPage,
                name:           AscCommon.CreateGUID(),
                type:           AscPDF.ANNOTATIONS_TYPES.Underline,
                creationDate:   (new Date().getTime()).toString(),
                modDate:        (new Date().getTime()).toString(),
                hidden:         false
            }

            let oAnnot = this.AddAnnotByProps(oProps);

            oAnnot.SetQuads(aQuads);
            oAnnot.SetStrokeColor([private_correctRGBColorComponent(r)/255, private_correctRGBColorComponent(g)/255, private_correctRGBColorComponent(b)/255]);
            oAnnot.SetOpacity(opacity / 100);
        }

        if (this.bOffMarkerAfterUsing) {
            editor.sendEvent("asc_onMarkerFormatChanged", AscPDF.ANNOTATIONS_TYPES.Underline, false);
            editor.SetMarkerFormat(AscPDF.ANNOTATIONS_TYPES.Underline, false);
        }
    };
    CPDFDoc.prototype.SetStrikeout = function(r, g, b, opacity) {
        this.StrikeoutColor = {
            r: r,
            g: g,
            b: b,
            a: opacity
        };

        let oDrawing        = this.activeDrawing;
        let oViewer         = editor.getDocumentRenderer();
        let oFile           = oViewer.file;
        let aSelQuads       = null == oDrawing ? oFile.getSelectionQuads() : oDrawing.GetSelectionQuads();

        if (aSelQuads.length == 0) return;

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
                rect:           [MinX - 3, MinY - 1, MaxX + 3, MaxY + 1],
                page:           nPage,
                name:           AscCommon.CreateGUID(),
                type:           AscPDF.ANNOTATIONS_TYPES.Strikeout,
                creationDate:   (new Date().getTime()).toString(),
                modDate:        (new Date().getTime()).toString(),
                hidden:         false
            }

            let oAnnot = this.AddAnnotByProps(oProps);

            oAnnot.SetQuads(aQuads);
            oAnnot.SetStrokeColor([private_correctRGBColorComponent(r)/255, private_correctRGBColorComponent(g)/255, private_correctRGBColorComponent(b)/255]);
            oAnnot.SetOpacity(opacity / 100);
        }

        if (this.bOffMarkerAfterUsing) {
            editor.sendEvent("asc_onMarkerFormatChanged", AscPDF.ANNOTATIONS_TYPES.Strikeout, false);
            editor.SetMarkerFormat(AscPDF.ANNOTATIONS_TYPES.Strikeout, false);
        }
    };
    CPDFDoc.prototype.SetParagraphSpacing = function(oSpacing) {
        let oController = this.GetController();
        oController.checkSelectedObjectsAndCallback(oController.setParagraphSpacing, [oSpacing], false, AscDFH.historydescription_Presentation_SetParagraphSpacing);
    };
    CPDFDoc.prototype.SetParagraphNumbering = function(oBullet) {
        let oController = this.GetController();
        oController.checkSelectedObjectsAndCallback(oController.setParagraphNumbering, [oBullet], false, AscDFH.historydescription_Presentation_SetParagraphNumbering);
    };
    CPDFDoc.prototype.IncreaseDecreaseFontSize = function(bIncrease) {
        let oController = this.GetController();
        
        oController.checkSelectedObjectsAndCallback(function () {
            oController.paragraphIncDecFontSize(bIncrease);
        }, [], false, AscDFH.historydescription_Presentation_ParagraphIncDecFontSize);
    };
    CPDFDoc.prototype.ChangeTextCase = function(nType) {
        let oController = this.GetController();
        let oActiveObj  = this.GetActiveObject();
        
        oController.changeTextCase(nType);
    };
    CPDFDoc.prototype.SetParagraphAlign = function(Align) {
        let oController = this.GetController();
        oController.checkSelectedObjectsAndCallback(oController.setParagraphAlign, [Align], false, AscDFH.historydescription_Presentation_SetParagraphAlign);
    };
    CPDFDoc.prototype.SetVerticalAlign = function(Align) {
        let oController = this.GetController();
        oController.checkSelectedObjectsAndCallback(oController.applyDrawingProps, [{verticalTextAlign: Align}], false, AscDFH.historydescription_Presentation_SetVerticalAlign);
    };
    CPDFDoc.prototype.IncreaseDecreaseIndent = function(bIncrease) {
        let oController = this.GetController();
        oController.checkSelectedObjectsAndCallback(oController.paragraphIncDecIndent, [bIncrease], false, AscDFH.historydescription_Presentation_ParagraphIncDecIndent);
    };
    CPDFDoc.prototype.ClearParagraphFormatting = function(isClearParaPr, isClearTextPr) {
        let oController = this.GetController();
        oController.checkSelectedObjectsAndCallback(oController.paragraphClearFormatting, [isClearParaPr, isClearTextPr], false, AscDFH.historydescription_Presentation_ParagraphClearFormatting);
    };
        
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// For drawings
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    CPDFDoc.prototype.EditPage = function(nPage) {
        if (null == this.Viewer.pagesInfo.pages[nPage] || this.Viewer.file.pages[nPage].isRecognized) {
            return;
        }

        let oFile = this.Viewer.file;
        let nOriginIndex = oFile.pages[nPage].originIndex;
        if (nOriginIndex == undefined) {
            return;
        }

        this.StartAction(AscDFH.historydescription_Pdf_EditPage);
        if (this.IsSelectionLocked(AscDFH.historydescription_Pdf_EditPage, [nPage])) {
            this.FinalizeAction(true);
            return;
        }

        Asc.editor.canSave = false;
        let oDrDoc = this.GetDrawingDocument();

        let oPageInfo = this.GetPageInfo(nPage);
        oPageInfo.SetRecognized(true);
        
        let aSpsXmls        = oFile.nativeFile["scanPage"](nOriginIndex, 1);
        let oParserContext  = new AscCommon.XmlParserContext();
        let oTableStyles    = this.GetTableStyles();
        Object.keys(this.TableStylesIdMap).forEach(function(styleId) {
            oParserContext.addTableStyle(oTableStyles.Get(styleId).GetStyleId(), oTableStyles.Get(styleId));
        });
        let aPageDrawings   = [];
        let oXmlReader;
        
        oParserContext.DrawingDocument = oDrDoc;

        for (let i = 0; i < aSpsXmls.length; i++) {
            oXmlReader = new AscCommon.StaxParser(aSpsXmls[i], undefined, oParserContext);
            oXmlReader.parseNode(0);

            let _t = this;
            let oDrawing;
            oXmlReader.rels = {
                getRelationship : function(rId) {
                    let url =  _t.Viewer.file.nativeFile["getImageBase64"](parseInt(rId.substring(3)));
                    if ("data:" === url.substring(0, 5)) {
                        return {
                            targetMode : "InternalBase64",
                            base64 : url,
                            drawing: oDrawing
                        }
                    } else {
                        return {
                            targetMode: "InternalLoaded",
                            targetFullName: url,
                            drawing: oDrawing
                        }
                    }
                }
            };

            switch (oXmlReader.GetName()) {
                case 'p:sp': {
                    oDrawing = new AscPDF.CPdfShape();
                    break;
                }
                case 'p:graphicFrame': {
                    oDrawing = new AscPDF.CPdfGraphicFrame();
                    break;
                }
                case 'p:pic': {
                    oDrawing = new AscPDF.CPdfImage();
                    break;
                }
            }
            
            if (oDrawing) {
                oDrawing.fromXml(oXmlReader);
                if (oDrawing.IsShape()) {
                    let new_body_pr = oDrawing.getBodyPr();
                    if (new_body_pr) {
                        new_body_pr = new_body_pr.createDuplicate();
                        new_body_pr.textFit = new AscFormat.CTextFit();
                        new_body_pr.textFit.type = AscFormat.text_fit_Auto;

                        if (oDrawing.txBody) {
                            oDrawing.txBody.setBodyPr(new_body_pr);
                        }
                    }
                }
                
                oDrawing.setBDeleted(false);
                aPageDrawings.push(oDrawing);
                oDrawing.CheckTextOnOpen();
            }
        }

        let _t = this;
        let oImageMap = oParserContext.imageMap;
        let aUrls = [];
        let aBase64Img = []
        for(let sImg in oImageMap) {
            if(oImageMap.hasOwnProperty(sImg)) {
                aUrls.push(sImg);
                let aImg = oImageMap[sImg];
                for(let nIdx = 0; nIdx < aImg.length; ++nIdx) {
                    let oImg = aImg[nIdx];
                    aBase64Img.push(new AscCommon.CBuilderImages(oImg.blipFill, sImg, oImg.drawing, null, null));
                }
            }
        }

        let fEndCallback = function () {
            aPageDrawings.forEach(function(drawing) {
                drawing.SetFromScan(true);
                _t.AddDrawing(drawing, nPage);
                drawing.SetNeedRecalc(true);
            });
            _t.FinalizeAction();
            _t.Viewer.file.removeSelection();
            _t.Viewer.paint(function() {
                _t.Viewer.thumbnails._repaintPage(nPage);
            });
            Asc.editor.canSave = true;
        };
        if(aUrls.length > 0) {
            AscCommon.sendImgUrls(Asc.editor, aUrls, function (data) {
                let oObjectsForDownload = AscCommon.GetObjectsForImageDownload(aBase64Img);
                AscCommon.ResetNewUrls(data, aUrls, oObjectsForDownload.aBuilderImagesByUrl, oImageMap);
                let aLoadUrls = [];
                for(let nIdx = 0; nIdx < data.length; ++nIdx) {
                    if(data[nIdx].url) {
                        aLoadUrls.push(data[nIdx].url);
                    }
                }
                Asc.editor.ImageLoader.LoadImagesWithCallback(aLoadUrls, fEndCallback, []);

                let _file = _t.Viewer.file;
                for (let url in aUrls) {
                    _file.nativeFile["changeImageUrl"](aUrls[url], oImageMap[url]);
                }
            });
        }
        else {
            fEndCallback();
        }
    };
    CPDFDoc.prototype.InsertContent2 = function(aSelContent, nIndex) {
        if (true == this.Api.isRestrictionView()) {
            return false;
        }

        let oThis = this;
        return oThis.InsertContent(aSelContent[nIndex].copy());
    };
    CPDFDoc.prototype.InsertContent = function(oSelContent) {
        let oThis       = this;
        let oController = this.GetController();
        let nCurPage    = this.GetCurPage();

        let bResult = false;

        // во view шейпы не вставляем
        if (true == this.Api.isRestrictionView()) {
            return bResult;
        }

        if (oSelContent.Drawings.length != 0) {
            this.BlurActiveObject();

            let aDrToPaste = oSelContent.Drawings.map(function(pasteObj) {
                return pasteObj.Drawing;
            });

            aDrToPaste.forEach(function(drawing, index) {
                let oXfrm = drawing.getXfrm();
                let oPos = private_computeDrawingAddingPos(nCurPage, oXfrm.extX, oXfrm.extY);
                oXfrm.setOffX(oPos.x);
                oXfrm.setOffY(oPos.y);

                // чуть-чуть смещаем при вставке, чтобы было видно вставленную фигуру
                let nShift = oController.getDrawingsPasteShift([drawing]);

                if (nShift > 0) {
                    oXfrm.shift(nShift, nShift);
                }

                oThis.AddDrawing(drawing, oThis.GetCurPage());

                if (drawing.IsGraphicFrame()) {
                    oController.Check_GraphicFrameRowHeight(drawing);
                }
                
                if (index == 0) {
                    oThis.SetMouseDownObject(drawing);
                }
                drawing.select(oController, nCurPage);
            });

            bResult = true;
        }
        if (oSelContent.DocContent) {
            oSelContent.DocContent.EndCollect(this);
            if (oSelContent.DocContent.Elements.length > 0) {
                let oTargetTextObject = AscFormat.getTargetTextObject(oController);
                let oTargetDocContent = oController.getTargetDocContent(true), paragraph, NearPos;

                if (oTargetDocContent) {
                    if (oTargetDocContent.Selection.Use) {
                        oController.removeCallback(1, undefined, undefined, undefined, undefined, undefined);
                    }

                    paragraph = oTargetDocContent.Content[oTargetDocContent.CurPos.ContentPos];
                    if (null != paragraph && paragraph.IsParagraph()) {
                        NearPos = {Paragraph: paragraph, ContentPos: paragraph.Get_ParaContentPos(false, false)};
                        paragraph.Check_NearestPos(NearPos);
                        oSelContent.DocContent.Insert(NearPos);
                    }
                    
                    oTargetTextObject && oTargetTextObject.checkExtentsByDocContent && oTargetTextObject.checkExtentsByDocContent();
                    oTargetTextObject.SetNeedRecalc(true);
                    AscCommon.History.SetSourceObjectsToPointPdf([oTargetTextObject]);
                }
                else {
                    this.CreateAndAddShapeFromSelectedContent(oSelContent.DocContent);
                }
            }

            bResult = true;
        }

        return bResult;
    };
    CPDFDoc.prototype.CreateAndAddShapeFromSelectedContent = function (oDocContent) {
        let nPage   = this.GetCurPage();
        let oTrack = new AscFormat.NewShapeTrack("textRect", 0, 0, this.GetTheme(), undefined, undefined, this, nPage);
        oTrack.track({}, 0, 0);

        let oShape = oTrack.getShape(false, this.GetDrawingDocument(), this.GetController());
        oShape.spPr.setLn(null);
        this.AddDrawing(oShape, nPage);

        oDocContent.ReplaceContent(oShape.txBody.content);

        let nPageW = this.GetPageWidthMM(nPage);
        let nPageH = this.GetPageHeightMM(nPage);
        let oBodyPr = oShape.getBodyPr();

        let nExtX = oShape.txBody.getMaxContentWidth(nPageW / 2, true) + oBodyPr.lIns + oBodyPr.rIns;
        let nExtY = oShape.txBody.content.GetSummaryHeight() + oBodyPr.tIns + oBodyPr.bIns;

        let oPos = private_computeDrawingAddingPos(nPage, nExtX, nExtY);

        oShape.spPr.xfrm.setOffX(oPos.x);
        oShape.spPr.xfrm.setOffY(oPos.y);

        oShape.spPr.xfrm.setExtX(nExtX);
        oShape.spPr.xfrm.setExtY(nExtY);

        return oShape;
    };
    CPDFDoc.prototype.AddDrawing = function(oDrawing, nPage, nPosInTree) {
        let oPagesInfo = this.GetPageInfo(nPage);
        if (!oPagesInfo)
            return;

        this.drawings.push(oDrawing);

        oDrawing.SetDocument(this);
        oPagesInfo.AddDrawing(oDrawing, nPosInTree);
        this.ClearSearch();
    };
    CPDFDoc.prototype.AddTextArt = function(nStyle, nPage) {
        let oPagesInfo = this.GetPageInfo(nPage);
        if (!oPagesInfo)
            return;

        let oController = this.GetController();

        let oTextArt = this.GetController().createTextArt(nStyle, false);
        oTextArt.SetDocument(this);
        oTextArt.setParent(oPagesInfo);
        oTextArt.SetPage(nPage);
        oTextArt.Recalculate();
        oTextArt.checkExtentsByDocContent();
        
        let oXfrm       = oTextArt.getXfrm();
        let nRotAngle    = this.Viewer.getPageRotate(nPage);

        let nExtX   = oXfrm.extX;
        let nExtY   = oXfrm.extY;
        let oPos    = private_computeDrawingAddingPos(nPage, nExtX, nExtY);

        if (nRotAngle != 0) {
            oXfrm.setRot(-nRotAngle * Math.PI / 180);
        }

        oXfrm.setOffX(oPos.x);
        oXfrm.setOffY(oPos.y);

        this.drawings.push(oTextArt);
        oPagesInfo.AddDrawing(oTextArt);
        oTextArt.SetNeedRecalc(true);
        
        this.SetMouseDownObject(oTextArt);
        oTextArt.select(oController, nPage);
        this.SetMouseDownObject(oTextArt);
        oController.selection.textSelection = oTextArt;
        oController.selectAll();
    };
    CPDFDoc.prototype.AddSmartArt = function(nSmartArtType, oPlaceholder, nPage) {
        let oPagesInfo = this.Viewer.pagesInfo;
        if (!oPagesInfo.pages[nPage])
            return;

        let nPageW      = this.GetPageWidthMM(nPage);
        let nPageH      = this.GetPageHeightMM(nPage);
        let nRotAngle   = this.Viewer.getPageRotate(nPage);

        let nExtX   = nPageW * 2 /3;
        let nExtY   = nPageH / 5;

        let oController = this.GetController();
        let oSmartArt   = new AscPDF.CPdfSmartArt();

        oSmartArt.fillByPreset(nSmartArtType);
        oSmartArt.fitForSizes(nExtY, nExtX);
        oSmartArt.fitFontSize();
        oSmartArt.recalculateBounds();

        // oSmartArt.changeSize(nExtX / oSmartArt.extX, nExtY / oSmartArt.extY);
        let oXfrm   = oSmartArt.getXfrm();
        let oPos    = private_computeDrawingAddingPos(nPage, nExtX, nExtY);

        if (nRotAngle != 0) {
            oXfrm.setRot(-nRotAngle * Math.PI / 180);
        }
        oXfrm.setOffX(oPos.x);
        oXfrm.setOffY(oPos.y);

        oSmartArt.normalize();
        oSmartArt.setRecalculateInfo();
		
        let oPh;
        if (oPlaceholder) {
			this.Api.WordControl.m_bIsMouseLock = false;
			oPh = AscCommon.g_oTableId.Get_ById(oPlaceholder.id);
			if (oPh) {
				const nWidth = oPh.extX;
                const nHeight = oPh.extY;
                oSmartArt.fitForSizes(nHeight, nWidth);
                const nX = oPh.x + oPh.extX / 2 - oSmartArt.spPr.xfrm.extX / 2;
                const nY = oPh.y + oPlaceholder.extY / 2 - oSmartArt.spPr.xfrm.extY / 2;
                oSmartArt.spPr.xfrm.setOffX(nX);
                oSmartArt.spPr.xfrm.setOffY(nY);
			}
		}
		
		oSmartArt.checkDrawingBaseCoords();
		oSmartArt.fitFontSize();
		oController.checkChartTextSelection();
		oController.resetSelection();
        oController.resetTrackState();
		oSmartArt.select(oController, 0);
		this.SetMouseDownObject(oSmartArt);

        oController.clearTrackObjects();
        oController.clearPreTrackObjects();
        oController.changeCurrentState(new AscFormat.NullState(oController));
        oController.updateSelectionState();

        this.AddDrawing(oSmartArt, nPage);
        return oSmartArt;
    };
    CPDFDoc.prototype.AddChartByBinary = function(chartBinary, isFromInterface, oPlaceholder, nPage) {
        let oPagesInfo = this.Viewer.pagesInfo;
        if (!oPagesInfo.pages[nPage])
            return;

        let oThis       = this;
        let oController = this.GetController();
        let nRotAngle    = this.Viewer.getPageRotate(nPage);

        let oChart  = oController.getChartSpace2(chartBinary, null);
        let oXfrm   = oChart.getXfrm();

        let nExtX   = oXfrm.extX;
        let nExtY   = oXfrm.extY;
        let oPos    = private_computeDrawingAddingPos(nPage, nExtX, nExtY);

        if (oPlaceholder) {
            let oPh = AscCommon.g_oTableId.Get_ById(oPlaceholder.id);

            if (oPh) {
                oPos.x = oPh.x;
                oPos.y = oPh.y;
                oXfrm.setExtX(oPh.extX);
                oXfrm.setExtY(oPh.extY);
            }
            else {
                return;
            }
        }

        if (nRotAngle != 0) {
            oXfrm.setRot(-nRotAngle * Math.PI / 180);
        }
        oXfrm.setOffX(oPos.x);
        oXfrm.setOffY(oPos.y);

        oController.resetSelection();
        oController.resetTrackState();
        oController.selectObject(oChart, 0);

        if (isFromInterface) {
            AscFonts.FontPickerByCharacter.checkText("", this, function () {
                oThis.AddDrawing(oChart, nPage);
            }, false, false, false);
        }
        else {
            this.AddDrawing(oChart, nPage);
        }
    };
    CPDFDoc.prototype.AddTable = function(nCol, nRow, sStyleId, nPage) {
        let oPagesInfo  = this.Viewer.pagesInfo;
        let oController = this.GetController();
        if (!oPagesInfo.pages[nPage])
            return;

        let oGrFrame = this.private_Create_TableGraphicFrame(nCol, nRow, sStyleId || this.DefaultTableStyleId, undefined, undefined, undefined, undefined, nPage);
        
        this.AddDrawing(oGrFrame, nPage);
        oController.Check_GraphicFrameRowHeight(oGrFrame);
        this.SetMouseDownObject(oGrFrame);
        oGrFrame.select(this.GetController(), nPage);
    };
    CPDFDoc.prototype.private_Create_TableGraphicFrame = function(Cols, Rows, StyleId, Width, Height, PosX, PosY, nPage, bInline) {
        let nPageW      = this.GetPageWidthMM(nPage);
        let nRotAngle    = this.Viewer.getPageRotate(nPage);

        if (false == AscFormat.isRealNumber(Width)) {
            Width = nPageW * 2 / 3;
        }

        let Grid = [];
    
        for (let Index = 0; Index < Cols; Index++)
            Grid[Index] = Width / Cols;
    
        let RowHeight = AscFormat.isRealNumber(Height) ? Height / Rows : 7.478268771701388;

        let X, Y;
        if (AscFormat.isRealNumber(PosX) && AscFormat.isRealNumber(PosY)) {
            X = PosX;
            Y = PosY;
        } else {
            let nExtX   = Width;
            let nExtY   = RowHeight * Rows;
            let oPos    = private_computeDrawingAddingPos(nPage, nExtX, nExtY);
            
            X = oPos.x;
            Y = oPos.y;
        }
        
        let Inline = false;
        if (AscFormat.isRealBool(bInline)) {
            Inline = bInline;
	    }

        let graphic_frame = new AscPDF.CPdfGraphicFrame();
        graphic_frame.setSpPr(new AscFormat.CSpPr());
        graphic_frame.spPr.setParent(graphic_frame);
        graphic_frame.spPr.setXfrm(new AscFormat.CXfrm());

        let oXfrm = graphic_frame.getXfrm();
        oXfrm.setParent(graphic_frame.spPr);
        if (nRotAngle != 0) {
            oXfrm.setRot(-nRotAngle * Math.PI / 180);
        }

        oXfrm.setOffX(X);
        oXfrm.setOffY(Y);
        oXfrm.setExtX(Width);
        oXfrm.setExtY(RowHeight * Rows);
        graphic_frame.setNvSpPr(new AscFormat.UniNvPr());
    
        let table = new CTable(this.GetDrawingDocument(), graphic_frame, Inline, Rows, Cols, Grid, true);
        table.Reset(Inline ? X : 0, Inline ? Y : 0, Width, 100000, 0, 0, 1, 0);
        if (!Inline) {
            table.Set_PositionH(Asc.c_oAscHAnchor.Page, false, 0);
            table.Set_PositionV(Asc.c_oAscVAnchor.Page, false, 0);
        }
        table.SetTableLayout(tbllayout_Fixed);
        if (typeof StyleId === "string") {
            table.Set_TableStyle(StyleId);
        }
        table.Set_TableLook(new AscCommon.CTableLook(false, true, false, false, true, false));
        for (let i = 0; i < table.Content.length; ++i) {
            let Row = table.Content[i];
            if (AscFormat.isRealNumber(RowHeight)) {
                Row.Set_Height(RowHeight, Asc.linerule_AtLeast);
            }
        }
        graphic_frame.setGraphicObject(table);
        graphic_frame.setBDeleted(false);
        
        return graphic_frame;
    };
    CPDFDoc.prototype.AddFreeTextAnnot = function(nType, nPage) {
        let oController = this.GetController();
        let nRotAngle   = this.Viewer.getPageRotate(nPage);
        let oFile       = this.Viewer.file;
        let oViewRect   = this.Viewer.getViewingRect(nPage);
        let oNativePage = oFile.pages[nPage];
        let nPageW      = oNativePage.W;
        let nPageH      = oNativePage.H;
        let oUser       = Asc.editor.User;

        let nExtX = 200;
        let nExtY = 85;

        let X1, Y1, X2, Y2;
        switch (nRotAngle) {
            case 0:
                X1 = nPageW * ((oViewRect.x + oViewRect.r) / 2) - nExtX / 2;
                Y1 = nPageH * ((oViewRect.y + oViewRect.b) / 2) - nExtY / 2;
                X2 = X1 + nExtX;
                Y2 = Y1 + nExtY;
                break;
            case 90:
                X1 = nPageW * ((oViewRect.y + oViewRect.b) / 2) - nExtY / 2;
                Y1 = nPageH - nPageH * ((oViewRect.x + oViewRect.r) / 2) - nExtX / 2;
                X2 = X1 + nExtY;
                Y2 = Y1 + nExtX;
                break;
            case 180:
                X1 = nPageW - nPageW * ((oViewRect.x + oViewRect.r) / 2) - nExtX / 2;
                Y1 = nPageH - nPageH * ((oViewRect.y + oViewRect.b) / 2) - nExtY / 2;
                X2 = X1 + nExtX;
                Y2 = Y1 + nExtY;
                break;
            case 270:
                X1 = nPageW - nPageW * ((oViewRect.y + oViewRect.b) / 2) - nExtY / 2;
                Y1 = nPageH * ((oViewRect.x + oViewRect.r) / 2) - nExtX / 2;
                X2 = X1 + nExtY;
                Y2 = Y1 + nExtX;
                break;
        }

        let nCurTime = new Date().getTime();

        let oProps = {
            rect:           [X1, Y1, X2, Y2],
            page:           nPage,
            name:           AscCommon.CreateGUID(),
            type:           AscPDF.ANNOTATIONS_TYPES.FreeText,
            author:         oUser.asc_getUserName(),
            modDate:        nCurTime,
            creationDate:   nCurTime,
            contents:       '',
            hidden:         false
        }

        let oFreeText = this.AddAnnotByProps(oProps);
        oFreeText.SetRotate(nRotAngle);
        
        oFreeText.SetFillColor([1, 1, 1]);
        oFreeText.SetStrokeColor([0, 0, 0]);
        oFreeText.SetWidth(1);
        oFreeText.SetAlign(AscPDF.ALIGN_TYPE.left);
        oFreeText.SetIntent(nType);
        
        this.SetMouseDownObject(oFreeText);
        oController.selection.groupSelection = oFreeText;
        oFreeText.selectStartPage = nPage;
        oFreeText.spTree.forEach(function(sp) {
            sp.selectStartPage = nPage;
        });

        switch (nType) {
            case AscPDF.FREE_TEXT_INTENT_TYPE.FreeText: {
                oFreeText.SetIntent(AscPDF.FREE_TEXT_INTENT_TYPE.FreeText);
                oFreeText.SetSubject('Text box');
                break;
            }
            // прописываем RD и Callout
            case AscPDF.FREE_TEXT_INTENT_TYPE.FreeTextCallout: {
                oFreeText.SetIntent(AscPDF.FREE_TEXT_INTENT_TYPE.FreeTextCallout);
                oFreeText.SetLineEnd(AscPDF.LINE_END_TYPE.OpenArrow);
                oFreeText.SetSubject('Text callout');
                
                let oTxBoxRect;
                let x1, y1, x2, y2, x3, y3;
                switch (nRotAngle) {
                    case 0:
                        oFreeText.SetRectangleDiff([nExtX / 2, 3 / 4 * nExtY, 0.5, 0.5]);
                        oTxBoxRect = oFreeText.GetTextBoxRect();
                        
                        x1 = X1;
                        y1 = Y1;
                        x2 = oTxBoxRect[0] - oFreeText.defaultPerpLength;
                        y2 = oTxBoxRect[1] + (oTxBoxRect[3] - oTxBoxRect[1]) / 2;
                        x3 = oTxBoxRect[0];
                        y3 = oTxBoxRect[1] + (oTxBoxRect[3] - oTxBoxRect[1]) / 2;
                        break;
                    case 90:
                        oFreeText.SetRectangleDiff([3 / 4 * nExtY, 0.5, 0.5, nExtX / 2]);
                        oTxBoxRect = oFreeText.GetTextBoxRect();

                        x1 = X1;
                        y1 = Y2;
                        x2 = oTxBoxRect[0] + (oTxBoxRect[2] - oTxBoxRect[0]) / 2;
                        y2 = oTxBoxRect[3] + oFreeText.defaultPerpLength;
                        x3 = oTxBoxRect[0] + (oTxBoxRect[2] - oTxBoxRect[0]) / 2;
                        y3 = oTxBoxRect[3]
                        break;
                    case 180:
                        oFreeText.SetRectangleDiff([0.5, 0.5, nExtX / 2, 3 / 4 * nExtY]);
                        oTxBoxRect = oFreeText.GetTextBoxRect();
                        
                        x1 = X2;
                        y1 = Y2;
                        x2 = oTxBoxRect[2] + oFreeText.defaultPerpLength;
                        y2 = oTxBoxRect[1] + (oTxBoxRect[3] - oTxBoxRect[1]) / 2;
                        x3 = oTxBoxRect[2];
                        y3 = oTxBoxRect[1] + (oTxBoxRect[3] - oTxBoxRect[1]) / 2;
                        break;
                    case 270:
                        oFreeText.SetRectangleDiff([0.5, nExtX / 2, 3 / 4 * nExtY, 0.5]);
                        oTxBoxRect = oFreeText.GetTextBoxRect();
                        
                        x1 = X2;
                        y1 = Y1;
                        x2 = oTxBoxRect[0] + (oTxBoxRect[2] - oTxBoxRect[0]) / 2;
                        y2 = oTxBoxRect[1] - oFreeText.defaultPerpLength;
                        x3 = oTxBoxRect[0] + (oTxBoxRect[2] - oTxBoxRect[0]) / 2;
                        y3 = oTxBoxRect[1];
                        break;
                }

                oFreeText.SetCallout([x1, y1, x2, y2, x3, y3]);
                break;
            }
        }

        oFreeText.SetInTextBox(true);
    };
    CPDFDoc.prototype.AddStampAnnot = function(nType, nPage, oImage) {
        let oController = this.GetController();
        let nRotAngle   = this.Viewer.getPageRotate(nPage);
        let oFile       = this.Viewer.file;
        let oViewRect   = this.Viewer.getViewingRect(nPage);
        let oNativePage = oFile.pages[nPage];
        let nPageW      = oNativePage.W;
        let nPageH      = oNativePage.H;
        let oUser       = Asc.editor.User;

        let nExtX = 200;
        let nExtY = 40;

        if (oImage) {
            nExtX   = Math.max(1, oImage.Image.width * g_dKoef_pix_to_mm);
            nExtY   = Math.max(1, oImage.Image.height * g_dKoef_pix_to_mm);
            let nKoeff  = Math.min(1.0, 1.0 / Math.max(nExtX / nPageW, nExtY / nPageH));

            nExtX = Math.max(5, nExtX * nKoeff); 
            nExtY = Math.max(5, nExtY * nKoeff); 
        }

        let X1, Y1, X2, Y2;
        switch (nRotAngle) {
            case 0:
                X1 = nPageW * ((oViewRect.x + oViewRect.r) / 2) - nExtX / 2;
                Y1 = nPageH * ((oViewRect.y + oViewRect.b) / 2) - nExtY / 2;
                X2 = X1 + nExtX;
                Y2 = Y1 + nExtY;
                break;
            case 90:
                X1 = nPageW * ((oViewRect.y + oViewRect.b) / 2) - nExtY / 2;
                Y1 = nPageH - nPageH * ((oViewRect.x + oViewRect.r) / 2) - nExtX / 2;
                X2 = X1 + nExtY;
                Y2 = Y1 + nExtX;
                break;
            case 180:
                X1 = nPageW - nPageW * ((oViewRect.x + oViewRect.r) / 2) - nExtX / 2;
                Y1 = nPageH - nPageH * ((oViewRect.y + oViewRect.b) / 2) - nExtY / 2;
                X2 = X1 + nExtX;
                Y2 = Y1 + nExtY;
                break;
            case 270:
                X1 = nPageW - nPageW * ((oViewRect.y + oViewRect.b) / 2) - nExtY / 2;
                Y1 = nPageH * ((oViewRect.x + oViewRect.r) / 2) - nExtX / 2;
                X2 = X1 + nExtY;
                Y2 = Y1 + nExtX;
                break;
        }

        let nCurTime = new Date().getTime();

        let oProps = {
            rect:           [X1, Y1, X2, Y2],
            page:           nPage,
            name:           AscCommon.CreateGUID(),
            type:           AscPDF.ANNOTATIONS_TYPES.Stamp,
            author:         oUser.asc_getUserName(),
            modDate:        nCurTime,
            creationDate:   nCurTime,
            contents:       '',
            hidden:         false
        }

        let oStamp = this.AddAnnotByProps(oProps);
        oStamp.SetRotate(nRotAngle);
        oStamp.SetWidth(1);

        if (oImage) {
            let oUniFill = new AscFormat.CUniFill();
            let oBlipFill = new AscFormat.CBlipFill();
            oUniFill.setFill(oBlipFill);
			oBlipFill.setRasterImageId(AscFormat.checkRasterImageId(oImage.src));
			oBlipFill.setStretch(true);
			oStamp.setFill(oUniFill);
        }
        // oStamp.SetSubject('Text box');
        // oStamp.SetIntent(nType);
        
        this.SetMouseDownObject(oStamp);
        oStamp.selectStartPage = nPage;
        oController.selectObject(oStamp, nPage);
    };
    CPDFDoc.prototype.AddImages = function(arrImages) {
        let oViewer     = this.Viewer;
        let nCurPage    = oViewer.currentPage;
        let oController = this.GetController();
        let nPageW      = this.GetPageWidthMM(nCurPage);
        let nPageH      = this.GetPageHeightMM(nCurPage);
        
        this.BlurActiveObject();

        for (let i = 0; i < arrImages.length; i++) {
            let _image  = arrImages[i];

            let nExtX   = Math.max(1, _image.Image.width * g_dKoef_pix_to_mm);
            let nExtY   = Math.max(1, _image.Image.height * g_dKoef_pix_to_mm);
            let nKoeff  = Math.min(1.0, 1.0 / Math.max(nExtX / nPageW, nExtY / nPageH));
            
            let nNewExtX    = Math.max(5, nExtX * nKoeff); 
            let nNewExtY    = Math.max(5, nExtY * nKoeff); 
            let oPos        = private_computeDrawingAddingPos(nCurPage, nNewExtX, nNewExtY);

            let oImage = new AscPDF.CPdfImage();
            AscFormat.fillImage(oImage, _image.src, 0, 0, nNewExtX, nNewExtY, _image.videoUrl, _image.audioUrl);

            let oXfrm = oImage.getXfrm();
            oXfrm.setOffX(oPos.x);
            oXfrm.setOffY(oPos.y);

            this.AddDrawing(oImage, nCurPage);

            if (i == 0) {
                this.SetMouseDownObject(oImage);
            }

            oImage.select(oController, nCurPage);
        }
    };
    CPDFDoc.prototype.ShapeApply = function(shapeProps) {
        let oController = this.GetController();
        let aAdditionalObjects = null;
        if (AscFormat.isRealNumber(shapeProps.Width) && AscFormat.isRealNumber(shapeProps.Height)) {
            aAdditionalObjects = oController.getConnectorsForCheck2();
        }

        oController.checkSelectedObjectsAndCallback(oController.applyDrawingProps, [shapeProps], false, AscDFH.historydescription_Presentation_SetShapeProps, aAdditionalObjects);
    };
    CPDFDoc.prototype.ChangeShapeType = function (sShapeType) {
        let oController = this.GetController();
        oController.checkSelectedObjectsAndCallback(oController.applyDrawingProps, [{type: sShapeType}], false, AscDFH.historydescription_Presentation_ChangeShapeType);
    };
    CPDFDoc.prototype.SetImageProps = function(oPr) {
        let oController         = this.GetController();
        let aAdditionalObjects  = null;
        if (AscFormat.isRealNumber(oPr.Width) && AscFormat.isRealNumber(oPr.Height)) {
            aAdditionalObjects = oController.getConnectorsForCheck2();
        }

        oController.checkSelectedObjectsAndCallback(oController.applyDrawingProps, [oPr], false, AscDFH.historydescription_Presentation_SetImageProps, aAdditionalObjects);
    };
    CPDFDoc.prototype.FitImagesToPage = function () {
        let oController = this.GetController();
        oController.fitImagesToPage();
    };
    CPDFDoc.prototype.AddTextWithPr = function(sText, oSettings) {
        let oController = this.GetController();
        oController.addTextWithPr(sText, oSettings);
    };
    CPDFDoc.prototype.GetChartObject = function(nType) {
        let oController = this.GetController();
        return oController.getChartObject(nType);
    };
    CPDFDoc.prototype.GetController = function() {
        return this.DrawingObjects;
    };
    CPDFDoc.prototype.BringToFront = function() {
        this.DoAction(function() {
            let oController = this.GetController();

            if (!(oController.selection.groupSelection)) {
                for (let i = oController.selectedObjects.length - 1; i > -1; --i) {
                    let oShape = oController.selectedObjects[i];
                    let oDrawings = oController.getDrawingObjects(oShape.GetPage());
                    this.ChangeDrawingPosInPageTree(oShape, oDrawings.length - 1);
                }
            }
            else {
                oController.selection.groupSelection.bringToFront();
            }
        }, AscDFH.historydescription_Document_GrObjectsBringToFront, this);
    };
    CPDFDoc.prototype.BringForward = function() {
        this.DoAction(function() {
            let oController = this.GetController();
            
            if (!(oController.selection.groupSelection)) {
                for (let i = oController.selectedObjects.length - 1; i > -1; --i) {
                    let oShape = oController.selectedObjects[i];
                    let oDrawings = oController.getDrawingObjects(oShape.GetPage());
                    let nCurPos = oDrawings.indexOf(oShape);

                    this.ChangeDrawingPosInPageTree(oShape, nCurPos + 1);
                }
            }
            else {
                oController.selection.groupSelection.bringForward();
            }
        }, AscDFH.historydescription_Document_GrObjectsBringForwardGroup, this);
    };
    CPDFDoc.prototype.SendToBack = function() {
        this.DoAction(function() {
            let oController = this.GetController();
            
            if (!(oController.selection.groupSelection)) {
                for (let i = oController.selectedObjects.length - 1; i > -1; --i) {
                    let oShape = oController.selectedObjects[i];
                    this.ChangeDrawingPosInPageTree(oShape, 0);
                }
            }
            else {
                oController.selection.groupSelection.sendToBack();
            }
        }, AscDFH.historydescription_Document_GrObjectsSendToBackGroup, this);
    };
    CPDFDoc.prototype.BringBackward = function() {
        this.DoAction(function() {
            let oController = this.GetController();
            
            if (!(oController.selection.groupSelection)) {
                for (let i = oController.selectedObjects.length - 1; i > -1; --i) {
                    let oShape = oController.selectedObjects[i];
                    let oDrawings = oController.getDrawingObjects(oShape.GetPage());
                    let nCurPos = oDrawings.indexOf(oShape);

                    this.ChangeDrawingPosInPageTree(oShape, nCurPos - 1);
                }
            }
            else {
                oController.selection.groupSelection.bringForward();
            }
        }, AscDFH.historydescription_Document_GrObjectsBringForwardGroup, this);
    };
    CPDFDoc.prototype.ChangeDrawingPosInPageTree = function(oDrawing, nNewPos) {
        let oController = this.GetController();

        let aDrawings   = oController.getDrawingObjects(oDrawing.GetPage());
        let nPage       = oDrawing.GetPage();
        let nOldPos     = aDrawings.indexOf(oDrawing);

        if (nNewPos >= aDrawings.length || nNewPos < 0 || nOldPos == nNewPos) {
            return;
        }

        this.RemoveDrawing(oDrawing.GetId());
        this.AddDrawing(oDrawing, nPage, nNewPos);
        oDrawing.select(oController, nPage);
        this.SetMouseDownObject(oDrawing);
        oDrawing.AddToRedraw();
    };
    CPDFDoc.prototype.PutShapesAlign = function(nType, nAlignToType) {
        let oController = this.GetController();
        
        if (!AscFormat.isRealNumber(nAlignToType)) {
			nAlignToType = Asc.c_oAscObjectsAlignType.Page;
		}

        let bSelected = nAlignToType === Asc.c_oAscObjectsAlignType.Selected;
		switch (nType) {
			case c_oAscAlignShapeType.ALIGN_LEFT: {
				oController.checkSelectedObjectsAndCallback(oController.alignLeft, [bSelected]);
				break;
			}
			case c_oAscAlignShapeType.ALIGN_RIGHT: {
				oController.checkSelectedObjectsAndCallback(oController.alignRight, [bSelected]);
				break;
			}
			case c_oAscAlignShapeType.ALIGN_TOP: {
				oController.checkSelectedObjectsAndCallback(oController.alignTop, [bSelected]);
				break;
			}
			case c_oAscAlignShapeType.ALIGN_BOTTOM: {
				oController.checkSelectedObjectsAndCallback(oController.alignBottom, [bSelected]);
				break;
			}
			case c_oAscAlignShapeType.ALIGN_CENTER: {
				oController.checkSelectedObjectsAndCallback(oController.alignCenter, [bSelected]);
				break;
			}
			case c_oAscAlignShapeType.ALIGN_MIDDLE: {
				oController.checkSelectedObjectsAndCallback(oController.alignMiddle, [bSelected]);
				break;
			}
			default:
				break;
		}
    };
    CPDFDoc.prototype.DistributeDrawingsHorizontally = function(alignType) {
        let oController = this.GetController();
        let bSelected = alignType === Asc.c_oAscObjectsAlignType.Selected;

        oController.checkSelectedObjectsAndCallback(oController.distributeHor, [bSelected]);
    };
    CPDFDoc.prototype.DistributeDrawingsVertically = function(alignType) {
        let oController = this.GetController();
        let bSelected = alignType === Asc.c_oAscObjectsAlignType.Selected;
        
        oController.checkSelectedObjectsAndCallback(oController.distributeVer, [bSelected]);
    };

    //-----------------------------------------------------------------------------------
    // Функции для работы с таблицами
    //-----------------------------------------------------------------------------------

    CPDFDoc.prototype.GetTableForPreview = function () {
        return AscFormat.ExecuteNoHistory(function () {
            let _x_mar = 10;
            let _y_mar = 10;
            let _r_mar = 10;
            let _b_mar = 10;
            let _pageW = 297;
            let _pageH = 210;
            let W = (_pageW - _x_mar - _r_mar);
            let H = (_pageH - _y_mar - _b_mar);
            let oGrFrame = this.private_Create_TableGraphicFrame(5, 5, this.DefaultTableStyleId, W, H, _x_mar, _y_mar, this.Viewer.currentPage, true);
            oGrFrame.setBDeleted(true);
            oGrFrame.SetPage(this.Viewer.currentPage);
            return oGrFrame.graphicObject;
        }, this, []);
    };
    CPDFDoc.prototype.CheckTableForPreview = function (oTable) {};

    CPDFDoc.prototype.SetTableProps = function(oTablePr) {
        this.DoAction(function() {
            let oController = this.GetController();
            oController.setTableProps(oTablePr);
        }, AscDFH.historydescription_Presentation_TblApply, this);
    };

    CPDFDoc.prototype.ApplyTableFunction = function (Function, bBefore, bAll, Cols, Rows) {
        let oController = this.GetController();
        if(!oController)
            return null;
    
        let result = null;
    
        let args;
        if (AscFormat.isRealNumber(Rows) && AscFormat.isRealNumber(Cols)) {
            args = [Cols, Rows];
        } else {
            args = [bBefore];
        }
        let oTargetText = AscFormat.getTargetTextObject(oController);
        let oTable;
        if (oTargetText && oTargetText.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
            oTable = oTargetText.graphicObject;
        }
        else {
            let oByTypes = oController.getSelectedObjectsByTypes(true);
            if (oByTypes.tables.length === 1) {
                let oGrFrame = oByTypes.tables[0];
                oTable = oGrFrame.graphicObject;
                if (Function !== AscWord.CTable.prototype.DistributeTableCells) {
                    oGrFrame.Set_CurrentElement();
                    if (!(bAll === true)) {
                        if (bBefore) {
                            oTable.MoveCursorToStartPos();
                        } else {
                            oTable.MoveCursorToStartPos();
                        }
                    } else {
                        oTable.SelectAll();
                    }
                }
            }
        }
        if (oTable) {
            oTable.Parent.SetNeedRecalc(true);
            result = Function.apply(oTable, args);
            if (oTable.Content.length === 0) {
                this.RemoveDrawing(oTable.Parent.GetId());
                return result;
            }
        }
        return result;
    };
    
    
    CPDFDoc.prototype.AddTableRow = function (bBefore) {
        this.ApplyTableFunction(CTable.prototype.AddTableRow, bBefore);
    };
    
    CPDFDoc.prototype.AddTableColumn = function (bBefore) {
        this.ApplyTableFunction(CTable.prototype.AddTableColumn, bBefore);
    };
    
    CPDFDoc.prototype.RemoveTableRow = function () {
        this.ApplyTableFunction(CTable.prototype.RemoveTableRow, undefined);
    };
    
    CPDFDoc.prototype.RemoveTableColumn = function () {
        this.ApplyTableFunction(CTable.prototype.RemoveTableColumn, true);
    };
    
    CPDFDoc.prototype.DistributeTableCells = function (isHorizontally) {
        return this.ApplyTableFunction(CTable.prototype.DistributeTableCells, isHorizontally);
    };
    
    CPDFDoc.prototype.MergeTableCells = function () {
        this.ApplyTableFunction(CTable.prototype.MergeTableCells, false, true);
    };
    
    CPDFDoc.prototype.SplitTableCells = function (Cols, Rows) {
        this.ApplyTableFunction(CTable.prototype.SplitTableCells, true, true, parseInt(Cols, 10), parseInt(Rows, 10));
    };
    
    CPDFDoc.prototype.SelectTable = function (Type) {
        const oController = this.GetController();
        if (oController) {
            let oByTypes = oController.getSelectedObjectsByTypes(true);
            if (oByTypes.tables.length === 1) {
                let oGrFrame = oByTypes.tables[0];
                oGrFrame.Set_CurrentElement();
                oGrFrame.graphicObject.SelectTable(Type);
                this.Viewer.onUpdateOverlay();
                return oGrFrame;
            }
        }
    };
    
    CPDFDoc.prototype.Table_CheckFunction = function (Function) {
        let oController = this.GetController();
        if (oController) {
            let oTextDrawing = AscFormat.getTargetTextObject(oController);
            if (oTextDrawing && oTextDrawing.getObjectType() === AscDFH.historyitem_type_GraphicFrame) {
                return Function.apply(oTextDrawing.graphicObject, []);
            }
        }
        return false;
    };
    
    CPDFDoc.prototype.CanMergeTableCells = function () {
        return this.Table_CheckFunction(CTable.prototype.CanMergeTableCells);
    };
    
    CPDFDoc.prototype.CanSplitTableCells = function () {
        return this.Table_CheckFunction(CTable.prototype.CanSplitTableCells);
    };
    
    CPDFDoc.prototype.CheckTableCoincidence = function (Table) {
        return false;
    };


    CPDFDoc.prototype.InitDefaultTextListStyles = function() {
        let oTextStyles     = new AscFormat.CTextStyles();
        let oTextListStyle  = new AscFormat.TextListStyle();

        oTextStyles.otherStyle = oTextListStyle;

        let nDefTab     = 25.4
        let nIndStep    = 12.7;
        let nJc         = AscCommon.align_Left;

        for (let i = 0; i < 10; i++) {
            let oParaPr = new AscWord.CParaPr();
            oTextListStyle.levels[i] = oParaPr;

            if (i == 9)
                break;

            oParaPr.DefaultTab  = nDefTab;
            oParaPr.Ind.Left    = i * nIndStep;
            oParaPr.Jc          = nJc;
        }

        this.styles.txStyles = oTextStyles;
    };
    CPDFDoc.prototype.InitDefaultTableStyles = function () {
        this.globalTableStyles = new CStyles(false);
    
        this.globalTableStyles.Id = AscCommon.g_oIdCounter.Get_NewId();
        AscCommon.g_oTableId.Add(this.globalTableStyles, this.globalTableStyles.Id);
        this.DefaultTableStyleId = AscFormat.CreatePresentationTableStyles(this.globalTableStyles, this.TableStylesIdMap);
    };

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Collaborative editing
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    CPDFDoc.prototype.PauseRecalculate = function() {};
    CPDFDoc.prototype.EndPreview_MailMergeResult = function() {};
    CPDFDoc.prototype.Get_SelectionState2 = function() {};
    CPDFDoc.prototype.Save_DocumentStateBeforeLoadChanges = function() {};
    CPDFDoc.prototype.Load_DocumentStateAfterLoadChanges = function() {};
    CPDFDoc.prototype.Check_MergeData = function() {};
    CPDFDoc.prototype.Set_SelectionState2 = function() {};
    CPDFDoc.prototype.ResumeRecalculate = function() {};
    CPDFDoc.prototype.RecalculateByChanges = function(arrChanges, nStartIndex, nEndIndex) {
        this.ClearSearch();

        // Обновляем позицию курсора
        this.NeedUpdateTarget = true;

        // Увеличиваем номер пересчета
        this.RecalcId++;

        this.History.Get_RecalcData(null, arrChanges, nStartIndex, nEndIndex);
    };
    CPDFDoc.prototype.UpdateTracks = function() {};
    CPDFDoc.prototype.GetOFormDocument = function() {};
    CPDFDoc.prototype.Continue_FastCollaborativeEditing = function() {
        let oController = this.GetController();

        if (true === this.CollaborativeEditing.Get_GlobalLock()) {
            if (this.Api.forceSaveUndoRequest) {
                this.Api.asc_Save(true);
            }

            return;
        }

        if (this.Api.isLongAction()) {
            return;
        }

        if (Asc.editor.FontLoader.isWorking()) {
            return;
        }
        
        if (true !== this.CollaborativeEditing.Is_Fast() || true === this.CollaborativeEditing.Is_SingleUser()) {
            return;
        }

        if (true === this.Api.isStartAddShape || oController.isTrackingDrawings() || this.Api.isOpenedChartFrame) {
            return;
        }

        let HaveChanges = this.History.Have_Changes(true);
        if (true !== HaveChanges && (true === this.CollaborativeEditing.Have_OtherChanges() || 0 !== this.CollaborativeEditing.getOwnLocksLength())) {
            // Принимаем чужие изменения. Своих нет, но функцию отсылки надо вызвать, чтобы снять локи.
            this.CollaborativeEditing.Apply_Changes();
            this.CollaborativeEditing.Send_Changes();
        }
        else if (true === HaveChanges || true === this.CollaborativeEditing.Have_OtherChanges()) {
            this.Api.asc_Save(true);
        }

        let CurTime = new Date().getTime();
        if (true === this.NeedUpdateTargetForCollaboration && (CurTime - this.LastUpdateTargetTime > 1000) || this.NeedUpdateTargetForCollaborationForce) {
            
            if (true !== HaveChanges || this.NeedUpdateTargetForCollaborationForce) {
                
                let CursorInfo = this.History.Get_DocumentPositionBinary();
                this.Api.CoAuthoringApi.sendCursor(CursorInfo);
                this.LastUpdateTargetTime = CurTime;
            }
            else {
                this.LastUpdateTargetTime = CurTime;
            }

            this.NeedUpdateTargetForCollaboration = false;
            this.NeedUpdateTargetForCollaborationForce = false;
        }
    };
    CPDFDoc.prototype.private_UpdateTargetForCollaboration = function(bForce) {
        this.NeedUpdateTargetForCollaboration = true;
        if (bForce) {
            this.NeedUpdateTargetForCollaborationForce = bForce;
        }
    };
    CPDFDoc.prototype.Get_DocumentPositionInfoForCollaborative = function() {
        let oActiveObj = this.GetActiveObject();
        
        if (!oActiveObj || (Asc.editor.isRestrictionView()) && !this.IsViewerObject(oActiveObj)) {
            return;
        }
        
        if (oActiveObj.IsDrawing() && oActiveObj.IsInTextBox()) {
            let oContent = oActiveObj.GetDocContent();
            if (oContent) {
                let aDocPos = oContent.getDocumentContentPosition();
                return aDocPos[aDocPos.length - 1];
            }
        }
        else {
            return {Class: oActiveObj, Position: undefined};
        }
    };
    CPDFDoc.prototype.Update_ForeignCursor = function(CursorInfo, UserId, Show, UserShortId) {
        if (!this.Api.User)
            return;
    
        if (UserId === this.Api.CoAuthoringApi.getUserConnectionId())
            return;
    
        let sUserName = this.Api.CoAuthoringApi.getParticipantName(UserId);
    
        // "" - это означает, что курсор нужно удалить
        if (!CursorInfo || "" === CursorInfo || !AscCommon.UserInfoParser.isUserVisible(sUserName))
        {
            this.Remove_ForeignCursor(UserId);
            this.Viewer.onUpdateOverlay();
            return;
        }
    
        let Changes = new AscCommon.CCollaborativeChanges();
        let Reader  = Changes.GetStream(CursorInfo);
    
        let oObjId   = Reader.GetString2();
        let InRunPos = Reader.GetLong();
    
        let oTargetObj = AscCommon.g_oTableId.Get_ById(oObjId);
        let bUpdateOverlayOnRemoveCursor = false;
        if (Array.isArray(this.CollaborativeEditing.m_oSelectedObjects[UserId]) && this.CollaborativeEditing.m_oSelectedObjects[UserId].length != 0) {
            bUpdateOverlayOnRemoveCursor = true;
        }

        this.Remove_ForeignCursor(UserId);
        if (bUpdateOverlayOnRemoveCursor) {
            this.Viewer.onUpdateOverlay();
        }

        if (!oTargetObj) {
            return;
        }
    
        if (oTargetObj instanceof AscWord.Run) {
            let CursorPos = [{Class : oTargetObj, Position : InRunPos}];
            oTargetObj.GetDocumentPositionFromObject(CursorPos);
            this.CollaborativeEditing.Add_ForeignCursor(UserId, CursorPos, UserShortId);
            
            if (true === Show) {
                this.CollaborativeEditing.Update_ForeignCursorPosition(UserId, oTargetObj, InRunPos, true);
                this.GetDrawingDocument().Collaborative_TargetsUpdate();
            }

            this.CollaborativeEditing.Show_ForeignCursorLabel(UserId);
        }
        else {
            this.CollaborativeEditing.Add_ForeignSelectedObject(UserId, oTargetObj, UserShortId);
            let color = AscCommon.getUserColorById(UserShortId, null, true);
            this.Show_ForeignSelectedObjectLabel(UserId, oTargetObj, color);
            this.Viewer.onUpdateOverlay();
        }
    };
    CPDFDoc.prototype.Draw_ForeingSelection = function(nDrawPage) {
        let oViewer             = this.Viewer;
        let oOverlay            = this.Viewer.overlay;
        let oCtx                = this.Viewer.overlay.m_oContext;
        let oSelectedObjects    = this.CollaborativeEditing.m_oSelectedObjects;

        for (let userId in oSelectedObjects) {
            let aObjects = oSelectedObjects[userId];

            let oColor = AscCommon.getUserColorById(this.CollaborativeEditing.m_aForeignCursorsId[userId], null, true);

            for (let i = 0; i < aObjects.length; i++) {
                let aRect = aObjects[i].GetOrigRect();
                let nPage = aObjects[i].GetPage();
                let nRotRad = aObjects[i].GetRot ? aObjects[i].GetRot() : 0;

                if (nDrawPage != nPage) {
                    continue;
                }

                let nScale = oViewer.getDrawingPageScale(nPage) * AscCommon.AscBrowser.retinaPixelRatio * oViewer.zoom;

                let xCenter = oViewer.width >> 1;
                if (oViewer.documentWidth > oViewer.width)
                {
                    xCenter = (oViewer.documentWidth >> 1) - (oViewer.scrollX) >> 0;
                }
                let yPos    = oViewer.scrollY >> 0;
                let page    = oViewer.drawingPages[nPage];
                let w       = (page.W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
                let h       = (page.H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
                let indLeft = ((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - (w >> 1);
                let indTop  = ((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

                if (true == oViewer.isLandscapePage(nPage))
                    indLeft = indLeft + (w - h) / 2;

                let X = aRect[0] * nScale - 0.5 >> 0;
                let Y = aRect[1] * nScale - 0.5 >> 0;
                let W = (aRect[2] - aRect[0]) * nScale + 0.5 >> 0;
                let H = (aRect[3] - aRect[1]) * nScale + 0.5 >> 0;

                oCtx.strokeStyle = "rgb(" + oColor.r + "," + oColor.g + "," + oColor.b + ")";
                oCtx.lineWidth = 2;

                if (nRotRad != 0) {
                    let cosTheta = Math.cos(nRotRad);
                    let sinTheta = Math.sin(nRotRad);
                
                    let centerX = indLeft + X + W / 2;
                    let centerY = indTop + Y + H / 2;
                
                    let corners = [
                        { x: -W / 2, y: -H / 2 }, // верхний левый
                        { x: W / 2, y: -H / 2 },  // верхний правый
                        { x: W / 2, y: H / 2 },   // нижний правый
                        { x: -W / 2, y: H / 2 },  // нижний левый
                    ];
                
                    let rotatedCorners = corners.map(function (pt) {
                        return {
                            x: centerX + pt.x * cosTheta - pt.y * sinTheta,
                            y: centerY + pt.x * sinTheta + pt.y * cosTheta
                        };
                    });
                
                    oCtx.beginPath();
                    oOverlay.CheckPoint(rotatedCorners[0].x, rotatedCorners[0].y);
                    oCtx.moveTo(rotatedCorners[0].x, rotatedCorners[0].y);
                    for (let i = 1; i < rotatedCorners.length; i++) {
                        oOverlay.CheckPoint(rotatedCorners[i].x, rotatedCorners[i].y);
                        oCtx.lineTo(rotatedCorners[i].x, rotatedCorners[i].y);
                    }
                    oCtx.closePath();
                    oCtx.stroke();
                } else {
                    oOverlay.CheckPoint(indLeft + X, indTop + Y);
                    oOverlay.CheckPoint(indLeft + X + W, indTop + Y + H);

                    oCtx.beginPath();
                    oCtx.rect(indLeft + X, indTop + Y, W, H);
                    oCtx.stroke();
                }
            }
        }
    };
    CPDFDoc.prototype.Show_ForeignSelectedObjectLabel = function(userId, foreignSelectObj, color) {
		let oApi = Asc.editor;

		if (foreignSelectObj.LabelTimer)
			clearTimeout(foreignSelectObj.LabelTimer);

		foreignSelectObj.LabelTimer = setTimeout(function()
		{
			foreignSelectObj.LabelTimer = null;
			oApi.sync_HideForeignCursorLabel(userId);
		}, AscCommon.FOREIGN_CURSOR_LABEL_HIDETIME);

        let oOrigRect = foreignSelectObj.GetOrigRect();
        let nPage = foreignSelectObj.GetPage();
		let oTr = this.pagesTransform[nPage].invert;

		let oPoint = oTr.TransformPoint(oOrigRect[0], oOrigRect[1]);
		
		oApi.sync_ShowForeignCursorLabel(userId, oPoint.x, oPoint.y, color);
	};
    CPDFDoc.prototype.Remove_ForeignCursor = function(UserId, oObject) {
        this.CollaborativeEditing.Remove_ForeignCursor(UserId);
        this.CollaborativeEditing.Remove_FiregnSelectedObject(UserId, oObject);
    };

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Required extensions
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    CPDFDoc.prototype.Is_Inline = function() {};
    CPDFDoc.prototype.OnChangeForm = function() {};
    CPDFDoc.prototype.TurnOffCheckChartSelection = function() {};
    CPDFDoc.prototype.TurnOnCheckChartSelection = function() {};
    CPDFDoc.prototype.UpdateRulers = function() {};
    CPDFDoc.prototype.UpdateSelection = function() {};
    CPDFDoc.prototype.Get_Api = function() {
        return Asc.editor;
    };
    CPDFDoc.prototype.sendEvent = function() {
        if (!this.Api)
            return;
        
        this.Api.sendEvent.apply(this.Api, arguments);
    };
    CPDFDoc.prototype.Get_CollaborativeEditing = function() {
        return this.CollaborativeEditing;
    };
    CPDFDoc.prototype.Clear_ContentChanges = function() {
		this.pagesContentChanges.Clear();
    };
    CPDFDoc.prototype.UpdateMaxApIdx = function(nApIdx) {
        AscCommon.g_oIdCounter.m_nIdCounterEdit = Math.max(nApIdx, AscCommon.g_oIdCounter.m_nIdCounterEdit);
    };
    CPDFDoc.prototype.Add_ContentChanges = function(Changes) {
        this.pagesContentChanges.Add(Changes);
    };
    CPDFDoc.prototype.Refresh_ContentChanges = function() {
		this.pagesContentChanges.Refresh();
    };
    CPDFDoc.prototype.GetRecalcId = function () {
        return Infinity;
    };
    CPDFDoc.prototype.RemoveBeforePaste = function() {};
    CPDFDoc.prototype.Document_UpdateInterfaceState = function() {};
    CPDFDoc.prototype.IsViewModeInReview = function() {
        return false;
    };
    CPDFDoc.prototype.Is_OnRecalculate = function() {
        return false;
    };
	CPDFDoc.prototype.IsActionStarted = function() {
		return false;
	};
    CPDFDoc.prototype.Get_AbsolutePage = function () {
        return 0;
    };
    CPDFDoc.prototype.Get_AbsoluteColumn = function () {
        return 0;
    };
    CPDFDoc.prototype.GetPrevElementEndInfo = function (CurElement) {
        return null;
    };
    CPDFDoc.prototype.Get_TextBackGroundColor = function () {
        return new CDocumentColor(255, 255, 255, false);
    };
    CPDFDoc.prototype.IsCell = function (isReturnCell) {
        if (isReturnCell)
            return null;
    
        return false;
    };
    CPDFDoc.prototype.GetCurPage = function() {
        return this.Viewer.currentPage;
    };
    CPDFDoc.prototype.GetPageInfo = function(nPage) {
        return this.Viewer.pagesInfo.pages[nPage];
    };
    CPDFDoc.prototype.GetThumbnails = function() {
        return this.Viewer.thumbnails;
    };
    CPDFDoc.prototype.GetPagesCount = function() {
        return this.Viewer.file.pages.length;
    };
    CPDFDoc.prototype.GetSelectionState = function() {
        const oSelectionState = {};

        let oController = this.GetController();

        oSelectionState.CurPage             = this.Viewer.currentPage;
        oSelectionState.activeObject        = this.GetActiveObject();
        oSelectionState.drawingSelection    = oController.getSelectionState();
        oSelectionState.HistoryIndex        = this.History.Index;

        return oSelectionState;
    };
    CPDFDoc.prototype.SetSelectionState = function(oState) {
        let oController = this.GetController();

        this.SetMouseDownObject(oState.activeObject);
        if (false == this.Api.isRestrictionView()) {
            oController.setSelectionState(oState.drawingSelection);
        }

        if (oState.CurPage != -1 && oState.CurPage != this.Viewer.currentPage)
	        this.Viewer.navigateToPage(oState.CurPage);
    };
    CPDFDoc.prototype.IsSelectionLocked = function (nCheckType, oAdditionalData, isDontLockInFastMode, isIgnoreCanEditFlag) {
        return this.Document_Is_SelectionLocked(nCheckType, oAdditionalData, isIgnoreCanEditFlag, undefined, isDontLockInFastMode);
    };
    CPDFDoc.prototype.Document_Is_SelectionLocked = function (nCheckType, AdditionalData, isIgnoreCanEditFlag, aAdditionaObjects, DontLockInFastMode) {
        let oController         = this.GetController();
        let oCurPageInfo        = this.GetPageInfo(this.GetCurPage());
        let isRestrictionView   = Asc.editor.isRestrictionView();

        if (true === this.CollaborativeEditing.Get_GlobalLock()) {
            return true;
        }

        AscCommon.CollaborativeEditing.OnStart_CheckLock();

        if (nCheckType === AscCommon.changestype_None && AscCommon.isRealObject(AdditionalData) && AdditionalData.CheckType === AscCommon.changestype_Table_Properties) {
            nCheckType = AscCommon.changestype_Drawing_Props;
        }

        if ([AscCommon.changestype_Paragraph_TextProperties, AscCommon.changestype_Paragraph_Content].includes(nCheckType)) {
            let oTargetTextObject = oController.getTargetDocContent(false, true);
            if (oTargetTextObject) {
                nCheckType = AscCommon.changestype_Drawing_Props;
            }
            else {
                return false;
            }
        }

        switch (nCheckType) {
            // check selected object lock
            case AscCommon.changestype_Drawing_Props:
            case AscCommon.changestype_Delete: {
                if (false == [AscCommon.c_oAscLockTypes.kLockTypeMine, AscCommon.c_oAscLockTypes.kLockTypeNone].includes(oCurPageInfo.Lock.Type))
                    return true;

                let selected_objects = oController.selectedObjects.slice();
                if (oController.selection.groupSelection) {
                    selected_objects.push(oController.selection.groupSelection);
                }
                if (this.mouseDownAnnot) {
                    selected_objects.push(this.mouseDownAnnot);
                }

                if (isRestrictionView) {
                    for (let i = 0; i < selected_objects.length; i++) {
                        if (selected_objects[i].IsDrawing()) {
                            return true;
                        }
                    }
                }

                for (let i = 0; i < selected_objects.length; ++i) {
                    let check_obj = {
                        "type":     AscLockTypeElemPDF.Object,
                        "pageId":   oCurPageInfo.GetId(),
                        "objId":    selected_objects[i].GetId(),
                        "guid":     selected_objects[i].GetId()
                    };

                    selected_objects[i].Lock.Check(check_obj);
                }

                break;
            }
            case AscDFH.historydescription_Pdf_RemovePage:
            case AscDFH.historydescription_Pdf_EditPage:
            case AscDFH.historydescription_Pdf_RotatePage: {
                let aSelectedPagesIdx = AdditionalData;
                let aPagesInfo = this.Viewer.pagesInfo.pages;

                aSelectedPagesIdx.forEach(function(pageIdx) {
                    let oPageInfo = aPagesInfo[pageIdx];

                    if (oPageInfo) {
                        let oCheckData = {
                            "type":     AscLockTypeElemPDF.Page,
                            "pageId":   oPageInfo.GetId(),
                            "guid":     oPageInfo.GetId()
                        };

                        oPageInfo.Lock.Check(oCheckData);
                    }
                });

                break;
            }
            case AscCommon.changestype_2_Comment: {
                let sCommentId = AdditionalData;

                let selected_objects = oController.selectedObjects.slice();
                if (oController.selection.groupSelection) {
                    selected_objects.push(oController.selection.groupSelection);
                }
                if (this.mouseDownAnnot) {
                    selected_objects.push(this.mouseDownAnnot);
                }

                for (let i = 0; i < selected_objects.length; ++i) {
                    if (selected_objects[i].GetId() == sCommentId) {
                        let check_obj = {
                            "type":     AscLockTypeElemPDF.Object,
                            "pageId":   oCurPageInfo.GetId(),
                            "objId":    selected_objects[i].GetId(),
                            "guid":     selected_objects[i].GetId()
                        };
    
                        selected_objects[i].Lock.Check(check_obj);
                        break;
                    }
                }

                break;
            }
        }
        
        let isLocked = AscCommon.CollaborativeEditing.OnEnd_CheckLock(DontLockInFastMode);
        return isLocked;
    };
    
    CPDFDoc.prototype.Document_UpdateUndoRedoState = function() {};
    CPDFDoc.prototype.SetHighlightRequiredFields = function() {};
    CPDFDoc.prototype.SetLocalTrackRevisions = function() {};
    CPDFDoc.prototype.Document_UpdateUndoRedoState = function() {
        this.UpdateUndoRedo();
    };
    CPDFDoc.prototype.RecalculateCurPos = function() {};
    CPDFDoc.prototype.HaveRevisionChanges = function() {};
    CPDFDoc.prototype.ContinueSpellCheck = function() {};
    CPDFDoc.prototype.ContinueTrackRevisions = function() {};
    CPDFDoc.prototype.StartCollaborationEditing = function() {};
    CPDFDoc.prototype.Viewer_OnChangePosition = function() {};
    CPDFDoc.prototype.Document_CreateFontMap = function() { return {}};
    CPDFDoc.prototype.TurnOffSpellCheck = function() {};
    CPDFDoc.prototype.TrackDocumentPositions = function (arrPositions) {
        this.CollaborativeEditing.Clear_DocumentPositions();
    
        for (var nIndex = 0, nCount = arrPositions.length; nIndex < nCount; ++nIndex) {
            this.CollaborativeEditing.Add_DocumentPosition(arrPositions[nIndex]);
        }
    };
    /**
     * Обновляем отслеживаемые позиции
     * @param arrPositions
     */
    CPDFDoc.prototype.RefreshDocumentPositions = function (arrPositions) {
        for (var nIndex = 0, nCount = arrPositions.length; nIndex < nCount; ++nIndex) {
            this.CollaborativeEditing.Update_DocumentPosition(arrPositions[nIndex]);
        }
    };
    CPDFDoc.prototype.RemoveSelection = function(bNoResetChartSelection) {
        let oController = this.GetController();
        if (oController) {
            oController.resetSelection(undefined, bNoResetChartSelection);
        }
    };
    CPDFDoc.prototype.Set_TargetPos = function() {};
    CPDFDoc.prototype.GetSelectedDrawingObjectsCount = function () {
        var oController = this.GetController();
        var aSelectedObjects = oController.selection.groupSelection && !oController.selection.groupSelection.IsAnnot() ? oController.selection.groupSelection.selectedObjects : oController.selectedObjects;
        return aSelectedObjects.filter(function(obj) {
            return obj.IsDrawing();
        }).length;
    };
    CPDFDoc.prototype.isShapeChild = function() {};
    CPDFDoc.prototype.Recalculate = function(){};
    CPDFDoc.prototype.GetDocPosType = function() {};
    CPDFDoc.prototype.GetSelectedContent = function() {};
    CPDFDoc.prototype.Is_ShowParagraphMarks = function() {};
    CPDFDoc.prototype.CheckTargetUpdate = function(force) {
		if (force)
			this.NeedUpdateTarget = true;
		
		if (!this.NeedUpdateTarget)
			return
		
		let textController = this.getTextController();
		if (!textController)
		{
			this.NeedUpdateTarget = false;
			return;
		}
		
		if (textController.IsNeedRecalc())
			return;
		
		textController.GetDocContent().RecalculateCurPos();
		this.NeedUpdateTarget = false;
	};
    CPDFDoc.prototype.SetWordSelection = function(){};
    
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
		return this.styles;
	};
	CPDFDoc.prototype.GetTheme = function() {
		return this.theme;
	};
	CPDFDoc.prototype.Get_Theme = function() {
		return this.theme;
	};
	CPDFDoc.prototype.GetStyles = function() {
		return this.Get_Styles();
	};
    CPDFDoc.prototype.GetDefaultLanguage = function() {
        return this.GetStyles().Default.TextPr.Lang.Val;
    };
    CPDFDoc.prototype.GetTableStyles = function() {
        return this.globalTableStyles;
    };
    CPDFDoc.prototype.GetAllTableStyles = function() {
        return this.globalTableStyles.Style;
    };

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Extension required for CGraphicObjects
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    CPDFDoc.prototype.Get_ColorMap = function() {
        return this.clrSchemeMap;
    };
    CPDFDoc.prototype.GetColorMap = function() {
        return this.clrSchemeMap;
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
    CPDFDoc.prototype.GetPageWidthEMU = function(nPage) {
        nPage = nPage != undefined ? nPage : this.Viewer.currentPage;
        let oNativePage = this.Viewer.file.pages[nPage];

        return oNativePage.W * g_dKoef_pt_to_mm * g_dKoef_mm_to_emu;
    };
    CPDFDoc.prototype.GetPageHeightEMU = function(nPage) {
        nPage = nPage != undefined ? nPage : this.Viewer.currentPage;
        let oNativePage = this.Viewer.file.pages[nPage];

        return oNativePage.H * g_dKoef_pt_to_mm * g_dKoef_mm_to_emu;
    };
    CPDFDoc.prototype.GetPageWidth = function(nPage) {
        nPage = nPage != undefined ? nPage : this.Viewer.currentPage;
        let oNativePage = this.Viewer.file.pages[nPage];

        return oNativePage.W;
    };
    CPDFDoc.prototype.GetPageHeight = function(nPage) {
        nPage = nPage != undefined ? nPage : this.Viewer.currentPage;
        let oNativePage = this.Viewer.file.pages[nPage];
        
        return oNativePage.H;
    };
    CPDFDoc.prototype.GetPageWidthMM = function(nPage) {
        nPage = nPage != undefined ? nPage : this.Viewer.currentPage;
        let oNativePage = this.Viewer.file.pages[nPage];

        return oNativePage.W * g_dKoef_pt_to_mm;
    };
    CPDFDoc.prototype.GetPageHeightMM = function(nPage) {
        nPage = nPage != undefined ? nPage : this.Viewer.currentPage;
        let oNativePage = this.Viewer.file.pages[nPage];
        
        return oNativePage.H * g_dKoef_pt_to_mm;
    };
	CPDFDoc.prototype.GetApi = function() {
		return editor;
	};
	CPDFDoc.prototype.CanEdit = function() {
		return this.Api.canEdit();
	};
    CPDFDoc.prototype.IsShowShapeAdjustments = function() {
        return this.Api.canEdit() && false == this.Api.IsCommentMarker();
    };
    CPDFDoc.prototype.IsShowTableAdjustments = function() {
        return this.Api.canEdit() && false == this.Api.IsCommentMarker();
    };
    CPDFDoc.prototype.IsViewerObject = function(oObject) {
        return !!(oObject && oObject.IsAnnot && (oObject.IsAnnot() || oObject.IsForm() || oObject.group && oObject.group.IsAnnot()));
    };
    CPDFDoc.prototype.IsFillingFormMode = function() {
		return false;
	};
	CPDFDoc.prototype.getDrawingObjects = function() {
		return null;
	};
	CPDFDoc.prototype.checkDefaultFieldFonts = function(callback) {
		
		if (1 === this.defaultFontsLoaded)
			return true;
		
		if (0 === this.defaultFontsLoaded)
			return false;
		
		this.defaultFontsLoaded = 0;
		let _t = this;
		this.fontLoader.LoadFonts([AscPDF.DEFAULT_FIELD_FONT],
			function()
			{
				_t.defaultFontsLoaded = 1;
				if (callback)
					callback();
			}
		);

		return 1 === this.defaultFontsLoaded;
	};
    CPDFDoc.prototype.checkDefaultFonts = function(callback) {
        return this.checkFonts(["Arial", "Symbol", "Wingdings", "Courier New", "Times New Roman"], callback);
    };
    CPDFDoc.prototype.checkFieldFont = function(oField, callback) {
        if (!oField)
            return true;
        
        // при клике по кнопке внешний вид остается прежним, поэтому грузить шрифт не надо
        if (oField.GetType() == AscPDF.FIELD_TYPES.button && oField.IsNeedDrawFromStream())
            return true;

		let sFontName = oField.GetTextFontActual();

        if (!sFontName)
            return true;
        
		if (this.loadedFonts.includes(sFontName))
            return true;
		
		let _t = this;
		this.fontLoader.LoadFonts([sFontName],
			function()
			{
				_t.loadedFonts.push(sFontName);
				if (callback)
					callback();
			}
		);
	
		return false;
	};
	CPDFDoc.prototype.checkFonts = function(aFontsNames, callback) {
		let fontsToLoad = [];
		let fontMap     = {};

        let aExtended = [];
        AscFonts.FontPickerByCharacter.extendFonts(aExtended);
        aExtended.forEach(function(font) {
            aFontsNames.push(font.name);
        });
        
		for (let i = 0; i < aFontsNames.length; i++) {
			if (this.loadedFonts.includes(aFontsNames[i]) || fontsToLoad.includes(aFontsNames[i]))
				continue;
			
			fontsToLoad.push(aFontsNames[i]);
			fontMap[aFontsNames[i]] = true;
		}
		
		if (!fontsToLoad.length)
			return true;
		
		let _t = this;
		this.fontLoader.LoadFonts(fontMap,
			function()
			{
				_t.loadedFonts = _t.loadedFonts.concat(fontsToLoad);
				if (callback)
					callback();
			}
		);
		
		return false;
	};
    CPDFDoc.prototype.GetAllSignatures = function() {
        return [];
    };
	CPDFDoc.prototype.IsWordSelection = function() {
		return false;
	};
    CPDFDoc.prototype.GetCursorRealPosition = function() {
        return {
            X: this.CurPosition.X,
            Y: this.CurPosition.Y
        };
    };

	CPDFDoc.prototype.getTextController = function() {
        let oController = this.GetController();

		let activeForm      = this.activeForm;
		let activeAnnot     = this.mouseDownAnnot;
		let activeDrawing   = this.activeDrawing;
		
        if (oController.getSelectedArray().length > 1) {
            return null;
        }

		if (activeForm && this.checkFieldFont(activeForm) && activeForm.IsCanEditText()) {
			return activeForm;
		}
		else if (activeAnnot && activeAnnot.IsFreeText() && activeAnnot.IsInTextBox()) {
			return activeAnnot;
		}
		else if (activeDrawing && activeDrawing.GetDocContent()) {
			return activeDrawing;
		}
		
		return null;
	};

    // for touch
    CPDFDoc.prototype.GetSelectionBounds = function()
    {
        return null;
    };
    CPDFDoc.prototype.IsInForm = function(x, y, pageIndex)
    {
        // TODO:
        return (null != this.Viewer.getPageFieldByMouse(false)) ? true : false;
    };
    CPDFDoc.prototype.IsFormFieldEditing = function()
    {
        return false;
    };
    CPDFDoc.prototype.IsNumberingSelection = function()
    {
        return false;
    };
    CPDFDoc.prototype.Get_TargetPos = function()
    {
        return null;
    };

    function CActionQueue(oDoc) {
        this.doc                = oDoc;
        this.actions            = [];
        this.isInProgress       = false;
        this.curAction          = null;
        this.curActionIdx       = -1;
        this.callbackAfterFocus = null;
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
        this.callbackAfterFocus = null;
    };
    CActionQueue.prototype.Stop = function() {
        Asc.editor.canSave = true;

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
            this.doc.DoAction(function() {
                let oFirstAction = this.actions[0];
                if (oFirstAction) {
                    Asc.editor.canSave = false;
                    this.SetInProgress(true);
                    this.SetCurActionIdx(0);
                    oFirstAction.Do();
                }
            }, AscDFH.historydescription_Pdf_ExecActions, this);
        }
    };
    CActionQueue.prototype.Continue = function() {
        let oNextAction = this.GetNextAction();
        if (this.callbackAfterFocus && this.curAction.triggerType == AscPDF.FORMS_TRIGGERS_TYPES.OnFocus && (!oNextAction || oNextAction.triggerType != AscPDF.FORMS_TRIGGERS_TYPES.OnFocus))
            this.callbackAfterFocus();

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
        let sCrDate     = oProps.creationDate;
        let sModDate    = oProps.modDate;
        let sText       = oProps.contents;
        let isHidden    = !!oProps.hidden;
        
        let oAnnot;

        switch (nAnnotType) {
            case AscPDF.ANNOTATIONS_TYPES.Text:
                oAnnot = new AscPDF.CAnnotationText(sName, nPageNum, aRect, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Ink:
                oAnnot = new AscPDF.CAnnotationInk(sName, nPageNum, aRect, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Highlight:
                oAnnot = new AscPDF.CAnnotationHighlight(sName, nPageNum, aRect, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Underline:
                oAnnot = new AscPDF.CAnnotationUnderline(sName, nPageNum, aRect, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Strikeout:
                oAnnot = new AscPDF.CAnnotationStrikeout(sName, nPageNum, aRect, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Squiggly:
                oAnnot = new AscPDF.CAnnotationSquiggly(sName, nPageNum, aRect, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Caret:
                oAnnot = new AscPDF.CAnnotationCaret(sName, nPageNum, aRect, oPdfDoc);
                oAnnot.SetQuads([[aRect[0], aRect[1], aRect[2], aRect[1], aRect[0], aRect[3], aRect[2], aRect[3]]]);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Line:
                oAnnot = new AscPDF.CAnnotationLine(sName, nPageNum, aRect, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Square:
                oAnnot = new AscPDF.CAnnotationSquare(sName, nPageNum, aRect, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Circle:
                oAnnot = new AscPDF.CAnnotationCircle(sName, nPageNum, aRect, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Polygon:
                oAnnot = new AscPDF.CAnnotationPolygon(sName, nPageNum, aRect, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.PolyLine:
                oAnnot = new AscPDF.CAnnotationPolyLine(sName, nPageNum, aRect, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.FreeText:
                oAnnot = new AscPDF.CAnnotationFreeText(sName, nPageNum, aRect, oPdfDoc);
                break;
            case AscPDF.ANNOTATIONS_TYPES.Stamp:
                oAnnot = new AscPDF.CAnnotationStamp(sName, nPageNum, aRect, oPdfDoc);
                break;
            default:
                return null;
        }

        oAnnot.SetCreationDate(sCrDate);
        oAnnot.SetModDate(sModDate);
        oAnnot.SetAuthor(sAuthor);
        oAnnot.SetDisplay(isHidden ? window["AscPDF"].Api.Objects.display["hidden"] : window["AscPDF"].Api.Objects.display["visible"]);
        oAnnot.SetContents(sText);

        return oAnnot;
    }

    function CreateAscAnnotPropFromObj(annot) {
        let oProps = new Asc.asc_CAnnotProperty();

        oProps.asc_putType(annot.GetType());
        
        let oFillColor  = annot.GetFillColor();
        let oFillRGB    = annot.GetRGBColor(oFillColor);
        let oFill       = AscFormat.CreateSolidFillRGBA(oFillRGB.r, oFillRGB.g, oFillRGB.b, 255);
        if (isRealObject(oFill)) {
            oProps.asc_putFill(AscFormat.CreateAscFill(oFill));
        }

        let oStrokeColor    = annot.GetStrokeColor();
        let oStrokeRGB      = annot.GetRGBColor(oStrokeColor);
        let oStrokeFill     = AscFormat.CreateSolidFillRGBA(oStrokeRGB.r, oStrokeRGB.g, oStrokeRGB.b, 255);
        if (isRealObject(oStrokeFill)) {
            oProps.asc_putStroke(AscFormat.CreateAscStroke(oStrokeFill, true));
        }

        if (annot.IsFreeText() && annot.IsInTextBox()) {
            oProps.asc_setCanEditText(true);
        }
        // obj.Position = new Asc.CPosition({X: shapeProp.x, Y: shapeProp.y});

        return oProps;
    }

    function private_correctRGBColorComponent(component) {
        if (typeof(component) != "number") {
            component = 0;
        }

        return component;
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

    function private_computeDrawingAddingPos(nPage, nExtX, nExtY) {
        let oDoc        = Asc.editor.getPDFDoc();
        let oViewRect   = oDoc.Viewer.getViewingRect(nPage);
        let nRotAngle   = oDoc.Viewer.getPageRotate(nPage);
        let nPageW      = oDoc.GetPageWidthMM(nPage);
        let nPageH      = oDoc.GetPageHeightMM(nPage);
        let nPosX;
        let nPosY;

        switch (nRotAngle) {
            case 0:
                nPosX = nPageW * ((oViewRect.x + oViewRect.r) / 2) - nExtX / 2;
                nPosY = nPageH * ((oViewRect.y + oViewRect.b) / 2) - nExtY / 2;
                break;
            case 90:
                nPosX = nPageW * ((oViewRect.y + oViewRect.b) / 2) - nExtX / 2;
                nPosY = nPageH - nPageH * ((oViewRect.x + oViewRect.r) / 2) - nExtY / 2;
                break;
            case 180:
                nPosX = nPageW - nPageW * ((oViewRect.x + oViewRect.r) / 2) - nExtX / 2;
                nPosY = nPageH - nPageH * ((oViewRect.y + oViewRect.b) / 2) - nExtY / 2;
                break;
            case 270:
                nPosX = nPageW - nPageW * ((oViewRect.y + oViewRect.b) / 2) - nExtX / 2;
                nPosY = nPageH * ((oViewRect.x + oViewRect.r) / 2) - nExtY / 2;
                break;
        }

        return {x: nPosX, y: nPosY}
    }
    
    if (!window["AscPDF"])
	    window["AscPDF"] = {};
	
	/**
	 * Speical class for handling the composite input in the pdf-editor
	 * @param textController
	 * @param pdfDocument
	 * @constructor
	 */
	function CPDFCompositeInput(textController, pdfDocument) {
		this.textController = textController;
		this.runInput       = new AscWord.RunCompositeInput(false);
		this.contentState   = textController.GetDocContent().GetSelectionState();
		this.pdfDocument    = pdfDocument;
		this.startPoint     = -1;
	}
	CPDFCompositeInput.begin = function(pdfDocument) {
		if (!pdfDocument)
			return null;
		
		if (pdfDocument.IsNeedSkipHistory() || pdfDocument.Viewer.IsOpenFormsInProgress || pdfDocument.Viewer.IsOpenAnnotsInProgress || AscCommon.History.UndoRedoInProgress)
			return null
		
		let textController = pdfDocument.getTextController();
		if (!textController || !textController.canBeginCompositeInput())
			return null;
		if (textController.IsDrawing() && Asc.editor.isRestrictionView())
			return null;
		
		let compositeInput = new CPDFCompositeInput(textController, pdfDocument);
		compositeInput.createNewHistoryPoint(AscDFH.historydescription_Document_CompositeInput);
		textController.beforeCompositeInput();
		let docContent = textController.GetDocContent();
		let run = docContent.GetCurrentRun();
		if (!run) {
			compositeInput.undoAll();
			return null;
		}
		compositeInput.runInput.begin(run);
		
		compositeInput.startPoint = pdfDocument.GetHistory().Index;
		
		this.inUse = true;
		return compositeInput;
	};
	CPDFCompositeInput.prototype.end = function() {
		let codePoints = this.runInput.getCodePoints();
		this.runInput.end();
		this.runInput = null;
		
		if (this.canSquashChanges()) {
			this.undoAll();
			
			this.textController.GetDocContent().SetSelectionState(this.contentState);
			
			this.pdfDocument.DoAction(function() {
				this.textController.EnterText(codePoints);
			}, AscDFH.historydescription_Document_AddLetter, this);
		}
		
		this.inUse = false;
	};
	CPDFCompositeInput.prototype.add = function(codePoint) {
		this.doAction(function() {
			this.runInput.add(codePoint);
			this.textController.SetNeedRecalc(true);
		}, AscDFH.historydescription_Document_CompositeInputReplace);
	};
	CPDFCompositeInput.prototype.remove = function(count) {
		this.doAction(function() {
			this.runInput.remove(count);
			this.textController.SetNeedRecalc(true);
		}, AscDFH.historydescription_Document_CompositeInputReplace);
	};
	CPDFCompositeInput.prototype.replace = function(codePoints) {
		this.doAction(function() {
			this.runInput.replace(codePoints);
			this.textController.SetNeedRecalc(true);
		}, AscDFH.historydescription_Document_CompositeInputReplace);
	};
	CPDFCompositeInput.prototype.setPos = function(pos) {
		return this.runInput.setPos(pos);
	};
	CPDFCompositeInput.prototype.get = function(pos) {
		return this.runInput.getPos(pos);
	};
	CPDFCompositeInput.prototype.getMaxPos = function() {
		return this.runInput.getLength();
	};
	CPDFCompositeInput.prototype.createNewHistoryPoint = function(description) {
		let localHistory = this.pdfDocument.GetHistory();
		if (!localHistory.IsOn())
			localHistory.TurnOn();
		
		localHistory.Create_NewPoint(description);
		localHistory.SetSourceObjectsToPointPdf(this.textController);
	};
	CPDFCompositeInput.prototype.undoAll = function() {
		if (-1 === this.startPoint)
			return;
		
		let localHistory = this.pdfDocument.GetHistory();
		while (localHistory.Index >= this.startPoint)
			localHistory.Undo();
		
		this.startPoint = -1;
	};
	CPDFCompositeInput.prototype.doAction = function(action, description) {
		this.pdfDocument.DoAction(function() {
			action.bind(this)();
		}, description, this);
	};
	CPDFCompositeInput.prototype.checkState = function() {};
	CPDFCompositeInput.prototype.canSquashChanges = function() {
		let localHistory = this.pdfDocument.GetHistory();
		return (!AscCommon.CollaborativeEditing.Is_Fast() || AscCommon.CollaborativeEditing.Is_SingleUser()) && localHistory.Index >= this.startPoint;
	};
    CPDFDoc.prototype.getCompositeInput = function() {
        if (!this.compositeInput) {
            this.compositeInput = new AscWord.DocumentCompositeInput(this);
        }
        
        return this.compositeInput;
    };

    /**
	 * Converts global coords to page coords.
     * Note: use scaled coordinates like pagePos_ from field, and not original like _origRect from field.
     * @param {Number} x
     * @param {Number} y
     * @param {Number} nPage
     * @param {boolean} [isNotMM = false] - coordinates in millimeters or not 
	 * @typeofeditors ["PDF"]
	 */
    function GetPageCoordsByGlobalCoords(x, y, nPage, isNotMM) {
        // конвертация из глобальных x, y к mm кординатам самой страницы
        let oViewer = editor.getDocumentRenderer();
        var pageObject = oViewer.getPageByCoords(x, y);

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
	 * Converts page (native) coords to global coords.
     * Note: use scaled coordinates like pagePos_ from field, and not original like _origRect from field.
     * @param {Number} x
     * @param {Number} y
     * @param {Number} nPage
	 * @typeofeditors ["PDF"]
	 */
    function GetGlobalCoordsByPageCoords(x, y, nPage) {
        let oViewer = Asc.editor.getDocumentRenderer();
        let oDoc = oViewer.getPDFDoc();
        let oTr = oDoc.pagesTransform[nPage].invert;
        
        let result = {};

        let oPoint = oTr.TransformPoint(x, y);
        result.X = result["X"] = oPoint.x;
        result.Y = result["Y"] = oPoint.y;

        return result;
    }

    /**
     * Corverts page coords (in mm) from one page to another page.
     * @param {Number} x
     * @param {Number} y
     * @param {Number} curPage
     * @param {Number} needPage
	 * @typeofeditors ["PDF"]
     * @returns {Object}
	 */
    function ConvertCoordsToAnotherPage(x, y, curPage, needPage) {
        let oViewer     = Asc.editor.getDocumentRenderer();
        let oFile       = oViewer.file;
        let oDoc        = oViewer.getPDFDoc();
        let oCurPageTr  = oDoc.pagesTransform[curPage].normal.CreateDublicate(); // с помощью этой получаем глобальные координаты
        let oNeedPageTr = oDoc.pagesTransform[needPage].normal.CreateDublicate(); // с помощью этой получаем координаты на странице
        
        let inchCur     = (25.4 / oFile.pages[curPage].Dpi);
        let inchNeed    = (25.4 / oFile.pages[needPage].Dpi);
        AscCommon.global_MatrixTransformer.ScaleAppend(oCurPageTr, inchCur, inchCur);
        AscCommon.global_MatrixTransformer.ScaleAppend(oNeedPageTr, inchNeed, inchNeed);
        oCurPageTr.Invert();

        let oGlobalCoords = oCurPageTr.TransformPoint(x, y);
        
        let oNeedPageCoords = oNeedPageTr.TransformPoint(oGlobalCoords.x, oGlobalCoords.y);
        oNeedPageCoords.X = oNeedPageCoords.x;
        oNeedPageCoords.Y = oNeedPageCoords.y;

        return oNeedPageCoords;
    }

    function DrawingCopyObject(Drawing, X, Y, ExtX, ExtY, ImageUrl) {
        this.Drawing = Drawing;
        this.X = X;
        this.Y = Y;
        this.ExtX = ExtX;
        this.ExtY = ExtY;
        this.ImageUrl = ImageUrl;
    }
    
    DrawingCopyObject.prototype.copy = function (oIdMap) {
    
        var _copy = this.Drawing;
        var oPr = new AscFormat.CCopyObjectProperties();
        oPr.idMap = oIdMap;
        if (this.Drawing) {
            _copy = this.Drawing.copy(oPr);
            if (AscCommon.isRealObject(oIdMap)) {
                oIdMap[this.Drawing.Id] = _copy.Id;
            }
        }
        return new DrawingCopyObject(this.Drawing ? _copy : this.Drawing, this.X, this.Y, this.ExtX, this.ExtY, this.ImageUrl);
    
    };
    
    window["AscPDF"].CPDFDoc                    = CPDFDoc;
    window["AscPDF"].CreateAnnotByProps         = CreateAnnotByProps;
    window["AscPDF"].CreateAscAnnotPropFromObj  = CreateAscAnnotPropFromObj;
    window["AscPDF"].CPDFCompositeInput         = CPDFCompositeInput;
    window["AscPDF"].PDFSelectedContent         = PDFSelectedContent;
    window["AscPDF"].DrawingCopyObject          = DrawingCopyObject;
    window["AscPDF"].AscLockTypeElemPDF         = AscLockTypeElemPDF;
    window["AscPDF"]["GetPageCoordsByGlobalCoords"] = window["AscPDF"].GetPageCoordsByGlobalCoords = GetPageCoordsByGlobalCoords;
    window["AscPDF"]["GetGlobalCoordsByPageCoords"] = window["AscPDF"].GetGlobalCoordsByPageCoords = GetGlobalCoordsByPageCoords;
    window["AscPDF"]["ConvertCoordsToAnotherPage"]  = window["AscPDF"].ConvertCoordsToAnotherPage = ConvertCoordsToAnotherPage;

})();
