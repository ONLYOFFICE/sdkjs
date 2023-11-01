/*
 * (c) Copyright Ascensio System SIA 2010-2023
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

"use strict";

(function(window, document)
{
	// TODO: Пока тут идет наследование от класса asc_docs_api для документов
	//       По логике нужно от этого уйти и сделать наследование от базового класса и добавить тип AscCommon.c_oEditorId.PDF
	// TODO: Возможно стоит перенести инициализацию initDocumentRenderer и тогда не придется в каждом методе проверять
	//       наличие this.DocumentRenderer
	
	/**
	 * @param config
	 * @constructor
	 * @extends {AscCommon.DocumentEditorApi}
	 */
	function PDFEditorApi(config) {
		AscCommon.DocumentEditorApi.call(this, config, AscCommon.c_oEditorId.Word);
		
		this.DocumentRenderer = null;
		this.DocumentType     = 1;
	}
	
	PDFEditorApi.prototype = Object.create(AscCommon.DocumentEditorApi.prototype);
	PDFEditorApi.prototype.constructor = PDFEditorApi;
	
	PDFEditorApi.prototype.openDocument = function(file) {
		let perfStart = performance.now();
		
		this.isOnlyReaderMode     = false;
		this.ServerIdWaitComplete = true;
		
		window["AscViewer"]["baseUrl"] = (typeof document !== 'undefined' && document.currentScript) ? "" : "./../../../../sdkjs/pdf/src/engine/";
		window["AscViewer"]["baseEngineUrl"] = "./../../../../sdkjs/pdf/src/engine/";
		
		// TODO: Возможно стоит перенести инициализацию в
		this.initDocumentRenderer();
		this.DocumentRenderer.open(file.data);
		
		AscCommon.InitBrowserInputContext(this, "id_target_cursor", "id_viewer");
		if (AscCommon.g_inputContext)
			AscCommon.g_inputContext.onResize(this.HtmlElementName);
		
		if (this.isMobileVersion)
			this.WordControl.initEventsMobile();
		
		// destroy unused memory
		let isEditForms = true;
		if (isEditForms == false) {
			AscCommon.pptx_content_writer.BinaryFileWriter = null;
			AscCommon.History.BinaryWriter = null;
		}
		
		this.WordControl.OnResize(true);

		this.FontLoader.LoadDocumentFonts(this.WordControl.m_oDrawingDocument.CheckFontNeeds(), false);

		let perfEnd = performance.now();
		AscCommon.sendClientLog("debug", AscCommon.getClientInfoString("onOpenDocument", perfEnd - perfStart), this);
	};
	PDFEditorApi.prototype.isPdfEditor = function() {
		return true;
	};
	PDFEditorApi.prototype.getLogicDocument = function() {
		return this.getPDFDoc();
	};
	PDFEditorApi.prototype.getDocumentRenderer = function() {
		return this.DocumentRenderer;
	};
	PDFEditorApi.prototype.getPDFDoc = function() {
		if (!this.DocumentRenderer)
			return null;
		
		return this.DocumentRenderer.getPDFDoc();
	};
	PDFEditorApi.prototype.IsNeedDefaultFonts = function() {
		return false;
	};
	PDFEditorApi.prototype["asc_setViewerThumbnailsZoom"] = function(value) {
		if (this.haveThumbnails())
			this.DocumentRenderer.Thumbnails.setZoom(value);
	};
	PDFEditorApi.prototype["asc_setViewerThumbnailsUsePageRect"] = function(value) {
		if (this.haveThumbnails())
			this.DocumentRenderer.Thumbnails.setIsDrawCurrentRect(value);
	};
	PDFEditorApi.prototype["asc_viewerThumbnailsResize"] = function() {
		if (this.haveThumbnails())
			this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.Thumbnails.resize();
	};
	PDFEditorApi.prototype["asc_viewerNavigateTo"] = function(value) {
		if (!this.DocumentRenderer)
			return;
		
		this.DocumentRenderer.navigate(value);
	};
	PDFEditorApi.prototype["asc_setViewerTargetType"] = function(type) {
		if (!this.DocumentRenderer)
			return;
		
		this.DocumentRenderer.setTargetType(type);
	};
	PDFEditorApi.prototype["asc_getPageSize"] = function(pageIndex) {
		if (!this.DocumentRenderer)
			return null;
		
		let page = this.DocumentRenderer.file.pages[pageIndex];
		if (!page)
			return null;

		return {
			"W": 25.4 * page.W / page.Dpi,
			"H": 25.4 * page.H / page.Dpi
		}
	};
	PDFEditorApi.prototype.Undo           = function()
	{
		var oDoc = this.getPDFDoc();
		if (!oDoc)
			return;

		oDoc.DoUndo();
	};
	PDFEditorApi.prototype.Redo           = function()
	{
		var oDoc = this.getPDFDoc();
		if (!oDoc)
			return;

		oDoc.DoRedo();
	};
	PDFEditorApi.prototype.asc_CheckCopy = function(_clipboard /* CClipboardData */, _formats) {
		if (!this.DocumentRenderer)
			return;
		
		var _text_object = (AscCommon.c_oAscClipboardDataFormat.Text & _formats) ? {Text : ""} : null;
		var _html_data;
		let oDoc = this.DocumentRenderer.getPDFDoc();
		var oActiveForm = oDoc.activeForm;
		if (oActiveForm && oActiveForm.content.IsSelectionUse()) {
			let sText = oActiveForm.content.GetSelectedText(true);
			if (!sText)
				return;
			
			_text_object.Text = sText;
			_html_data = "<div><p><span>" + sText + "</span></p></div>";
		}
		else {
			_html_data = this.DocumentRenderer.Copy(_text_object)
		}

		if (AscCommon.c_oAscClipboardDataFormat.Text & _formats)
			_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Text, _text_object.Text);

		if (AscCommon.c_oAscClipboardDataFormat.Html & _formats)
			_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Html, _html_data);
	};
	PDFEditorApi.prototype.asc_SelectionCut = function() {
		if (!this.DocumentRenderer)
			return;
		let oDoc = this.DocumentRenderer.getPDFDoc();
		let oField = oDoc.activeForm;
		if (oField && (oField.GetType() === AscPDF.FIELD_TYPES.text || (oField.GetType() === AscPDF.FIELD_TYPES.combobox && oField.IsEditable()))) {
			if (oField.content.IsSelectionUse()) {
				oField.Remove(-1);
				this.DocumentRenderer._paint();
				this.DocumentRenderer.onUpdateOverlay();
				this.WordControl.m_oDrawingDocument.TargetStart();
				this.WordControl.m_oDrawingDocument.showTarget(true);
				oDoc.UpdateCopyCutState();
			}
		}
	};
	PDFEditorApi.prototype.asc_PasteData = function(_format, data1, data2, text_data, useCurrentPoint, callback, checkLocks) {
		if (!this.DocumentRenderer)
			return;
		
		let oDoc = this.DocumentRenderer.getPDFDoc();
		let oField = oDoc.activeForm;
		let data = text_data || data1;

		if (!data)
			return;

		if (oField && (oField.GetType() != AscPDF.FIELD_TYPES.text || oField.IsMultiline() == false))
			data = data.trim().replace(/[\n\r]/g, ' ');

		if (oField && (oField.GetType() === AscPDF.FIELD_TYPES.text || (oField.GetType() === AscPDF.FIELD_TYPES.combobox && oField.IsEditable()))) {
			let aChars = [];
			for (let i = 0; i < data.length; i++)
				aChars.push(data[i].charCodeAt(0));
			
			oField.EnterText(aChars);
			this.WordControl.m_oDrawingDocument.showTarget(true);
			this.WordControl.m_oDrawingDocument.TargetStart();
			this.DocumentRenderer._paint();
			this.DocumentRenderer.onUpdateOverlay();
			oDoc.UpdateCopyCutState();
		}
	};
	PDFEditorApi.prototype.asc_setAdvancedOptions = function(idOption, option) {
		if (this.advancedOptionsAction !== AscCommon.c_oAscAdvancedOptionsAction.Open
			|| AscCommon.EncryptionWorker.asc_setAdvancedOptions(this, idOption, option)
			|| !this.DocumentRenderer)
			return;
		
		this.DocumentRenderer.open(null, option.asc_getPassword());
	};
	PDFEditorApi.prototype.can_CopyCut = function() {
		if (!this.DocumentRenderer)
			return false;
		
		return this.DocumentRenderer.isCanCopy();
	};
	PDFEditorApi.prototype.startGetDocInfo = function() {
		let renderer = this.DocumentRenderer;
		if (!renderer)
			return;
		
		this.sync_GetDocInfoStartCallback();
		
		this.DocumentRenderer.startStatistics();
		this.DocumentRenderer.onUpdateStatistics(0, 0, 0, 0);
		
		if (this.DocumentRenderer.isFullText)
			this.sync_GetDocInfoEndCallback();
	};
	PDFEditorApi.prototype.stopGetDocInfo = function() {
		this.sync_GetDocInfoStopCallback();
		this.DocumentRenderer.endStatistics();
	};
	PDFEditorApi.prototype.asc_searchEnabled = function(isEnabled) {
		if (!this.DocumentRenderer)
			return;
		
		this.DocumentRenderer.SearchResults.IsSearch = isEnabled;
		this.WordControl.OnUpdateOverlay();
	};
	PDFEditorApi.prototype.asc_findText = function(props, isNext, callback) {
		if (!this.DocumentRenderer)
			return 0;
		
		this.DocumentRenderer.SearchResults.IsSearch = true;

		let isAsync = (true === this.DocumentRenderer.findText(props.GetText().trim(), props.IsMatchCase(), props.IsWholeWords(), isNext, this.sync_setSearchCurrent));
		let result = this.DocumentRenderer.SearchResults.Count;
		
		var CurMatchIdx = 0;
		if (this.DocumentRenderer.SearchResults.CurrentPage === 0) {
			CurMatchIdx = this.DocumentRenderer.SearchResults.Current;
		}
		else {
			// чтобы узнать, под каким номером в списке текущее совпадение
			// нужно посчитать сколько совпадений было до текущего на текущей странице
			for (var nPage = 0; nPage <= this.DocumentRenderer.SearchResults.CurrentPage; nPage++) {
				for (var nMatch = 0; nMatch < this.DocumentRenderer.SearchResults.Pages[nPage].length; nMatch++) {
					if (nPage === this.DocumentRenderer.SearchResults.CurrentPage && nMatch === this.DocumentRenderer.SearchResults.Current)
						break;
					
					CurMatchIdx++;
				}
			}
		}
		
		this.DocumentRenderer.SearchResults.CurMatchIdx = CurMatchIdx;
		
		this.sync_setSearchCurrent(CurMatchIdx, result);
		
		if (!isAsync && callback)
			callback(result);
		
		return result;
	};
	PDFEditorApi.prototype.asc_endFindText = function() {
		if (!this.DocumentRenderer)
			return;
		
		this.DocumentRenderer.file.SearchResults.IsSearch = false;
		this.DocumentRenderer.file.onUpdateOverlay();
	};
	PDFEditorApi.prototype.asc_isSelectSearchingResults = function() {
		if (!this.DocumentRenderer)
			return false;
		
		return this.DocumentRenderer.SearchResults.Show;
	};
	PDFEditorApi.prototype.asc_StartTextAroundSearch = function() {
		if (!this.DocumentRenderer)
			return false;
		
		this.DocumentRenderer.file.startTextAround();
	};
	PDFEditorApi.prototype.asc_SelectSearchElement = function(id) {
		if (!this.DocumentRenderer)
			return false;
		
		this.DocumentRenderer.SelectSearchElement(id);
	};
	PDFEditorApi.prototype.ContentToHTML = function() {
		if (!this.DocumentRenderer)
			return "";
		
		this.DocumentReaderMode = new AscCommon.CDocumentReaderMode();
		
		this.DocumentRenderer.selectAll();
		
		var text_data = {
			data : "",
			pushData : function(format, value) { this.data = value; }
		};
		
		this.asc_CheckCopy(text_data, 2);
		
		this.DocumentRenderer.removeSelection();
		
		return text_data.data;
	};
	PDFEditorApi.prototype.goToPage = function(pageNum) {
		if (!this.DocumentRenderer)
			return;
		
		return this.DocumentRenderer.navigateToPage(pageNum);
	};
	PDFEditorApi.prototype.getCountPages = function() {
		return this.DocumentRenderer ? this.DocumentRenderer.getPagesCount() : 0;
	};
	PDFEditorApi.prototype.getCurrentPage = function() {
		return this.DocumentRenderer ? this.DocumentRenderer.currentPage : 0;
	};
	PDFEditorApi.prototype.asc_getPdfProps = function() {
		return  this.DocumentRenderer ? this.DocumentRenderer.getDocumentInfo() : null;
	};
	PDFEditorApi.prototype.asc_enterText = function(text) {
		if (!this.DocumentRenderer)
			return false;
		
		let viewer	= this.DocumentRenderer;
		let oDoc	= viewer.getPDFDoc();
		if (!viewer
			|| !oDoc.checkDefaultFieldFonts()
			|| !oDoc.activeForm
			|| !oDoc.activeForm.IsEditable()) {
			return false;
		}
		
		oDoc.activeForm.EnterText(text);
		
		this.WordControl.m_oDrawingDocument.TargetStart();
		// Чтобы при зажатой клавише курсор не пропадал
		this.WordControl.m_oDrawingDocument.showTarget(true);
		
		return true;
	};
	PDFEditorApi.prototype.asc_GetSelectedText = function() {
		if (!this.DocumentRenderer)
			return "";

		var textObj = {Text : ""};
		this.DocumentRenderer.Copy(textObj);
		if (textObj.Text.trim() === "")
			return "";
		
		return textObj.Text;
	};
	PDFEditorApi.prototype.SetMarkerFormat          = function(nType, value, opacity, r, g, b)
	{
		this.isMarkerFormat	= value;
		this.curMarkerType	= nType;
		let oDoc			= this.getPDFDoc();
		
		if (value == true && oDoc.activeForm)
			oDoc.OnExitFieldByClick();

		if (this.isMarkerFormat) {
			switch (this.curMarkerType) {
				case AscPDF.ANNOTATIONS_TYPES.Highlight:
					this.SetHighlight(r, g, b, opacity);
					break;
				case AscPDF.ANNOTATIONS_TYPES.Underline:
					this.SetUnderline(r, g, b, opacity);
					break;
				case AscPDF.ANNOTATIONS_TYPES.Strikeout:
					this.SetStrikeout(r, g, b, opacity);
					break;
			}
		}
	};
	PDFEditorApi.prototype.Paste = function()
	{
		if (AscCommon.g_clipboardBase.IsWorking())
			return false;

		return AscCommon.g_clipboardBase.Button_Paste();
	};
	PDFEditorApi.prototype.asc_setSkin = function(theme)
    {
        AscCommon.updateGlobalSkin(theme);

        if (this.isUseNativeViewer)
        {
            if (this.WordControl && this.WordControl.m_oDrawingDocument && this.WordControl.m_oDrawingDocument.m_oDocumentRenderer)
            {
                this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.updateSkin();
            }
        }

        if (this.WordControl && this.WordControl.m_oBody)
        {
            this.WordControl.OnResize(true);
            if (this.WordControl.m_oEditor && this.WordControl.m_oEditor.HtmlElement)
            {
                this.WordControl.m_oEditor.HtmlElement.fullRepaint = true;
                this.WordControl.OnScroll();
            }
        }
    };
	PDFEditorApi.prototype.asc_SelectPDFFormListItem = function(sId) {
		let nIdx = parseInt(sId);
		let oViewer = this.DocumentRenderer;
		let oDoc	= oViewer.getPDFDoc();
		let oField = oDoc.activeForm;
		if (!oField)
			return;
		
		oField.SelectOption(nIdx);
		let isNeedRedraw = oField.IsNeedCommit();
		if (oField._commitOnSelChange && oField.IsNeedCommit()) {
			oField.Commit();
			isNeedRedraw = true;
			
			oDoc.activeForm = null;
			oField.SetDrawHighlight(true);
			
			this.WordControl.m_oDrawingDocument.TargetEnd();
		}
		
		
		if (isNeedRedraw) {
			oViewer._paint();
		}
	};
	PDFEditorApi.prototype.SetDrawingFreeze = function(bIsFreeze)
	{
		if (!this.WordControl)
			return;

		this.WordControl.DrawingFreeze = bIsFreeze;

		var elem = document.getElementById("id_main");
		if (elem)
		{
			if (bIsFreeze)
			{
				elem.style.display = "none";
			}
			else
			{
				elem.style.display = "block";
			}
		}

		if (!bIsFreeze)
			this.WordControl.OnScroll();
	};
	// composite input
	PDFEditorApi.prototype.Begin_CompositeInput = function()
	{
		let viewer = this.DocumentRenderer;
		if (!viewer)
			return false;
		
		let pdfDoc = viewer.getPDFDoc();
		if (!pdfDoc.activeForm || !pdfDoc.activeForm.IsEditable())
			return false;
		
		function begin() {
			pdfDoc.activeForm.beginCompositeInput();
		}
		
		if (!pdfDoc.checkDefaultFieldFonts(begin))
			return true;
		
		begin();
		return true;
	};
	PDFEditorApi.prototype.Add_CompositeText = function(codePoint) {
		let form = this._getActiveForm();
		if (!form || !form.IsEditable())
			return;
		
		form.addCompositeText(codePoint);
	};
	PDFEditorApi.prototype.Remove_CompositeText = function(count) {
		let form = this._getActiveForm();
		if (!form || !form.IsEditable())
			return;
		
		form.removeCompositeText(count);
	};
	PDFEditorApi.prototype.Replace_CompositeText = function(codePoints) {
		let form = this._getActiveForm();
		if (!form || !form.IsEditable())
			return;
		
		form.replaceCompositeText(codePoints);
	};
	PDFEditorApi.prototype.End_CompositeInput = function()
	{
		let form = this._getActiveForm();
		if (!form || !form.IsEditable())
			return;
		
		form.endCompositeInput();
	};
	PDFEditorApi.prototype.Set_CursorPosInCompositeText = function(pos) {
		let form = this._getActiveForm();
		if (!form || !form.IsEditable())
			return;
		
		form.setPosInCompositeInput(pos);
	};
	PDFEditorApi.prototype.Get_CursorPosInCompositeText = function() {
		let form = this._getActiveForm();
		if (!form || !form.IsEditable())
			return 0;
		
		return form.getPosInCompositeInput();
	};
	PDFEditorApi.prototype.Get_MaxCursorPosInCompositeText = function() {
		let form = this._getActiveForm();
		if (!form || !form.IsEditable())
			return 0;
		
		return form.getMaxPosInCompositeInput();
	};
	PDFEditorApi.prototype._getActiveForm = function() {
		let viewer = this.DocumentRenderer;
		if (!viewer)
			return null;
		
		let pdfDoc = viewer.getPDFDoc();
		return pdfDoc.activeForm;
	};


	// for comments
	PDFEditorApi.prototype.can_AddQuotedComment = function()
	{
		return true;
	};
	PDFEditorApi.prototype.asc_addComment = function(AscCommentData)
	{
		var oDoc = this.getPDFDoc();
		let oViewer = editor.getDocumentRenderer();

		if (!oDoc)
			return null;

		let oCommentData = new AscCommon.CCommentData();
		oCommentData.Read_FromAscCommentData(AscCommentData);

		let oComment = oDoc.AddComment(AscCommentData);
		//this.sync_AddComment(oComment.GetId(), oCommentData);

		oComment.AddToRedraw();
		oViewer._paint();

		return oComment.GetId()
	};
	PDFEditorApi.prototype.asc_showComments = function()
	{
		let oDoc = this.getPDFDoc();
		oDoc.HideShowAnnots(false);
		oDoc.Viewer._paint();
        oDoc.Viewer.onUpdateOverlay();
	};

	PDFEditorApi.prototype.asc_hideComments = function()
	{
		let oDoc = this.getPDFDoc();
		oDoc.HideShowAnnots(true);
		oDoc.Viewer._paint();
        oDoc.Viewer.onUpdateOverlay();
	};
	PDFEditorApi.prototype.asc_getAnchorPosition = function()
	{
		let oViewer		= editor.getDocumentRenderer();
		let pageObject	= oViewer.getPageByCoords(AscCommon.global_mouseEvent.X - oViewer.x, AscCommon.global_mouseEvent.Y - oViewer.y);
		let nPage		= pageObject ? pageObject.index : oViewer.currentPage;

		let nScaleY			= oViewer.drawingPages[nPage].H / oViewer.file.pages[nPage].H;
        let nScaleX			= oViewer.drawingPages[nPage].W / oViewer.file.pages[nPage].W;
		let nCommentWidth	= 33 * nScaleX;
		let nCommentHeight	= 33 * nScaleY;
		let oDoc			= oViewer.getPDFDoc();

		if (!pageObject) {
			let oPos = AscPDF.GetGlobalCoordsByPageCoords(10, 10, nPage, true);
			oDoc.anchorPositionToAdd = {
				x: 10,
				y: 10
			};
			return new AscCommon.asc_CRect(oPos["X"] + nCommentWidth, oPos["Y"] + nCommentHeight / 2, 0, 0);
		}

		oDoc.anchorPositionToAdd = {
			x: pageObject.x,
			y: pageObject.y
		};

		if (oDoc.mouseDownAnnot) {
			let aRect = oDoc.mouseDownAnnot.GetRect();
			let oPos = AscPDF.GetGlobalCoordsByPageCoords(aRect[2], aRect[1] + (aRect[3] - aRect[1]) / 2, nPage, true);
			return new AscCommon.asc_CRect(oPos["X"], oPos["Y"], 0, 0);
		}
		
		return new AscCommon.asc_CRect(AscCommon.global_mouseEvent.X - oViewer.x, AscCommon.global_mouseEvent.Y - oViewer.y, 0, 0);
	};
	PDFEditorApi.prototype.asc_removeComment = function(Id)
	{
		let oDoc = this.getPDFDoc();
		if (!oDoc)
			return;

		oDoc.RemoveComment(Id);
	};
	PDFEditorApi.prototype.asc_changeComment = function(Id, AscCommentData)
	{
		var oDoc = this.getDocumentRenderer().getPDFDoc();
		if (!oDoc)
			return;

		var CommentData = new AscCommon.CCommentData();
		CommentData.Read_FromAscCommentData(AscCommentData);
		oDoc.EditComment(Id, CommentData);
	};
	PDFEditorApi.prototype.asc_selectComment = function(Id)
	{
		this.getPDFDoc().GoToAnnot(Id);
	};

	PDFEditorApi.prototype.asc_EditSelectAll = function()
	{
		let oViewer = this.getDocumentRenderer();
		let oDoc = oViewer.getPDFDoc();

		oViewer.file.selectAll();
		oDoc.UpdateCopyCutState();
	};
	PDFEditorApi.prototype.asc_showComment = function(Id)
	{
		if (Id instanceof Array)
			this.getPDFDoc().ShowComment(Id);
		else
			this.getPDFDoc().ShowComment([Id]);
	};
	// drawing pen
	PDFEditorApi.prototype.onInkDrawerChangeState = function() {
		const oViewer	= this.getDocumentRenderer();
		const oDoc		= this.getDocumentRenderer().getPDFDoc();

		if(!oDoc)
			return;

		oViewer.file.Selection = {
			Page1 : 0,
			Line1 : 0,
			Glyph1 : 0,

			Page2 : 0,
			Line2 : 0,
			Glyph2 : 0,

			IsSelection : false
		}

		oViewer.onUpdateOverlay();
		oViewer.DrawingObjects.onInkDrawerChangeState();
		oDoc.currInkInDrawingProcess = null;

		if (false == this.isInkDrawerOn()) {
			if (oViewer.MouseHandObject) {
				oViewer.setCursorType("pointer");
			}
			else {
				oViewer.setCursorType("default");
			}
		}
	};

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	PDFEditorApi.prototype.initDocumentRenderer = function() {
		let documentRenderer = new AscCommon.CViewer(this.HtmlElementName, this);
		
		let _t = this;
		documentRenderer.registerEvent("onNeedPassword", function(){
			_t.sendEvent("asc_onAdvancedOptions", Asc.c_oAscAdvancedOptionsID.DRM);
		});
		documentRenderer.registerEvent("onStructure", function(structure){
			_t.sendEvent("asc_onViewerBookmarksUpdate", structure);
		});
		documentRenderer.registerEvent("onCurrentPageChanged", function(pageNum){
			_t.sendEvent("asc_onCurrentPage", pageNum);
		});
		documentRenderer.registerEvent("onPagesCount", function(pagesCount){
			_t.sendEvent("asc_onCountPages", pagesCount);
		});
		documentRenderer.registerEvent("onZoom", function(value, type){
			_t.WordControl.m_nZoomValue = ((value * 100) + 0.5) >> 0;
			_t.sync_zoomChangeCallback(_t.WordControl.m_nZoomValue, type);
		});
		
		documentRenderer.registerEvent("onFileOpened", function() {
			_t.disableRemoveFonts = true;
			_t.onDocumentContentReady();
			_t.bInit_word_control = true;
			
			var thumbnailsDivId = "thumbnails-list";
			if (document.getElementById(thumbnailsDivId))
			{
				documentRenderer.Thumbnails = new AscCommon.ThumbnailsControl(thumbnailsDivId);
				documentRenderer.setThumbnailsControl(documentRenderer.Thumbnails);
				
				documentRenderer.Thumbnails.registerEvent("onZoomChanged", function (value) {
					_t.sendEvent("asc_onViewerThumbnailsZoomUpdate", value);
				});
			}
			documentRenderer.isDocumentContentReady = true;
		});
		documentRenderer.registerEvent("onHyperlinkClick", function(url){
			_t.sendEvent("asc_onHyperlinkClick", url);
		});
		
		documentRenderer.ImageMap = {};
		documentRenderer.InitDocument = function() {};
		
		this.DocumentRenderer = documentRenderer;
		this.WordControl.m_oDrawingDocument.m_oDocumentRenderer = documentRenderer;
	};
	PDFEditorApi.prototype.haveThumbnails = function() {
		return !!(this.DocumentRenderer && this.DocumentRenderer.Thumbnails);
	};
	PDFEditorApi.prototype.updateDarkMode = function() {
		if (!this.DocumentRenderer)
			return;
		
		this.DocumentRenderer.updateDarkMode();
	};
	PDFEditorApi.prototype.SetHighlight = function(r, g, b, opacity) {
		let oViewer	= this.getDocumentRenderer();
		let oDoc	= this.getPDFDoc();
		oDoc.SetHighlight(r, g, b, opacity);

		oViewer.file.Selection = {
			Page1 : 0,
			Line1 : 0,
			Glyph1 : 0,

			Page2 : 0,
			Line2 : 0,
			Glyph2 : 0,

			IsSelection : false
		}
		
		oViewer._paint();
		oViewer.onUpdateOverlay();
	};
	PDFEditorApi.prototype.SetStrikeout = function(r, g, b, opacity) {
		let oViewer	= this.getDocumentRenderer();
		let oDoc	= this.getPDFDoc();
		oDoc.SetStrikeout(r, g, b, opacity);

		oViewer.file.Selection = {
			Page1 : 0,
			Line1 : 0,
			Glyph1 : 0,

			Page2 : 0,
			Line2 : 0,
			Glyph2 : 0,

			IsSelection : false
		}
		
		oViewer._paint();
		oViewer.onUpdateOverlay();
	};
	PDFEditorApi.prototype.SetUnderline = function(r, g, b, opacity) {
		let oViewer	= this.getDocumentRenderer();
		let oDoc	= this.getPDFDoc();
		oDoc.SetUnderline(r, g, b, opacity);

		oViewer.file.Selection = {
			Page1 : 0,
			Line1 : 0,
			Glyph1 : 0,

			Page2 : 0,
			Line2 : 0,
			Glyph2 : 0,

			IsSelection : false
		}
		
		oViewer._paint();
		oViewer.onUpdateOverlay();
	};
	PDFEditorApi.prototype.updateSkin = function() {
		let obj_id_main = document.getElementById("id_main");
		if (obj_id_main) {
			obj_id_main.style.backgroundColor = AscCommon.GlobalSkin.BackgroundColor;
			document.getElementById("id_viewer").style.backgroundColor = AscCommon.GlobalSkin.BackgroundColor;
			document.getElementById("id_panel_right").style.backgroundColor = AscCommon.GlobalSkin.ScrollBackgroundColor;
			document.getElementById("id_horscrollpanel").style.backgroundColor = AscCommon.GlobalSkin.ScrollBackgroundColor;
		}
		
		if (!this.DocumentRenderer)
			return;
		
		this.DocumentRenderer.updateSkin();
	};
	PDFEditorApi.prototype._selectSearchingResults = function(isShow) {
		if (!this.DocumentRenderer)
			return;
		
		this.DocumentRenderer.SearchResults.Show = isShow;
		this.DocumentRenderer.onUpdateOverlay();
	};
	PDFEditorApi.prototype._printDesktop = function(options) {
		if (!this.DocumentRenderer)
			return false;
		
		let desktopOptions = {};
		if (options && options.advancedOptions)
			desktopOptions["nativeOptions"] = options.advancedOptions.asc_getNativeOptions();
		
		let viewer = this.DocumentRenderer;
		if (window["AscDesktopEditor"] && !window["AscDesktopEditor"]["IsLocalFile"]() && window["AscDesktopEditor"]["SetPdfCloudPrintFileInfo"])
		{
			if (!window["AscDesktopEditor"]["IsCachedPdfCloudPrintFileInfo"]())
				window["AscDesktopEditor"]["SetPdfCloudPrintFileInfo"](AscCommon.Base64.encode(viewer.getFileNativeBinary()));
		}
		window["AscDesktopEditor"]["Print"](JSON.stringify(desktopOptions), viewer.savedPassword ? viewer.savedPassword : "");
		return true;
	};
	PDFEditorApi.prototype.asyncImagesDocumentEndLoaded = function() {
		this.ImageLoader.bIsLoadDocumentFirst = false;
		
		if (!this.DocumentRenderer)
			return;
		
		if (this.EndActionLoadImages === 1) {
			this.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadDocumentImages);
		}
		else if (this.EndActionLoadImages === 2) {
			if (this.isPasteFonts_Images)
				this.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadImage);
			else
				this.sync_EndAction(Asc.c_oAscAsyncActionType.Information, Asc.c_oAscAsyncAction.LoadImage);
		}
		
		this.EndActionLoadImages = 0;
		
		this.WordControl.m_oDrawingDocument.OpenDocument();
		
		this.LoadedObject = null;
		
		this.bInit_word_control = true;
		
		this.WordControl.InitControl();
		
		if (this.isViewMode)
			this.asc_setViewMode(true);
	};
	PDFEditorApi.prototype.Input_UpdatePos = function() {
		if (this.DocumentRenderer)
			this.WordControl.m_oDrawingDocument.MoveTargetInInputContext();
	};
	PDFEditorApi.prototype.OnMouseUp = function(x, y) {
		if (!this.DocumentRenderer)
			return;
		
		this.DocumentRenderer.onMouseUp(x, y);
	};

	// disable drop
	PDFEditorApi.prototype.isEnabledDropTarget = function() {
		return false;
	};
	PDFEditorApi.prototype.checkDocumentTitleFonts = function() {
		// Do not load any fonts
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Export
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	window['Asc']['PDFEditorApi'] = PDFEditorApi;
	AscCommon.PDFEditorApi        = PDFEditorApi;
	
	PDFEditorApi.prototype['asc_setAdvancedOptions']       = PDFEditorApi.prototype.asc_setAdvancedOptions;
	PDFEditorApi.prototype['startGetDocInfo']              = PDFEditorApi.prototype.startGetDocInfo;
	PDFEditorApi.prototype['stopGetDocInfo']               = PDFEditorApi.prototype.stopGetDocInfo;
	PDFEditorApi.prototype['can_CopyCut']                  = PDFEditorApi.prototype.can_CopyCut;
	PDFEditorApi.prototype['asc_searchEnabled']            = PDFEditorApi.prototype.asc_searchEnabled;
	PDFEditorApi.prototype['asc_findText']                 = PDFEditorApi.prototype.asc_findText;
	PDFEditorApi.prototype['asc_endFindText']              = PDFEditorApi.prototype.asc_endFindText;
	PDFEditorApi.prototype['asc_isSelectSearchingResults'] = PDFEditorApi.prototype.asc_isSelectSearchingResults;
	PDFEditorApi.prototype['asc_StartTextAroundSearch']    = PDFEditorApi.prototype.asc_StartTextAroundSearch;
	PDFEditorApi.prototype['asc_SelectSearchElement']      = PDFEditorApi.prototype.asc_SelectSearchElement;
	PDFEditorApi.prototype['ContentToHTML']                = PDFEditorApi.prototype.ContentToHTML;
	PDFEditorApi.prototype['goToPage']                     = PDFEditorApi.prototype.goToPage;
	PDFEditorApi.prototype['getCountPages']                = PDFEditorApi.prototype.getCountPages;
	PDFEditorApi.prototype['getCurrentPage']               = PDFEditorApi.prototype.getCurrentPage;
	PDFEditorApi.prototype['asc_getPdfProps']              = PDFEditorApi.prototype.asc_getPdfProps;
	PDFEditorApi.prototype['asc_enterText']                = PDFEditorApi.prototype.asc_enterText;
	PDFEditorApi.prototype['asc_GetSelectedText']          = PDFEditorApi.prototype.asc_GetSelectedText;
	PDFEditorApi.prototype['asc_SelectPDFFormListItem']    = PDFEditorApi.prototype.asc_SelectPDFFormListItem;

	PDFEditorApi.prototype['SetDrawingFreeze']             = PDFEditorApi.prototype.SetDrawingFreeze;
	PDFEditorApi.prototype['OnMouseUp']                    = PDFEditorApi.prototype.OnMouseUp;

	PDFEditorApi.prototype['asc_addComment']               = PDFEditorApi.prototype.asc_addComment;
	PDFEditorApi.prototype['can_AddQuotedComment']         = PDFEditorApi.prototype.can_AddQuotedComment;
	PDFEditorApi.prototype['asc_showComments']             = PDFEditorApi.prototype.asc_showComments;
	PDFEditorApi.prototype['asc_showComment']              = PDFEditorApi.prototype.asc_showComment;
	PDFEditorApi.prototype['asc_hideComments']             = PDFEditorApi.prototype.asc_hideComments;
	PDFEditorApi.prototype['asc_removeComment']            = PDFEditorApi.prototype.asc_removeComment;
	PDFEditorApi.prototype['asc_changeComment']            = PDFEditorApi.prototype.asc_changeComment;
	PDFEditorApi.prototype['asc_selectComment']            = PDFEditorApi.prototype.asc_selectComment;

	PDFEditorApi.prototype['asc_setSkin']                  = PDFEditorApi.prototype.asc_setSkin;
	PDFEditorApi.prototype['asc_getAnchorPosition']        = PDFEditorApi.prototype.asc_getAnchorPosition;
	PDFEditorApi.prototype['SetMarkerFormat']              = PDFEditorApi.prototype.SetMarkerFormat;
	PDFEditorApi.prototype['asc_EditSelectAll']            = PDFEditorApi.prototype.asc_EditSelectAll;
	PDFEditorApi.prototype['Undo']                         = PDFEditorApi.prototype.Undo;
	PDFEditorApi.prototype['Redo']                         = PDFEditorApi.prototype.Redo;
	PDFEditorApi.prototype['asc_SelectionCut']             = PDFEditorApi.prototype.asc_SelectionCut;
	PDFEditorApi.prototype['asc_CheckCopy']                = PDFEditorApi.prototype.asc_CheckCopy;
	PDFEditorApi.prototype['Paste']                        = PDFEditorApi.prototype.Paste;
	PDFEditorApi.prototype['asc_PasteData']                = PDFEditorApi.prototype.asc_PasteData;

})(window, window.document);
