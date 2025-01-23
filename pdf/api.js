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
		
		this.compositeInput = null;
		this.isPdfViewer    = false; // Было решено, что флаг isViewMode присылается всегда false, т.к. пдф всегда
		                             // можно редактировать (во вьювере заполнять поля, например)
		                             // Данный флаг различает в каком режиме загружен документ (edit/view)
	}
	
	PDFEditorApi.prototype = Object.create(AscCommon.DocumentEditorApi.prototype);
	PDFEditorApi.prototype.constructor = PDFEditorApi;
	
	PDFEditorApi.prototype.openDocument = function(file) {
		let perfStart = performance.now();
		
		this.isOnlyReaderMode = false;
		
		window["AscViewer"]["baseUrl"] = (typeof document !== 'undefined' && document.currentScript) ? "" : "./../../../../sdkjs/pdf/src/engine/";
		window["AscViewer"]["baseEngineUrl"] = "./../../../../sdkjs/pdf/src/engine/";
		
		// TODO: Возможно стоит перенести инициализацию в
		this.initDocumentRenderer();
		this.DocumentRenderer.open(file.data);
		
		AscCommon.InitBrowserInputContext(this, "id_target_cursor", "id_viewer");
		if (AscCommon.g_inputContext)
			AscCommon.g_inputContext.onResize(this.HtmlElementName);
		
		this.WordControl.initEventsMobile();
		this.DocumentRenderer.touchManager = this.WordControl.MobileTouchManager;

		if (undefined !== this.startMobileOffset)
		{
			this.WordControl.setOffsetTop(this.startMobileOffset.offset, this.startMobileOffset.offsetScrollTop);
			delete this.startMobileOffset;
		}
		
		// destroy unused memory
		let isEditForms = true;
		if (isEditForms == false) {
			AscCommon.pptx_content_writer.BinaryFileWriter = null;
			AscCommon.History.BinaryWriter = null;
		}
		
		this.WordControl.OnResize(true);
		this.FontLoader.LoadDocumentFonts();

		let perfEnd = performance.now();
		AscCommon.sendClientLog("debug", AscCommon.getClientInfoString("onOpenDocument", perfEnd - perfStart), this);
	};
	PDFEditorApi.prototype.isPdfEditor = function() {
		return true;
	};
	PDFEditorApi.prototype.getLogicDocument = function() {
		return this.getPDFDoc();
	};
	PDFEditorApi.prototype.getDrawingDocument = function () {
		return this.WordControl.m_oDrawingDocument;
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
	PDFEditorApi.prototype.AddTextArt = function(nStyle) {
		let oDoc = this.getPDFDoc();
		oDoc.DoAction(function() {
			oDoc.AddTextArt(nStyle, this.getDocumentRenderer().currentPage);
		}, AscDFH.historydescription_Document_AddTextArt, this);
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
	PDFEditorApi.prototype["asc_nativeOpenFile"] = function(data, version) {
		this.SpellCheckUrl = '';

		this.User = new AscCommon.asc_CUser();
		this.User.setId("TM");
		this.User.setUserName("native");

		this.WordControl.m_bIsRuler = false;
		this.WordControl.Init();

		this.initDocumentRenderer();
		this.DocumentType   = 2;

		AscCommon.g_oIdCounter.Set_Load(true);

		if (undefined !== version)
			AscCommon.CurFileVersion = version;

		let viewer = this.DocumentRenderer;
		viewer.file = window["AscViewer"].createFile(data);

		if (viewer.file)
		{
			viewer.SearchResults = viewer.file.SearchResults;
			viewer.file.viewer = viewer;
		}

		viewer.afterOpen();

		AscCommon.g_oIdCounter.Set_Load(false);
		this.LoadedObject = 1;

		if (window["NATIVE_EDITOR_ENJINE"] === true && undefined != window["native"])
		{
			AscCommon.CDocsCoApi.prototype.askSaveChanges = function(callback)
			{
				callback({"saveLock" : false});
			};
			AscCommon.CDocsCoApi.prototype.saveChanges    = function(arrayChanges, deleteIndex, excelAdditionalInfo)
			{
				if (window["native"]["SaveChanges"])
					window["native"]["SaveChanges"](arrayChanges.join("\",\""), deleteIndex, arrayChanges.length);
			};
		}
	};
	PDFEditorApi.prototype["asc_nativeCalculateFile"] = function() {
		let pdfDoc = this.getPDFDoc();
		if (!pdfDoc)
			return;

		pdfDoc.RecalculateAll();
	};
	PDFEditorApi.prototype["asc_nativePrintPagesCount"] = function()
	{
		// the result is unimportant
		return 0;
	};
	PDFEditorApi.prototype["asc_nativeGetPDF"] = function(options)
	{
		if (this.DocumentRenderer)
		{
			let result = this.DocumentRenderer.Save();
			if (!result)
				return null;
			window["native"]["Save_End"]("", result.length);
			return result;
		}
		return null;
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

		let oDoc			= this.getPDFDoc();
		let oFile			= oDoc.Viewer.file;
		let oActiveForm		= oDoc.activeForm;
		let oActiveAnnot	= oDoc.mouseDownAnnot;
		let oActiveDrawing	= oDoc.activeDrawing;

		if (oActiveForm && oActiveForm.content.IsSelectionUse()) {
			let sText = oActiveForm.content.GetSelectedText(false, {NewLine: true});
			if (!sText)
				return;

			if (AscCommon.c_oAscClipboardDataFormat.Text & _formats)
				_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Text, sText);

			if (AscCommon.c_oAscClipboardDataFormat.Html & _formats)
				_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Html, "<div><p><span>" + sText + "</span></p></div>");
		}
		else if (oActiveAnnot && oActiveAnnot.IsFreeText() && oActiveAnnot.IsInTextBox()) {
			let sText = oActiveAnnot.GetDocContent().GetSelectedText(false, {NewLine: true});
			if (!sText)
				return;

			if (AscCommon.c_oAscClipboardDataFormat.Text & _formats)
				_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Text, sText);

			if (AscCommon.c_oAscClipboardDataFormat.Html & _formats)
				_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Html, "<div><p><span>" + sText + "</span></p></div>");
		}
		else if (oActiveDrawing) {
			let oContent = oActiveDrawing.GetDocContent();
			let sText = oContent ? oContent.GetSelectedText(false) : "";
			let _data, sBase64;

			// text
			if (AscCommon.c_oAscClipboardDataFormat.Text & _formats) {
				_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Text, sText);
			}
			//HTML
			if (AscCommon.c_oAscClipboardDataFormat.Html & _formats) {
				var oCopyProcessor = new AscCommon.CopyProcessor(this);
				sBase64            = oCopyProcessor.Start();
				_data              = oCopyProcessor.getInnerHtml();
	
				_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Html, _data)
			}
			//INTERNAL
			if (AscCommon.c_oAscClipboardDataFormat.Internal & _formats) {
				if (sBase64 === null)
				{
					var oCopyProcessor = new AscCommon.CopyProcessor(this);
					sBase64            = oCopyProcessor.Start();
				}
	
				_data = sBase64;
				_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Internal, _data)
			}
		}
		else {
			let _text_object = {Text: ""};
			let _html_data = this.DocumentRenderer.Copy(_text_object);

			if (AscCommon.c_oAscClipboardDataFormat.Text & _formats)
				_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Text, _text_object.Text);

			if (AscCommon.c_oAscClipboardDataFormat.Html & _formats)
				_clipboard.pushData(AscCommon.c_oAscClipboardDataFormat.Html, _html_data);
		}
	};
	PDFEditorApi.prototype.asc_SelectionCut = function() {
		if (!this.DocumentRenderer)
			return;
		
		let oDoc = this.DocumentRenderer.getPDFDoc();
		if (oDoc.CanCopyCut().cut) {
			oDoc.DoAction(function() {
				oDoc.Remove(1);
			}, AscDFH.historydescription_Cut);
		}
	};
	PDFEditorApi.prototype.onUpdateRestrictions = function() {
		let oDoc = this.getPDFDoc();
		if (!oDoc) {
			return;
		}

		if (this.isRestrictionView()) {
			let oActiveObj = oDoc.GetActiveObject();

			if (oActiveObj && oActiveObj.IsDrawing()) {
				oDoc.BlurActiveObject();
			}

			oDoc.drawings.forEach(function(drawing) {
				drawing.AddToRedraw();
			});
		}
		else {
			setTimeout(function() {
				oDoc.checkDefaultFonts();
			});
		}
	};
	PDFEditorApi.prototype.sync_CanUndoCallback = function(canUndo) {
		this.sendEvent("asc_onCanUndo", canUndo);
	};
	PDFEditorApi.prototype.sync_CanRedoCallback = function(canRedo) {
		if (true === AscCommon.CollaborativeEditing.Is_Fast() && true !== AscCommon.CollaborativeEditing.Is_SingleUser()) {
			canRedo = false;
		}

		this.sendEvent("asc_onCanRedo", canRedo);
	};
	PDFEditorApi.prototype.asc_PasteData = function(_format, data1, data2, text_data, useCurrentPoint, callback, checkLocks) {
		if (!this.DocumentRenderer)
			return;
		
		let oDoc			= this.DocumentRenderer.getPDFDoc();
		let data			= typeof(text_data) == "string" ? text_data : data1;
		let oActiveDrawing	= oDoc.activeDrawing;

		oDoc.StartAction(AscDFH.historydescription_Document_PasteHotKey);
		
		this.needPasteText = false; // если не вставили бинарник, то вставляем текст
		// пока что копирование бинарником только внутри drawings или самих drawings
		if ([AscCommon.c_oAscClipboardDataFormat.Internal, AscCommon.c_oAscClipboardDataFormat.HtmlElement, AscCommon.c_oAscClipboardDataFormat.Text].includes(_format) && ((oDoc.GetActiveObject() == null) || oActiveDrawing)) {
			if (this.isRestrictionView()) {
				oDoc.FinalizeAction(true)
				return;
			}

			window['AscCommon'].g_specialPasteHelper.Paste_Process_Start(arguments[5]);
			AscCommon.Editor_Paste_Exec(this, _format, data1, data2, text_data, undefined, callback);
		}
		else {
			this.needPasteText = true;
		}
		
		if (!this.needPasteText || typeof(data) != "string")
			return;

		AscFonts.FontPickerByCharacter.checkText(data, this, processPaste);

		function processPaste() {
			let aChars = [];
			for (let i = 0; i < data.length; i++)
				aChars.push(data[i].charCodeAt(0));

			this.asc_enterText(aChars, true);
		}
	};
	PDFEditorApi.prototype.asc_setAdvancedOptions = function(idOption, option) {
		if (this.advancedOptionsAction !== AscCommon.c_oAscAdvancedOptionsAction.Open
			|| AscCommon.EncryptionWorker.asc_setAdvancedOptions(this, idOption, option)
			|| !this.DocumentRenderer)
			return;
		
		this.DocumentRenderer.open(null, option.asc_getPassword());
	};
	PDFEditorApi.prototype.getGraphicController = function () {
		let oDoc = this.getPDFDoc();
		return oDoc.GetController();
	};
	PDFEditorApi.prototype.can_CopyCut = function() {
		if (!this.DocumentRenderer)
			return false;
		
		let oDoc = this.DocumentRenderer.getPDFDoc();
		if (!oDoc)
			return false;

		return oDoc.CanCopyCut().copy;
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
	PDFEditorApi.prototype.asc_findText = function(props, isNext) {
		let oViewer 		= this.getDocumentRenderer();
		let oDoc			= this.getPDFDoc();
		let oSearchEngine	= oDoc.SearchEngine;

		if (!oViewer)
			return 0;
		
		oViewer.IsSearch = true;

		let isAsync	= (true === oViewer.findText(props, isNext));
		let result	= oSearchEngine.Count;
		
		return result;
	};
	PDFEditorApi.prototype.asc_endFindText = function() {
		if (!this.DocumentRenderer)
			return;
		
		this.DocumentRenderer.IsSearch = false;
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
		
		let oDoc = this.getPDFDoc();
		oDoc.SearchEngine.StartTextAround();
	};
	PDFEditorApi.prototype.asc_SelectSearchElement = function(id) {
		if (!this.DocumentRenderer)
			return false;
		
		this.getPDFDoc().SelectSearchElement(id);
		this.DocumentRenderer.onUpdateOverlay();
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
	PDFEditorApi.prototype.asc_enterText = function(codePoints, isFromPaste) {
		if (!this.DocumentRenderer)
			return false;
		
		let viewer = this.DocumentRenderer;
		let doc    = viewer.getPDFDoc();
		let drDoc  = doc.GetDrawingDocument();
		let drController = doc.GetController();
		
		let textController = doc.getTextController();
		if (!textController)
			return false;
		if (textController.IsDrawing() && Asc.editor.isRestrictionView()) {
			return false;
		}
		
		if (doc.IsSelectionLocked(AscCommon.changestype_Drawing_Props)) {
			return false;
		}
		
		if (true != isFromPaste) {
			doc.StartAction(AscDFH.historydescription_Document_AddLetter);
		}

		let docContent = textController.GetDocContent();
		let result = textController.EnterText(codePoints);
		
		if (null == drController.getTargetTextObject() && false == textController.IsForm()) {
			if (textController.IsAnnot() && textController.IsFreeText()) {
				drController.selection.groupSelection = textController;
				textController.selection.textSelection = textController.GetTextBoxShape()
			}
			else {
				drController.selection.textSelection = textController;
			}
		}

		drDoc.TargetStart(true);
		
		if (docContent.IsSelectionUse() && !docContent.IsSelectionEmpty())
			drDoc.TargetEnd();
		
		if (true !== isFromPaste) {
			doc.FinalizeAction();
		}

		return result;
	};
	PDFEditorApi.prototype.asc_createSmartArt = function (nSmartArtType, oPlaceholderObject) {
		let oViewer	= this.getDocumentRenderer();
		let oDoc	= this.getPDFDoc();

		AscCommon.g_oBinarySmartArts.checkLoadDrawing().then(function()
		{
			return AscCommon.g_oBinarySmartArts.checkLoadData(nSmartArtType);
		}).then(function()
		{
			let oSmartArt;

			oDoc.DoAction(function() {
				oSmartArt = oDoc.AddSmartArt(nSmartArtType, oPlaceholderObject, oViewer.currentPage);
			}, AscDFH.historydescription_Document_AddSmartArt);

			return oSmartArt;
		});
	};
	PDFEditorApi.prototype.asc_undoAllChanges = function() {};
	PDFEditorApi.prototype.asc_addChartDrawingObject = function(chartBinary, Placeholder) {
		let oDoc	= this.getPDFDoc();
		let oViewer	= this.getDocumentRenderer();

		oDoc.DoAction(function() {
			AscFonts.IsCheckSymbols = true;
			oDoc.AddChartByBinary(chartBinary, true, Placeholder, oViewer.currentPage);
			AscFonts.IsCheckSymbols = false;
		}, AscDFH.historydescription_Document_AddChart);
	};
	PDFEditorApi.prototype.asc_correctEnterText = function(oldValue, newValue) {
		
		if (!this.DocumentRenderer)
			return false;
		
		let viewer = this.DocumentRenderer;
		let doc    = viewer.getPDFDoc();
		
		return doc.DoAction(function() {
			let textController = doc.getTextController();
			if (!textController)
				return false;
			
			return textController.CorrectEnterText(oldValue, newValue);
		}, AscDFH.historydescription_Document_AddLetter, this);
	};
	PDFEditorApi.prototype.asc_EditPage = function() {
		let oViewer	= this.getDocumentRenderer();
		let oDoc	= this.getPDFDoc();
		
		oDoc.EditPage(oViewer.currentPage);
	};
	PDFEditorApi.prototype.asc_AddPage = function(bBefore) {
		let oViewer = this.getDocumentRenderer();
		let oFile	= oViewer.file;
		let oDoc 	= this.getPDFDoc();

		let oThumbnails = oViewer.thumbnails;
		let aIndexes = oThumbnails.getSelectedPages();
		let nPos = bBefore ? Math.min.apply(null, aIndexes) : Math.max.apply(null, aIndexes) + 1;

		oDoc.DoAction(function() {
			let oPageToClone = bBefore ? oFile.pages[oViewer.currentPage] : (oFile.pages[oViewer.currentPage + 1] || oFile.pages[oViewer.currentPage]);
			let oPage = {
				fonts: [],
				Rotate: 0,
				Dpi: oPageToClone.Dpi,
				W: oPageToClone.W,
				H: oPageToClone.H
			}

			oDoc.AddPage(nPos, oPage);
			oDoc.FinalizeAction();
			
			oViewer.navigateToPage(nPos);

		}, AscDFH.historydescription_Pdf_AddPage, this);
	};
	PDFEditorApi.prototype.asc_RemovePage = function(aPages) {
		let oViewer = this.getDocumentRenderer();
		let oDoc 	= this.getPDFDoc();

		let oThumbnails = oViewer.thumbnails;
		aPages = aPages != undefined ? aPages : oThumbnails.getSelectedPages().slice();

		oDoc.DoAction(function() {
			oDoc.RemovePages(aPages);
        }, AscDFH.historydescription_Pdf_RemovePage, this, aPages);
	};
	PDFEditorApi.prototype.asc_GetSelectedText = function(bClearText, select_Pr) {
		if (!this.DocumentRenderer)
			return "";
		
		bClearText = (bClearText === true);
		
		let oDoc		= this.getPDFDoc();
		let oTxObject	= oDoc.getTextController();
		let textObj		= {Text : ""};
		
		if (oTxObject) {
			let oContent = oTxObject.GetDocContent();
			textObj.Text = oContent ? (oContent.GetSelectedText(bClearText, select_Pr) || "") : "";
		}
		else {
			this.DocumentRenderer.Copy(textObj);
		}
		
		if (!textObj.Text || textObj.Text.trim() === "")
			return "";
		
		return textObj.Text;
	};
	PDFEditorApi.prototype.asc_AddMath2 = function(Type) {
		let oDoc	= this.getPDFDoc();
		let oTextPr	= oDoc.GetDirectTextPr();

		let oMathElement = new AscCommonWord.MathMenu(Type, oTextPr ? oTextPr.Copy() : null);
		oDoc.AddToParagraph(oMathElement, false);
	};
	PDFEditorApi.prototype.asc_ConvertMathView = function(isToLinear, isAll)
	{
		let oDoc = this.getPDFDoc();
		oDoc.ConvertMathView(isToLinear, isAll);
	};
	PDFEditorApi.prototype.sync_shapePropCallback = function(pr) {
		let oDoc		= this.getPDFDoc();
		let oDrDoc		= oDoc.GetDrawingDocument();
		let oController	= oDoc.GetController();

		var obj = AscFormat.CreateAscShapePropFromProp(pr);
		if (pr.fill != null && pr.fill.fill != null && pr.fill.fill.type == Asc.c_oAscFill.FILL_TYPE_BLIP) {
			oDrDoc.DrawImageTextureFillShape(pr.fill.fill.RasterImageId);
		}
		
		obj.asc_setCanEditText(oController.canEditText());

		var oTextArtProperties = pr.textArtProperties;
		if (oTextArtProperties && oTextArtProperties.Fill && oTextArtProperties.Fill.fill && oTextArtProperties.Fill.fill.type == Asc.c_oAscFill.FILL_TYPE_BLIP) {
			oDrDoc.DrawImageTextureFillTextArt(oTextArtProperties.Fill.fill.RasterImageId);
		}
		else {
			oDrDoc.DrawImageTextureFillTextArt(null);
		}
		
		var _len = this.SelectedObjectsStack.length;
		if (_len > 0) {
			if (this.SelectedObjectsStack[_len - 1].Type == Asc.c_oAscTypeSelectElement.Shape) {
				this.SelectedObjectsStack[_len - 1].Value = obj;
				return;
			}
		}

		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new AscCommon.asc_CSelectedObject(Asc.c_oAscTypeSelectElement.Shape, obj);
	};
	PDFEditorApi.prototype.sync_annotPropCallback = function(annot) {
		var obj = AscPDF.CreateAscAnnotPropFromObj(annot);

		var _len = this.SelectedObjectsStack.length;
		if (_len > 0) {
			if (this.SelectedObjectsStack[_len - 1].Type == Asc.c_oAscTypeSelectElement.Annot) {
				this.SelectedObjectsStack[_len - 1].Value = obj;
				return;
			}
		}

		this.SelectedObjectsStack[this.SelectedObjectsStack.length] = new AscCommon.asc_CSelectedObject(Asc.c_oAscTypeSelectElement.Annot, obj);
	};
	PDFEditorApi.prototype.canUnGroup = function() {
		return false;
	};
	PDFEditorApi.prototype.canGroup = function() {
		return false;
	};
	PDFEditorApi.prototype.shapes_bringToFront = function() {
		this.getPDFDoc().BringToFront();
	};
	PDFEditorApi.prototype.shapes_bringForward = function() {
		this.getPDFDoc().BringForward();
	};
	PDFEditorApi.prototype.shapes_bringToBack = function() {
		this.getPDFDoc().SendToBack();
	};
	PDFEditorApi.prototype.shapes_bringBackward = function() {
		this.getPDFDoc().BringBackward();
	};
	PDFEditorApi.prototype.AddImageUrlAction = function(url, imgProp, obj) {
		var _image = this.ImageLoader.LoadImage(url, 1);
		if (null != _image) {
			this.AddImageUrlActionCallback(_image, obj);
		}
		else {
			this.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadImage);
			this.asyncImageEndLoaded2 = function(_image)
			{
				this.AddImageUrlActionCallback(_image, obj);
				this.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadImage);

				this.asyncImageEndLoaded2 = null;
			}
		}
	};
	PDFEditorApi.prototype.AddImageUrlActionCallback = function(_image, obj) {
		let oDoc = this.getPDFDoc();

		let src = _image.src;
		if (obj && obj.isShapeImageChangeUrl) {
			let AscShapeProp       = new Asc.asc_CShapeProperty();
			AscShapeProp.fill      = new Asc.asc_CShapeFill();
			AscShapeProp.fill.type = Asc.c_oAscFill.FILL_TYPE_BLIP;
			AscShapeProp.fill.fill = new Asc.asc_CFillBlip();
			AscShapeProp.fill.fill.asc_putUrl(src);
			if(obj.textureType !== null && obj.textureType !== undefined){
                AscShapeProp.fill.fill.asc_putType(obj.textureType);
			}
			this.ShapeApply(AscShapeProp);
		}
		else if (obj && obj.isImageChangeUrl) {
			let AscImageProp      = new Asc.asc_CImgProperty();
			AscImageProp.ImageUrl = src;
			this.ImgApply(AscImageProp);
		}
		else if (obj && obj.isTextArtChangeUrl) {
			let AscShapeProp = new Asc.asc_CShapeProperty();
			let oFill        = new Asc.asc_CShapeFill();
			oFill.type       = Asc.c_oAscFill.FILL_TYPE_BLIP;
			oFill.fill       = new Asc.asc_CFillBlip();
			oFill.fill.asc_putUrl(src);
            if(obj.textureType !== null && obj.textureType !== undefined){
				oFill.fill.asc_putType(obj.textureType);
            }
			AscShapeProp.textArtProperties = new Asc.asc_TextArtProperties();
			AscShapeProp.textArtProperties.asc_putFill(oFill);
			this.ShapeApply(AscShapeProp);
		}
		else {
			let srcLocal = AscCommon.g_oDocumentUrls.getImageLocal(src);
			if (srcLocal) {
				src = srcLocal;
			}

			oDoc.AddImages([_image], obj);
		}
	};
	PDFEditorApi.prototype.asc_addImage = function(obj){
		if (this.isEditOleMode){
			this.oSaveObjectForAddImage = obj;
			this.sendFromFrameToGeneralEditor({
				"type": AscCommon.c_oAscFrameDataType.ShowImageDialogInFrame,
			});
			return;
		}

		var t = this;
        if (this.WordControl) // после показа диалога может не прийти mouseUp
        	this.WordControl.m_bIsMouseLock = false;
		
		AscCommon.ShowImageFileDialog(this.documentId, this.documentUserId, this.CoAuthoringApi.get_jwt(), this.documentShardKey, this.documentWopiSrc, this.documentUserSessionId, function(error, files)
		{
			// ошибка может быть объектом в случае отмены добавления картинки в форму
			if (typeof(error) == "object")
				return;

			t._uploadCallback(error, files, obj);
		},
		function(error) {
			if (Asc.c_oAscError.ID.No !== error){
				t.sendEvent("asc_onError", error, Asc.c_oAscError.Level.NoCritical);
			}

			if (obj && obj.sendUrlsToFrameEditor && t.isOpenedChartFrame) {
				t.sendStartUploadImageActionToFrameEditor();
			}
			
			obj && obj.fStartUploadImageCallback && obj.fStartUploadImageCallback();
			t.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.UploadImage);
		});
	};
	PDFEditorApi.prototype.ChangeArtImageFromFile = function(type) {
		this.asc_addImage({isTextArtChangeUrl: true, textureType: type});
	};
	PDFEditorApi.prototype.SetInterfaceDrawImagePlaceTextArt = function(div_id) {
		let oDrDoc = this.getPDFDoc().GetDrawingDocument();
		
		if (!this.isLoadFullApi) {
			this.tmpTextArtDiv = div_id;
			return;
		}

		oDrDoc.InitGuiCanvasTextArt(div_id);
	};
	PDFEditorApi.prototype.asc_setInterfaceDrawImagePlaceShape = function(div_id) {
		let oDrDoc = this.getPDFDoc().GetDrawingDocument();

		if (!this.isLoadFullApi) {
			this.shapeElementId = div_id;
			return;
		}

		oDrDoc.InitGuiCanvasShape(div_id);
	};
	PDFEditorApi.prototype.remove_Hyperlink = function() {
		let oDoc = this.getPDFDoc();
		oDoc.RemoveHyperlink();
	};
	PDFEditorApi.prototype.change_Hyperlink = function(HyperProps) {
		let oDoc = this.getPDFDoc();
		oDoc.ModifyHyperlink(HyperProps);
	};
	PDFEditorApi.prototype.sync_HyperlinkClickCallback = function(Url) {
		if (AscCommon.IsLinkPPAction(Url)) {
			let oViewer		= this.getDocumentRenderer();
			let oDoc		= this.getPDFDoc();
			let nCurPage	= oDoc.GetCurPage();
			let nPagesCount	= oDoc.GetPagesCount();
			
			if (Url == "ppaction://hlinkshowjump?jump=firstslide") {
				oViewer.navigateToPage(0);
			}
			else if (Url == "ppaction://hlinkshowjump?jump=lastslide") {
				oViewer.navigateToPage(nPagesCount - 1);
			}
			else if (Url == "ppaction://hlinkshowjump?jump=nextslide") {
				oViewer.navigateToPage(nCurPage + 1);
			}
			else if (Url == "ppaction://hlinkshowjump?jump=previousslide") {
				oViewer.navigateToPage(nCurPage - 1);
			}
			else {
				let mask	= "ppaction://hlinksldjumpslide";
				let posStr	= Url.indexOf(mask);

				if (0 == posStr) {
					let pageNum = parseInt(Url.substring(mask.length));
					if (pageNum >= 0 && pageNum < nPagesCount)
						oViewer.navigateToPage(pageNum);
				}
			}
			return;
		}

		this.sendEvent("asc_onHyperlinkClick", Url);
	};

	PDFEditorApi.prototype.sync_VerticalTextAlign = function(align) {
		this.sendEvent("asc_onVerticalTextAlign", align);
	};
	PDFEditorApi.prototype.sync_Vert = function(vert) {
		this.sendEvent("asc_onVert", vert);
	};
	PDFEditorApi.prototype.asc_getHeaderFooterProperties = function() {
		return null;
	};
	PDFEditorApi.prototype.SetMarkerFormat = function(nType, value, opacity, r, g, b) {
		this.isMarkerFormat	= value;

		// from edit mode
		if (value == true && nType == undefined) {
			this.getPDFDoc().SetHighlight(r, g, b, opacity);
			return;
		}

		this.curMarkerType	= value ? nType : undefined;
		let oDoc			= this.getPDFDoc();
		let oActiveObj		= oDoc.GetActiveObject();
		let oViewer			= oDoc.Viewer;
		let oDrDoc			= oDoc.GetDrawingDocument();
		
		if (value == true) {
			if (oActiveObj) {
				oDoc.BlurActiveObject();
				oDoc.UpdateInterface();
			}
		}
		else {
			oViewer.onUpdateOverlay();
		}

		if (this.isMarkerFormat) {
			let aSelQuads = oViewer.file.getSelectionQuads();
        	if (aSelQuads.length == 0) {
				oDoc.bOffMarkerAfterUsing = false;
			}
			else {
				oDoc.bOffMarkerAfterUsing = true;
			}

			oDoc.DoAction(function() {
				oDrDoc.LockCursorType(AscCommon.Cursors.MarkerFormat);

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
			}, AscDFH.historydescription_Pdf_AddHighlightAnnot, this);
		}
		else {
			// SetMarkerFormat вызывается при включении ластика/рисовалки, курсор не сбрасываем
			if (false == this.isDrawInkMode() && false == this.isEraseInkMode()) {
				oDrDoc.UnlockCursorType();
				oViewer.setCursorType('default');
			}
			
			oDoc.bOffMarkerAfterUsing = true;
		}
	};
	PDFEditorApi.prototype.IsCommentMarker = function() {
        return this.curMarkerType !== undefined;
    };

	PDFEditorApi.prototype.get_PageWidth  = function(nPage) {
		let oDoc = this.getPDFDoc();
		return oDoc.GetPageWidthEMU();
	};
	PDFEditorApi.prototype.get_PageHeight = function(nPage) {
		let oDoc = this.getPDFDoc();
		return oDoc.GetPageHeightEMU();
	};
	PDFEditorApi.prototype.asc_SetFastCollaborative = function(isOn)
	{
		if (!AscCommon.CollaborativeEditing)
			return;
		
		AscCommon.CollaborativeEditing.Set_Fast(isOn);
	};
	/////////////////////////////////////////////////////////////
	///////// For annots
	////////////////////////////////////////////////////////////
	PDFEditorApi.prototype.AddFreeTextAnnot = function(nType) {
		let oDoc = this.getPDFDoc();

		oDoc.BlurActiveObject();
		function addFreeText() {
			oDoc.DoAction(function() {
				oDoc.AddFreeTextAnnot(nType, oDoc.Viewer.currentPage);
			}, AscDFH.historydescription_Pdf_AddAnnot, this);
		}

		if (oDoc.checkFonts(["Arial"], addFreeText)) {
			addFreeText();
		}
	};

	PDFEditorApi.prototype.AddStampAnnot = function(sType) {
		let oDoc = this.getPDFDoc();

		oDoc.BlurActiveObject();

		if (sType == AscPDF.STAMP_TYPES.Image) {
			let t = this;
			AscCommon.ShowImageFileDialog(this.documentId, this.documentUserId, this.CoAuthoringApi.get_jwt(), this.documentShardKey, this.documentWopiSrc, this.documentUserSessionId, function(error, files) {
				// ошибка может быть объектом в случае отмены добавления картинки в форму
				if (typeof(error) == "object")
					return;
		
				t._uploadCallback(error, files, {
					isStamp: true
				});
			},
			function(error) {
				if (Asc.c_oAscError.ID.No !== error) {
					t.sendEvent("asc_onError", error, Asc.c_oAscError.Level.NoCritical);
				}
		
				t.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.UploadImage);
			});
		}
		else {
			function addStamp() {
				oDoc.DoAction(function() {
					oDoc.AddStampAnnot(sType, oDoc.Viewer.currentPage);
				}, AscDFH.historydescription_Pdf_AddAnnot, this);
			}
	
			if (oDoc.checkFonts(["Arial"], addStamp)) {
				addStamp();
			}
		}
		
	};

	/////////////////////////////////////////////////////////////
	///////// For drawings
	////////////////////////////////////////////////////////////
	PDFEditorApi.prototype.StartAddShape = function(sPreset, is_apply) {
		let oDoc	= this.getPDFDoc();
		let oDrDoc	= oDoc.GetDrawingDocument();

		if (this.isDrawTablePen) {
			this.sync_TableDrawModeCallback(false);
        }
        if (this.isDrawTableErase) {
            this.sync_TableEraseModeCallback(false);
        }

		this.stopInkDrawer();
		this.cancelEyedropper();

		oDoc.BlurActiveObject();
		this.isStartAddShape = true;
		this.addShapePreset  = sPreset;
		if (is_apply) {
			oDrDoc.LockCursorType("crosshair");
		}
		else {
			editor.sync_EndAddShape();
			editor.sync_StartAddShapeCallback(false);
		}
	};
	PDFEditorApi.prototype.ShapeApply = function(prop) {
		let oDoc	= this.getPDFDoc();
		let oDrDoc	= oDoc.GetDrawingDocument();

		// нужно определить, картинка это или нет
		let image_url = "";
		let sToken = undefined;
		prop.Width    = prop.w;
		prop.Height   = prop.h;

		let bShapeTexture = true;
		if (prop.fill != null) {
			if (prop.fill.fill != null && prop.fill.type == Asc.c_oAscFill.FILL_TYPE_BLIP) {
				image_url	= prop.fill.fill.asc_getUrl();
				sToken		= prop.fill.fill.token;

				let _tx_id = prop.fill.fill.asc_getTextureId();
				if (null != _tx_id && 0 <= _tx_id && _tx_id < AscCommon.g_oUserTexturePresets.length) {
					image_url = AscCommon.g_oUserTexturePresets[_tx_id];
				}
			}
		}

		let oFill;
		if (prop.textArtProperties) {
			oFill = prop.textArtProperties.asc_getFill();

			if (oFill && oFill.fill != null && oFill.type == Asc.c_oAscFill.FILL_TYPE_BLIP) {
				image_url	= oFill.fill.asc_getUrl();
				sToken		= oFill.fill.token;

				let _tx_id = oFill.fill.asc_getTextureId();
				if (null != _tx_id && 0 <= _tx_id && _tx_id < AscCommon.g_oUserTexturePresets.length) {
					image_url = AscCommon.g_oUserTexturePresets[_tx_id];
				}

				bShapeTexture = false;
			}
		}

		if (!AscCommon.isNullOrEmptyString(image_url)) {
			let sImageUrl = null;
			if (!AscCommon.g_oDocumentUrls.getImageLocal(image_url)) {
				sImageUrl = image_url;
			}

			let oApi           = this;
			let fApplyCallback = function() {
				let _image   = oApi.ImageLoader.LoadImage(image_url, 1);
				let srcLocal = AscCommon.g_oDocumentUrls.getImageLocal(image_url);

				if (srcLocal) {
					image_url = srcLocal;
				}

				if (bShapeTexture) {
					prop.fill.fill.asc_putUrl(image_url); // erase documentUrl
				}
				else {
					oFill.fill.asc_putUrl(image_url);
				}

				if (null != _image || window["NATIVE_EDITOR_ENJINE"]) {
					oDoc.ShapeApply(prop);

					if (bShapeTexture) {
						oDrDoc.DrawImageTextureFillShape(image_url);
					}
					else {
						oDrDoc.DrawImageTextureFillTextArt(image_url);
					}
				}

				else {
					oApi.sync_StartAction(Asc.c_oAscAsyncActionType.Information, Asc.c_oAscAsyncAction.LoadImage);
					let oProp = prop;

					oApi.asyncImageEndLoaded2 = function(_image) {
						oDoc.ShapeApply(oProp);
						oDrDoc.DrawImageTextureFillShape(image_url);
						oApi.sync_EndAction(Asc.c_oAscAsyncActionType.Information, Asc.c_oAscAsyncAction.LoadImage);
						oApi.asyncImageEndLoaded2 = null;
					}
				}
			};
			
			if (!sImageUrl) {
				fApplyCallback();
			}
			else {

				if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]()) {
					image_url = window["AscDesktopEditor"]["LocalFileGetImageUrl"](sImageUrl);
					image_url = AscCommon.g_oDocumentUrls.getImageUrl(image_url);

					fApplyCallback();
					return;
				}

                AscCommon.sendImgUrls(this, [sImageUrl], function(data) {

                    if (data && data[0] && data[0].url !== "error") {
                        image_url = data[0].url;
                        fApplyCallback();
                    }

                }, undefined, sToken);
			}
		}
		else {
			oDoc.ShapeApply(prop);
		}
	};
	PDFEditorApi.prototype.ChangeShapeType = function(sShapetype) {
		let oDoc = this.getPDFDoc();

		oDoc.ChangeShapeType(sShapetype);
	};
	PDFEditorApi.prototype.ImgApply = function(obj) {
		let oDoc	= this.getPDFDoc();
		let ImagePr	= {};

		ImagePr.lockAspect = obj.lockAspect;
		ImagePr.Width      = null === obj.Width || undefined === obj.Width ? null : parseFloat(obj.Width);
		ImagePr.Height     = null === obj.Height || undefined === obj.Height ? null : parseFloat(obj.Height);

		ImagePr.title				= obj.title;
		ImagePr.bSetOriginalSize	= obj.bSetOriginalSize;
		ImagePr.description			= obj.description;
		ImagePr.name				= obj.name;
		ImagePr.rot					= obj.rot;
		ImagePr.rotAdd				= obj.rotAdd;
		ImagePr.flipH				= obj.flipH;
		ImagePr.flipV				= obj.flipV;
		ImagePr.flipHInvert			= obj.flipHInvert;
		ImagePr.flipVInvert			= obj.flipVInvert;
		ImagePr.resetCrop			= obj.resetCrop;
		ImagePr.transparent         = obj.transparent;

		if (undefined != obj.Position) {
			ImagePr.Position =
			{
				X : null === obj.Position.X || undefined === obj.Position.X ? null : parseFloat(obj.Position.X),
				Y : null === obj.Position.Y || undefined === obj.Position.Y ? null : parseFloat(obj.Position.Y)
			};
		}
		else {
			ImagePr.Position = {X : null, Y : null};
		}

		ImagePr.ImageUrl = obj.ImageUrl;

		if (window["NATIVE_EDITOR_ENJINE"]) {
		  	oDoc.SetImageProps(ImagePr);
		  	return;
		}
		if (!AscCommon.isNullOrEmptyString(ImagePr.ImageUrl)) {
			let sImageUrl = null, sToken = undefined;
			if (!AscCommon.g_oDocumentUrls.getImageLocal(ImagePr.ImageUrl)) {
				sImageUrl = ImagePr.ImageUrl;
				sToken = obj.Token;
			}

			let oApi           = this;
			let fApplyCallback = function() {
				let _img     = oApi.ImageLoader.LoadImage(ImagePr.ImageUrl, 1);
				let srcLocal = AscCommon.g_oDocumentUrls.getImageLocal(ImagePr.ImageUrl);

				if (srcLocal) {
					ImagePr.ImageUrl = srcLocal;
				}
				if (null != _img) {
					oDoc.SetImageProps(ImagePr);
				}
				else {
					oApi.asyncImageEndLoaded2 = function(_image) {
						oDoc.SetImageProps(ImagePr);
						oApi.asyncImageEndLoaded2 = null;
					}
				}
			};

			if (!sImageUrl) {
				fApplyCallback();
			}
			else {
				if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]()) {
                    this.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.UploadImage);
                    let _url = window["AscDesktopEditor"]["LocalFileGetImageUrl"](sImageUrl);
                    _url     = AscCommon.g_oDocumentUrls.getImageUrl(_url);
                    ImagePr.ImageUrl = _url;
                    fApplyCallback();
                    this.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.UploadImage);
                    return;
                }

                AscCommon.sendImgUrls(this, [sImageUrl], function(data) {

                    if (data && data[0] && data[0].url !== "error") {
                        ImagePr.ImageUrl = data[0].url;
                        fApplyCallback();
                    }

                }, undefined, sToken);
			}
		}
		else {
			ImagePr.ImageUrl = null;
			oDoc.SetImageProps(ImagePr);
		}
	};
	PDFEditorApi.prototype.asc_FitImagesToPage = function() {
		let oDoc = this.getPDFDoc();
		oDoc.FitImagesToPage();
	};

	PDFEditorApi.prototype.asc_SetFillColor = function(r, g, b) {
		let oDoc = this.getPDFDoc();
		let oController = oDoc.GetController();
		let oMouseDownAnnot = oDoc.mouseDownAnnot;

		if (!oMouseDownAnnot) {
			return false;
		}

		let aColor = [r / 255, g / 255, b / 255];
		
		return oDoc.DoAction(function() {
			oController.selectedObjects.forEach(function(annot) {
				if (annot.IsTextMarkup()) {
					annot.SetFillColor(aColor);
				}
			});

			return true;
        }, AscDFH.historydescription_Pdf_ChangeFillColor);
	};

	PDFEditorApi.prototype.asc_GetFillColor = function() {
		let oDoc = this.getPDFDoc();
		let oMouseDownAnnot = oDoc.mouseDownAnnot;

		if (!oMouseDownAnnot) {
			return null;
		}

		let oColor = oMouseDownAnnot.GetRGBColor(oMouseDownAnnot.GetFillColor());
		oColor["r"] = oColor.r;
        oColor["g"] = oColor.g;
        oColor["b"] = oColor.b;

		return oColor;
	};

	PDFEditorApi.prototype.asc_SetStrokeColor = function(r, g, b) {
		let oDoc = this.getPDFDoc();
		let oController = oDoc.GetController();
		let oMouseDownAnnot = oDoc.mouseDownAnnot;

		if (!oMouseDownAnnot) {
			return false;
		}

		let aColor = [r / 255, g / 255, b / 255];

		return oDoc.DoAction(function() {
			oController.selectedObjects.forEach(function(annot) {
				if (annot.IsTextMarkup()) {
					annot.SetStrokeColor(aColor);
				}
			});

			return true;
        }, AscDFH.historydescription_Pdf_ChangeStrokeColor);
	};

	PDFEditorApi.prototype.asc_GetStrokeColor = function() {
		let oDoc = this.getPDFDoc();
		let oMouseDownAnnot = oDoc.mouseDownAnnot;

		if (!oMouseDownAnnot) {
			return null;
		}

		let oColor = oMouseDownAnnot.GetRGBColor(oMouseDownAnnot.GetStrokeColor());
		oColor["r"] = oColor.r;
        oColor["g"] = oColor.g;
        oColor["b"] = oColor.b;

		return oColor;
	};

	PDFEditorApi.prototype.asc_SetOpacity = function(nValue) {
		let oDoc = this.getPDFDoc();
		let oController = oDoc.GetController();
		let oMouseDownAnnot = oDoc.mouseDownAnnot;

		if (!oMouseDownAnnot) {
			return false;
		}

		return oDoc.DoAction(function() {
			oController.selectedObjects.forEach(function(annot) {
				if (annot.IsTextMarkup()) {
					annot.SetOpacity(nValue / 100);
				}
			});

			return true;
        }, AscDFH.historydescription_Pdf_ChangeOpacity);
	};

	PDFEditorApi.prototype.asc_CloseFile = function() {
		AscCommon.History.Clear();
		AscCommon.g_oIdCounter.Clear();
		AscCommon.g_oTableId.Clear();
		AscCommon.CollaborativeEditing.Clear();
		this.isApplyChangesOnOpenEnabled = true;
		this.isDocumentLoadComplete = false;
		this.ServerImagesWaitComplete = false;
		this.turnOffSpecialModes();
		AscCommon.pptx_content_loader.ImageMapChecker = {};

		this.sendEvent("asc_onCloseFile");
	};

	PDFEditorApi.prototype.asc_GetOpacity = function() {
		let oDoc = this.getPDFDoc();
		let oMouseDownAnnot = oDoc.mouseDownAnnot;

		if (!oMouseDownAnnot) {
			return null;
		}

		return oMouseDownAnnot.GetOpacity() * 100;
	};

	/////////////////////////////////////////////////////////////
	///////// For table
	////////////////////////////////////////////////////////////
	PDFEditorApi.prototype.put_Table = function(col, row, placeholder, sStyleId) {
		let oDoc = this.getPDFDoc();
		oDoc.DoAction(function() {
			oDoc.AddTable(col, row, sStyleId, oDoc.GetCurPage());
		}, AscDFH.historydescription_Presentation_AddFlowTable, this);
	};
	PDFEditorApi.prototype.tblApply = function(oPr) {
		let oDoc = this.getPDFDoc();

		let oBorders = oPr.CellBorders;
		if (oPr.CellBorders){
			function fCheckBorder(oBorder) {
				if(!oBorder || !oBorder.Color)
					return;
				oBorder.Unifill =  AscFormat.CreateUnifillFromAscColor(oBorder.Color, 0);
			}
			fCheckBorder(oBorders.Left);
			fCheckBorder(oBorders.Top);
			fCheckBorder(oBorders.Right);
			fCheckBorder(oBorders.Bottom);
			fCheckBorder(oBorders.InsideH);
			fCheckBorder(oBorders.InsideV);
		}
		let oBackground = oPr.CellsBackground;
		if (oBackground && oBackground.Color) {
			if (oBackground.Value === Asc.c_oAscShd.Nil){
				oBackground.Value = Asc.c_oAscShd.Clear;
				oBackground.Unifill = AscFormat.CreateNoFillUniFill();
			}
			else {
				oBackground.Unifill = AscFormat.CreateUnifillFromAscColor(oBackground.Color, 0);
			}
		}
		oDoc.SetTableProps(oPr);
	};
	PDFEditorApi.prototype.asc_DistributeTableCells = function(isHorizontally) {
		let oDoc	= this.getPDFDoc();
		let bResult	= false;

		bResult = oDoc.DistributeTableCells(isHorizontally);
		return bResult;
	};
	PDFEditorApi.prototype.remColumn = function() {
		let oDoc = this.getPDFDoc();
		oDoc.RemoveTableColumn();
		return true;
	};
	PDFEditorApi.prototype.remTable = function() {
		let oDoc = this.getPDFDoc();
		let oObject = oDoc.GetActiveObject();
		
		if (oObject && oObject.IsDrawing() && oObject.IsGraphicFrame()) {
			oDoc.CreateNewHistoryPoint();
			oDoc.RemoveDrawing(oObject.GetId());
			oDoc.TurnOffHistory();
			return true;
		}

		return false;
	};
	PDFEditorApi.prototype.asc_getTableStylesPreviews = function(bUseDefault, arrIds) {
		let oDoc = this.getPDFDoc();

		this.private_CheckTableStylesPreviewGenerator();
		let arrPreviews = [];

		if (arrIds && arrIds.length) {
			let oStylesList = oDoc.GetTableStyles();
			for (let nIndex = 0, nCount = arrIds.length; nIndex < nCount; ++nIndex) {
				let oStyle   = oStylesList.Get(arrIds[nIndex]);
				let oPreview = this.TableStylesPreviewGenerator.GetPreview(oStyle);
				if (oPreview)
					arrPreviews.push(oPreview);
			}
		}
		else {
			arrPreviews = this.TableStylesPreviewGenerator.GetAllPreviews(bUseDefault);
		}

		return arrPreviews;
	};
	PDFEditorApi.prototype.asc_GetSelectionBounds = function() {
		return [[0,0], [0,0], [0,0], [0,0]];
	};

	PDFEditorApi.prototype.getPluginContextMenuInfo = function () {
		let oDoc		= this.getPDFDoc();
		let oController	= oDoc.GetController();
		
		return oController.getPluginSelectionInfo();
	};

	PDFEditorApi.prototype.SetShowTextSelectPanel = function(bShow) {
		this.showTextSelectPanel = bShow;
		
		this.getPDFDoc().UpdateSelectionTrackPos();
	};
	PDFEditorApi.prototype.NeedShowTextSelectPanel = function() {
		return this.showTextSelectPanel;
	};

	/////////////////////////////////////////////////////////////
	///////// For text
	////////////////////////////////////////////////////////////

	PDFEditorApi.prototype.put_TextPrBold = function(value) {
		this.getPDFDoc().AddToParagraph(new AscCommonWord.ParaTextPr({Bold : value}));
	};
	PDFEditorApi.prototype.put_TextPrItalic = function(value) {
		this.getPDFDoc().AddToParagraph(new AscCommonWord.ParaTextPr({Italic : value}));
	};
	PDFEditorApi.prototype.put_TextPrUnderline = function(value) {
		this.getPDFDoc().AddToParagraph(new AscCommonWord.ParaTextPr({Underline : value}));
	};
	PDFEditorApi.prototype.put_TextPrStrikeout = function(value) {
		this.getPDFDoc().AddToParagraph(new AscCommonWord.ParaTextPr({
			Strikeout  : value,
			DStrikeout : false
		}));
	};
	PDFEditorApi.prototype.put_PrLineSpacing = function(nType, nValue) {
		this.getPDFDoc().SetParagraphSpacing({LineRule : nType, Line : nValue});
	};
	PDFEditorApi.prototype.put_LineSpacingBeforeAfter = function(type, value) { //"type == 0" means "Before", "type == 1" means "After"
		switch (type) {
			case 0:
				this.getPDFDoc().SetParagraphSpacing({Before : value});
				break;
			case 1:
				this.getPDFDoc().SetParagraphSpacing({After : value});
				break;
		}
	};
	PDFEditorApi.prototype.FontSizeIn = function() {
		this.getPDFDoc().IncreaseDecreaseFontSize(true);
	};
	PDFEditorApi.prototype.FontSizeOut = function() {
		this.getPDFDoc().IncreaseDecreaseFontSize(false);
	};
	// 0- baseline, 2-subscript, 1-superscript
	PDFEditorApi.prototype.put_TextPrBaseline = function(value) {
		this.getPDFDoc().AddToParagraph(new AscCommonWord.ParaTextPr({VertAlign : value}));
	};
	PDFEditorApi.prototype.put_TextPrFontSize = function(size) {
		this.getPDFDoc().AddToParagraph(new AscCommonWord.ParaTextPr({FontSize : Math.min(size, 300)}));
	};
	PDFEditorApi.prototype.put_TextPrFontName = function(name) {
		var loader   = AscCommon.g_font_loader;
		var fontinfo = AscFonts.g_fontApplication.GetFontInfo(name);
		var isasync  = loader.LoadFont(fontinfo);

		if (false === isasync) {
			this.getPDFDoc().AddToParagraph(new AscCommonWord.ParaTextPr({
				FontFamily : {
					Name  : name,
					Index : -1
				}
			}));
		}
	};
	PDFEditorApi.prototype.put_TextColor = function(color) {
		var _unifill        = new AscFormat.CUniFill();
		_unifill.fill       = new AscFormat.CSolidFill();
		_unifill.fill.color = AscFormat.CorrectUniColor(color, _unifill.fill.color, 0);

		this.getPDFDoc().AddToParagraph(new AscCommonWord.ParaTextPr({Unifill : _unifill}), false);
	};
	PDFEditorApi.prototype.asc_ChangeTextCase = function(nType) {
		this.getPDFDoc().ChangeTextCase(nType);
	};
	PDFEditorApi.prototype.put_PrAlign = function(nType) {
		this.getPDFDoc().SetParagraphAlign(nType);
	};
	PDFEditorApi.prototype.setVerticalAlign = function(nType) {
		this.getPDFDoc().SetVerticalAlign(nType);
	};
	PDFEditorApi.prototype.IncreaseIndent = function() {
		this.getPDFDoc().IncreaseDecreaseIndent(true);
	};
	PDFEditorApi.prototype.DecreaseIndent = function(){
		this.getPDFDoc().IncreaseDecreaseIndent(false);
	};
	PDFEditorApi.prototype.ClearFormating = function() {
		this.getPDFDoc().ClearParagraphFormatting(false, true);
	};
	PDFEditorApi.prototype.UpdateParagraphProp = function(oParaPr) {
		let oDoc			= this.getPDFDoc();
		let TextPr			= oDoc.GetCalculatedTextPr(true);
		let oDrawingProps	= oDoc.Get_GraphicObjectsProps();

		if (oDrawingProps.shapeProps && oDrawingProps.shapeProps.locked
		|| oDrawingProps.chartProps && oDrawingProps.chartProps.locked
		|| oDrawingProps.tableProps && oDrawingProps.tableProps.Locked) {
			oParaPr.Locked = true;
		}
		
		oParaPr.Subscript   = ( TextPr.VertAlign === AscCommon.vertalign_SubScript ? true : false );
		oParaPr.Superscript = ( TextPr.VertAlign === AscCommon.vertalign_SuperScript ? true : false );
		oParaPr.Strikeout   = TextPr.Strikeout;
		oParaPr.DStrikeout  = TextPr.DStrikeout;
		oParaPr.AllCaps     = TextPr.Caps;
		oParaPr.SmallCaps   = TextPr.SmallCaps;
		oParaPr.TextSpacing = TextPr.Spacing;
		oParaPr.Position    = TextPr.Position;
		oParaPr.ListType	= AscFormat.fGetListTypeFromBullet(oParaPr.Bullet);

		this.sync_ParaSpacingLine(oParaPr.Spacing);
		this.Update_ParaInd(oParaPr.Ind);
		this.sync_PrAlignCallBack(oParaPr.Jc);
		this.sync_ParaStyleName(oParaPr.StyleName);
		this.sync_ListType(oParaPr.ListType);
		this.sync_PrPropCallback(oParaPr);
	};
	PDFEditorApi.prototype.paraApply = function(Props) {
		let oDoc		= this.getPDFDoc();
		let oController	= oDoc.GetController();
        
		let sLoadFont = null, sLoadText = null;
		let fCallback = function() {
			oController.paraApplyCallback(Props);
		};
		
		let oBullet = Props.asc_getBullet();
		if (oBullet) {
			sLoadFont = oBullet.asc_getFont();
			sLoadText = oBullet.asc_getSymbol();
		}

		if (typeof sLoadFont === "string" && sLoadFont.length > 0
		&& typeof sLoadText === "string" && sLoadText.length > 0) {
			let loader   = AscCommon.g_font_loader;
			let fontinfo = AscFonts.g_fontApplication.GetFontInfo(sLoadFont);
			let isasync  = loader.LoadFont(fontinfo);

			if (false === isasync) {
				AscFonts.FontPickerByCharacter.checkText(sLoadText, this, function () {
					oController.checkSelectedObjectsAndCallback(fCallback, [], false, AscDFH.historydescription_Presentation_ParaApply);
				});
			}
			else {
				this.asyncMethodCallback = function() {
					AscFonts.FontPickerByCharacter.checkText(sLoadText, this, function () {
						oController.checkSelectedObjectsAndCallback(fCallback, [], false, AscDFH.historydescription_Presentation_ParaApply);
					});
				}
			}
		}
		else {
			oController.checkSelectedObjectsAndCallback(fCallback, [], false, AscDFH.historydescription_Presentation_ParaApply);
		}
	};


	/////////////////////////////////////////////////////////////
	///////// For text
	////////////////////////////////////////////////////////////

	PDFEditorApi.prototype.sync_ListType = function(NumPr) {
		this.sendEvent("asc_onListType", new AscCommon.asc_CListType(NumPr));
	};
	PDFEditorApi.prototype.ParseBulletPreviewInformation = function(arrDrawingInfo) {
		const arrNumberingLvls = [];
		AscFormat.ExecuteNoHistory(function ()
		{
			for (let i = 0; i < arrDrawingInfo.length; i += 1)
			{
				const oDrawInfo = arrDrawingInfo[i];
				const oNumberingInfo = oDrawInfo["numberingInfo"];
				if (!oNumberingInfo) continue;
				const sDivId = oDrawInfo["divId"];
				if (!oNumberingInfo["bullet"])
				{
					const oPresentationBullet = new AscCommonWord.CPresentationBullet();
					const oTextPr = new AscCommonWord.CTextPr();
					oPresentationBullet.m_sChar = AscCommon.translateManager.getValue("None");
					oPresentationBullet.m_nType = AscFormat.numbering_presentationnumfrmt_Char;
					oPresentationBullet.m_bFontTx = false;
					oPresentationBullet.m_sFont   = "Arial";
					oTextPr.Unifill = AscFormat.CreateSolidFillRGB(0, 0, 0);
					oTextPr.FontSize = oTextPr.FontSizeCS = 65;
					oPresentationBullet.MergeTextPr(oTextPr);
					arrNumberingLvls.push({divId: sDivId, arrLvls: [oPresentationBullet], isRemoving: true});
				}
				else
				{
					const oBullet = window['AscJsonConverter'].ReaderFromJSON.prototype.BulletFromJSON(oNumberingInfo["bullet"]);
					const oPresentationBullet = oBullet.getPresentationBullet(AscFormat.GetDefaultTheme(), AscFormat.GetDefaultColorMap());
					oPresentationBullet.m_bFontTx = false;
					const oTextPr = new AscCommonWord.CTextPr();
					oTextPr.Unifill = AscFormat.CreateSolidFillRGB(0, 0, 0);
					oTextPr.FontSize = oTextPr.FontSizeCS = 65;
					oPresentationBullet.MergeTextPr(oTextPr);
					arrNumberingLvls.push({divId: sDivId, arrLvls: [oPresentationBullet]});
				}
			}
		}, this);
		return arrNumberingLvls;
	};
	PDFEditorApi.prototype.put_ListType = function(type, subtype, custom) {
		let blipUrl = custom && custom.imageId;
		if (blipUrl) {
			let checkImageUrlFromServer;

			let that		= this;
			let localUrl	= AscCommon.g_oDocumentUrls.getLocal(blipUrl);
			let fullUrl		= AscCommon.g_oDocumentUrls.getUrl(blipUrl);

			if (fullUrl) {
				checkImageUrlFromServer = fullUrl;
			}
			else if (localUrl) {
				checkImageUrlFromServer = blipUrl;
			}

			if (checkImageUrlFromServer) {
				blipUrl			= checkImageUrlFromServer;
				custom.imageId	= blipUrl;

				let isImageNotAttendInImageLoader = !this.ImageLoader.map_image_index[blipUrl];
				if (isImageNotAttendInImageLoader) {
					let tryToSetImageBulletAgain = function () {
						that.put_ListType(type, subtype, custom);
					}
					this.ImageLoader.LoadImagesWithCallback([blipUrl], tryToSetImageBulletAgain);
					return;
				}
			}
			else {
				let changeBlipFillUrlToLocalAndTrySetImageBulletAgain = function (data) {
					let uploadImageUrl = data[0].url;
					custom.imageId = uploadImageUrl;
					that.put_ListType(type, subtype, custom);
				}
				AscCommon.sendImgUrls(this, [blipUrl], changeBlipFillUrlToLocalAndTrySetImageBulletAgain, false, custom.token);
				return;
			}
		}

		let oDoc = this.getPDFDoc();
		let NumberInfo = {
			Type:		type,
			SubType:	subtype,
			Custom:		custom
		};
		let oBullet = AscFormat.fGetPresentationBulletByNumInfo(NumberInfo);
		let sBullet = oBullet.asc_getSymbol();

		let fCallback = function() {
			oDoc.SetParagraphNumbering(oBullet);
		};

		if(typeof sBullet === "string" && sBullet.length > 0) {
			AscFonts.FontPickerByCharacter.checkText(sBullet, this, fCallback);
		}
		else {
			fCallback();
		}
	};
	PDFEditorApi.prototype.asc_GetPossibleNumberingLanguage = function(){};
	PDFEditorApi.prototype._addImageUrl = function(arrUrls, oOptionObject) {
		let oDoc = this.getPDFDoc();
		
		if (oOptionObject) {
			if (oOptionObject.sendUrlsToFrameEditor && this.isOpenedChartFrame) {
				this.addImageUrlsFromGeneralToFrameEditor(arrUrls);
				return;
			}
			else if (oOptionObject.isImageChangeUrl || oOptionObject.isTextArtChangeUrl || oOptionObject.isShapeImageChangeUrl || oOptionObject["obj"] || (oOptionObject instanceof AscCommon.CContentControlPr && oOptionObject.GetInternalId()) || oOptionObject.fAfterUploadOleObjectImage) {
				this.AddImageUrlAction(arrUrls[0], undefined, oOptionObject);
				return;
			}
		}

		if (this.ImageLoader) {
			const oApi = this;
			this.ImageLoader.LoadImagesWithCallback(arrUrls, function() {
				if (oOptionObject) {
					const oImage = oApi.ImageLoader.LoadImage(arrUrls[0], 1);
					if (!oImage || !oImage.Image) {
						return;
					}

					if (oOptionObject.GetType && oOptionObject.GetType() === AscPDF.FIELD_TYPES.button) {
						oOptionObject.AddImage(oImage);
					}
					else if (oOptionObject.isStamp) {
						oDoc.AddStampAnnot(AscPDF.STAMP_TYPES.Image, oDoc.Viewer.currentPage, oImage);
					}
				}
				else {
					const arrImages = [];
					for (let i = 0; i < arrUrls.length; ++i) {
						const oImage = oApi.ImageLoader.LoadImage(arrUrls[i], 1);
						if (oImage.Image) {
							arrImages.push(oImage);
						}
					}
					if (arrImages.length) {
						oDoc.DoAction(function() {
							oDoc.AddImages(arrImages);
						}, AscDFH.historydescription_Presentation_AddFlowImage)
					}
				}
			}, []);
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
	PDFEditorApi.prototype.asc_SetTextFormDatePickerDate = function(oPr)
	{
		let oDoc = this.getPDFDoc();
		let oActiveForm = oDoc.activeForm;
		if (!oActiveForm)
			return;

		let oDate = new Asc.cDate(oPr.GetFullDate());
		let oCurDate = new Date();

		oDate.setMinutes(oCurDate.getMinutes());
		oDate.setSeconds(oCurDate.getSeconds());
		oDate.getMilliseconds(oCurDate.getMilliseconds());

		oDoc.lastDatePickerInfo = {
			value: AscPDF.FormatDateValue(oActiveForm.GetDateFormat(), oDate.getTime()),
			form: oActiveForm
		};

		oActiveForm.content.SelectAll();
		oActiveForm.EnterText(AscWord.CTextFormFormat.prototype.GetBuffer(oDoc.lastDatePickerInfo.value));
		oDoc.EnterDownActiveField();
		oDoc.lastDatePickerInfo = null;
	};
	PDFEditorApi.prototype.asc_SelectPDFFormListItem = function(sId) {
		let oViewer	= this.DocumentRenderer;
		let oDoc	= oViewer.getPDFDoc();
		let oField	= oDoc.activeForm;
		let nIdx	= parseInt(sId);
		if (!oField)
			return;
				
		oDoc.DoAction(function() {
			oField.SelectOption(nIdx);
			if (oField.IsCommitOnSelChange() && oField.IsNeedCommit()) {
				oDoc.EnterDownActiveField();
			}	
		}, AscDFH.historydescription_Pdf_FieldSelectOption, this);
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
	PDFEditorApi.prototype.SetDocumentModified = function(bValue) {
		this.isDocumentModify = bValue;
		this.sendEvent("asc_onDocumentModifiedChanged");

		if (undefined !== window["AscDesktopEditor"]) {
			window["AscDesktopEditor"]["onDocumentModifiedChanged"](bValue);
		}
	};
	PDFEditorApi.prototype.CheckChangedDocument = function() {
		if (true === AscCommon.History.Have_Changes()) {
			this.SetDocumentModified(true);
		}
		else {
			this.SetDocumentModified(false);
		}

		this._onUpdateDocumentCanSave();
	};
	PDFEditorApi.prototype._autoSaveInner = function() {
		let _curTime = new Date();
		let oDoc = this.getPDFDoc();
		let oHistory = oDoc.History;

		if (null === this.lastSaveTime) {
			this.lastSaveTime = _curTime;
		}
		
		if (AscCommon.CollaborativeEditing.Is_Fast() && !AscCommon.CollaborativeEditing.Is_SingleUser()) {
			this.WordControl.m_oLogicDocument.Continue_FastCollaborativeEditing();
		}
		else if (this.isLiveViewer()) {
			if (AscCommon.CollaborativeEditing.Have_OtherChanges()) {
				AscCommon.CollaborativeEditing.Apply_Changes();
			}
		}
		else {
			let _bIsWaitScheme = false;
			if (oHistory.Points && oHistory.Index >= 0 && oHistory.Index < oHistory.Points.length) {
				if ((_curTime - oHistory.Points[oHistory.Index].Time) < this.intervalWaitAutoSave) {
					_bIsWaitScheme = true;
				}
			}

			if (!_bIsWaitScheme) {
				let _interval = (AscCommon.CollaborativeEditing.m_nUseType <= 0) ? this.autoSaveGapSlow :
					this.autoSaveGapFast;

				if ((_curTime - this.lastSaveTime) > _interval) {
					if (oHistory.Have_Changes(true) == true && oDoc.LocalHistory !== AscCommon.History) {
						this.asc_Save(true);
					}
					this.lastSaveTime = _curTime;
				}
			}
		}
	};
	PDFEditorApi.prototype.asc_Save = function (isAutoSave, isIdle) {
		let oDoc = this.getPDFDoc();
		if (!isAutoSave)
			oDoc.BlurActiveObject();
		AscCommon.DocumentEditorApi.prototype.asc_Save.call(this, isAutoSave, isIdle);
	};
	PDFEditorApi.prototype._onEndLoadSdk = function() {
		AscCommon.DocumentEditorApi.prototype._onEndLoadSdk.call(this);

		this.stampAnnotPreviewManager = new AscPDF.StampAnnotPreviewManager();
	};
	PDFEditorApi.prototype.asc_getPropertyEditorStamps = function() {
		return this.stampAnnotPreviewManager.getStampPreviews();
	};
	PDFEditorApi.prototype.loadStampsJSON = function() {
		try {
			if (window["native_pdf_stamps"]) {
				AscPDF["STAMPS_JSON"] = AscPDF.STAMPS_JSON = window["native_pdf_stamps"];
				delete window["native_pdf_stamps"];
				return;
			}
			var xhr = new XMLHttpRequest();
			xhr.open("GET", "../../../../sdkjs/pdf/src/annotations/stamps.json", true);
			var t = this;
			xhr.onload = function()
			{
				if (this.status === 200 || location.href.indexOf("file:") === 0)
				{
					try
					{
						AscPDF["STAMPS_JSON"] = AscPDF.STAMPS_JSON = JSON.parse(this.responseText);
					}
					catch (err) {}
				}
			};
			xhr.send('');
		}
		catch (e) {}
	};
	PDFEditorApi.prototype._init = function() {
		AscCommon.DocumentEditorApi.prototype._init.call(this);
		this.loadStampsJSON();
	}
	PDFEditorApi.prototype._coAuthoringInitEnd = function() {
		AscCommon.DocumentEditorApi.prototype._coAuthoringInitEnd.call(this);

		let t = this;
		this.CoAuthoringApi.onConnectionStateChanged = function(e) {
			if (true === AscCommon.CollaborativeEditing.Is_Fast() && false === e['state']) {
				t.WordControl.m_oLogicDocument.Remove_ForeignCursor(e['id']);
				t.DocumentRenderer.onUpdateOverlay();
			}
			t.sendEvent("asc_onConnectionStateChanged", e);
		};
		this.CoAuthoringApi.onLocksAcquired = function(e)
        {
            if (t._coAuthoringCheckEndOpenDocument(t.CoAuthoringApi.onLocksAcquired, e)) {
                return;
            }
            let oDoc = t.getPDFDoc();
            let oThumbnails = oDoc.GetThumbnails();
            if (2 != e["state"]) {
                let Id    = e["block"];
                let Class = AscCommon.g_oTableId.Get_ById(Id);
                if (Class && e["blockValue"]["type"] == AscPDF.AscLockTypeElemPDF.Page) {
                    oThumbnails && oThumbnails._repaintPage(Class.GetIndex());
                }
                if (null != Class) {
					function updateLock(Class) {
						let Lock = Class.Lock;
						if (!Lock) {
							return;
						}

						// Выставляем ID пользователя, залочившего данный элемент
						Lock.Set_UserId(e["user"]);
						let OldType = Lock.Get_Type();
						if (AscCommon.c_oAscLockTypes.kLockTypeOther2 === OldType || AscCommon.c_oAscLockTypes.kLockTypeOther3 === OldType) {
							Lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeOther3, true);
						}
						else {
							Lock.Set_Type(AscCommon.c_oAscLockTypes.kLockTypeOther, true);
						}

						Class.AddToRedraw && Class.AddToRedraw();
					}

					if (Class.IsForm && Class.IsForm()) {
						let aWidgets = oDoc.GetAllWidgets(Class.GetFullName());
						aWidgets.forEach(function(widget) {
							updateLock(widget);
						});
					}
					else {
						updateLock(Class);
					}

                    if (Class.IsAnnot && Class.IsAnnot()) {
						// если аннотация коммент или аннотация с комментом то блокируем и комментарий тоже
						if (Class.IsComment() || (Class.IsUseContentAsComment() && Class.GetContents() != undefined) || Class.GetReply(0) != null) {
							t.sync_LockComment(Class.Get_Id(), e["user"]);
						}
					}
					
                    oDoc.UpdateInterface();
                }
                else {
                    AscCommon.CollaborativeEditing.Add_NeedLock(Id, e["user"]);
                }
            }
        };
		this.CoAuthoringApi.onLocksReleased = function(e, bChanges) {
			if (t._coAuthoringCheckEndOpenDocument(t.CoAuthoringApi.onLocksReleased, e, bChanges)) {
				return;
			}
		
			let oDoc = t.getPDFDoc();
            let oThumbnails = oDoc.GetThumbnails();
			let Id = e["block"]["guid"];
			let Class = AscCommon.g_oTableId.Get_ById(Id);
			if (Class && e["block"]["type"] == AscPDF.AscLockTypeElemPDF.Page) {
				oThumbnails && oThumbnails._repaintPage(Class.GetIndex());
			}
			if (null != Class) {
				function updateLock(Class) {
					let Lock = Class.Lock;
					if (!Lock) {
						return;
					}

					let CurType = Lock.Get_Type();
					let NewType = AscCommon.c_oAscLockTypes.kLockTypeNone;
		
					if (CurType === AscCommon.c_oAscLockTypes.kLockTypeOther) {
						if (true != bChanges) {
							NewType = AscCommon.c_oAscLockTypes.kLockTypeNone;
						} else {
							NewType = AscCommon.c_oAscLockTypes.kLockTypeOther2;
							AscCommon.CollaborativeEditing.Add_Unlock(Class);
						}
					} else if (CurType === AscCommon.c_oAscLockTypes.kLockTypeMine) {
						// Такого быть не должно
						NewType = AscCommon.c_oAscLockTypes.kLockTypeMine;
					} else if (CurType === AscCommon.c_oAscLockTypes.kLockTypeOther2 || CurType === AscCommon.c_oAscLockTypes.kLockTypeOther3) {
						NewType = AscCommon.c_oAscLockTypes.kLockTypeOther2;
					}
		
					Lock.Set_Type(NewType, true);

					if (NewType == AscCommon.c_oAscLockTypes.kLockTypeNone) {
						Class.AddToRedraw && Class.AddToRedraw();
						if (Class.IsAnnot && Class.IsAnnot()) {
							// if annot is comment or annot with comment then release locks for it too
							if (Class.IsComment() || (Class.IsUseContentAsComment() && Class.GetContents() != undefined) || Class.GetReply(0) != null) {
								Asc.editor.sync_UnLockComment(Class.Get_Id());
							}
						}
					}
				}
				

				if (Class.IsForm && Class.IsForm()) {
					let aWidgets = oDoc.GetAllWidgets(Class.GetFullName());
					aWidgets.forEach(function(widget) {
						updateLock(widget);
					});
				}
				else {
					updateLock(Class);
				}

				oDoc.UpdateInterface();
			} else {
				AscCommon.CollaborativeEditing.Remove_NeedLock(Id);
			}
		};
	};
	PDFEditorApi.prototype._autoSave = function () {

		if (this.canSave
			&& (!this.isViewMode || this.isLiveViewer())
			&& (this.canUnlockDocument || 0 !== this.autoSaveGap || AscCommon.CollaborativeEditing.Is_Fast()))
		{
			if (this.canUnlockDocument) {
				this.lastSaveTime = new Date();
				// Check edit mode after unlock document http://bugzilla.onlyoffice.com/show_bug.cgi?id=35971
				// Close cell edit without errors (isIdle = true)
				this.asc_Save(true, true);
			} else {
				this._autoSaveInner();
			}
		}
	};
	PDFEditorApi.prototype._haveChanges = function() {
		let oDoc = this.getPDFDoc();
		return oDoc.History.Have_Changes();
	};
	PDFEditorApi.prototype.pre_Save = function(_images) {
		this.isSaveFonts_Images = true;
		this.saveImageMap       = _images;
		this.FontLoader.LoadDocumentFonts2([]);
	};
	PDFEditorApi.prototype.asc_Print = function (options) {
		let oDoc = this.getPDFDoc();
		oDoc.BlurActiveObject();
		oDoc.RecalculateAll();
		AscCommon.DocumentEditorApi.prototype.asc_Print.call(this, options);
	};
	PDFEditorApi.prototype.asc_drawPrintPreview = function(index) {
		let oDoc = this.getPDFDoc();
		oDoc.BlurActiveObject();
		AscCommon.DocumentEditorApi.prototype.asc_drawPrintPreview.call(this, index);
	};
	PDFEditorApi.prototype.initCollaborativeEditing = function() {
		if (AscCommon.CollaborativeEditing)
			return;
		
		AscCommon.CollaborativeEditing = new AscPDF.CPDFCollaborativeEditing();
	};
	PDFEditorApi.prototype.ChangeReaderMode = function() {};
	PDFEditorApi.prototype.asc_getSelectedDrawingObjectsCount = function() {
		return this.WordControl.m_oLogicDocument.GetSelectedDrawingObjectsCount();
	};
	
	// composite input
	PDFEditorApi.prototype.Begin_CompositeInput = function() {
		if (this.compositeInput)
			return;
		
		this.compositeInput = AscPDF.CPDFCompositeInput.begin(this.getPDFDoc());
	};
	PDFEditorApi.prototype.Add_CompositeText = function(codePoint) {
		if (!this.compositeInput)
			return;
		
		this.compositeInput.add(codePoint);
	};
	PDFEditorApi.prototype.Remove_CompositeText = function(count) {
		if (!this.compositeInput)
			return;
		
		this.compositeInput.remove(count);
	};
	PDFEditorApi.prototype.Replace_CompositeText = function(codePoints) {
		if (!this.compositeInput)
			return;
		
		this.compositeInput.replace(codePoints);
	};
	PDFEditorApi.prototype.Set_CursorPosInCompositeText = function(pos) {
		if (!this.compositeInput)
			return;
		
		this.compositeInput.setPos(pos);
	};
	PDFEditorApi.prototype.Get_CursorPosInCompositeText = function() {
		if (!this.compositeInput)
			return 0;
		
		return this.compositeInput.getPos();
	};
	PDFEditorApi.prototype.Get_MaxCursorPosInCompositeText = function() {
		if (!this.compositeInput)
			return 0;
		
		return this.compositeInput.getMaxPos();
	};
	PDFEditorApi.prototype.End_CompositeInput = function() {
		if (!this.compositeInput)
			return;
		
		this.compositeInput.end();
		this.compositeInput = null;
	};
	PDFEditorApi.prototype.getAddedTextOnKeyDown = function() {
		return [];
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
		if (!oDoc)
			return null;

		let oCommentData = new AscCommon.CCommentData();
		oCommentData.Read_FromAscCommentData(AscCommentData);

		let oComment = oDoc.AddComment(AscCommentData);
		if (oComment) {
			return oComment.GetId()
		}
	};
	PDFEditorApi.prototype.asc_showComments = function()
	{
		let oDoc = this.getPDFDoc();
		oDoc.HideShowAnnots(false);
	};

	PDFEditorApi.prototype.asc_hideComments = function()
	{
		let oDoc = this.getPDFDoc();
		oDoc.HideShowAnnots(true);
	};
	PDFEditorApi.prototype.UpdateTextPr = function(TextPr)
	{
		let oDoc = this.getPDFDoc();
		let oDrDoc = oDoc.GetDrawingDocument();

		if ("undefined" != typeof(TextPr))
		{
			if (TextPr.Color !== undefined)
			{
				oDrDoc.TargetCursorColor.R = TextPr.Color.r;
				oDrDoc.TargetCursorColor.G = TextPr.Color.g;
				oDrDoc.TargetCursorColor.B = TextPr.Color.b;
			}
			if (TextPr.Bold === undefined)
				TextPr.Bold = false;
			if (TextPr.Italic === undefined)
				TextPr.Italic = false;
			if (TextPr.Underline === undefined)
				TextPr.Underline = false;
			if (TextPr.Strikeout === undefined)
				TextPr.Strikeout = false;
			if (TextPr.FontFamily === undefined)
				TextPr.FontFamily = {Index : 0, Name : ""};
			if (TextPr.FontSize === undefined)
				TextPr.FontSize = "";

			this.sync_BoldCallBack(TextPr.Bold);
			this.sync_ItalicCallBack(TextPr.Italic);
			this.sync_UnderlineCallBack(TextPr.Underline);
			this.sync_StrikeoutCallBack(TextPr.Strikeout);
			this.sync_TextPrFontSizeCallBack(TextPr.FontSize);
			this.sync_TextPrFontFamilyCallBack(TextPr.FontFamily);

			if (TextPr.VertAlign !== undefined)
				this.sync_VerticalAlign(TextPr.VertAlign);
			if (TextPr.Spacing !== undefined)
				this.sync_TextSpacing(TextPr.Spacing);
			if (TextPr.DStrikeout !== undefined)
				this.sync_TextDStrikeout(TextPr.DStrikeout);
			if (TextPr.Caps !== undefined)
				this.sync_TextCaps(TextPr.Caps);
			if (TextPr.SmallCaps !== undefined)
				this.sync_TextSmallCaps(TextPr.SmallCaps);
			if (TextPr.Position !== undefined)
				this.sync_TextPosition(TextPr.Position);
			if (TextPr.Lang !== undefined)
				this.sync_TextLangCallBack(TextPr.Lang);

			if (TextPr.Unifill !== undefined)
			{
				this.sync_TextColor2(TextPr.Unifill);
			}

			if (AscCommon.isRealObject(TextPr.HighlightColor))
			{
				var oRGB = TextPr.HighlightColor.RGBA;
				this.sendEvent("asc_onTextHighLight", new AscCommon.CColor(oRGB.R, oRGB.G, oRGB.B));
			}
			else
			{
				this.sendEvent("asc_onTextHighLight", AscCommonWord.highlight_None);
			}
		}
	};
	PDFEditorApi.prototype.sync_TextColor2 = function(unifill)
	{
		var _color;
		if (unifill.fill == null)
			return;
		var color;
		if (unifill.fill.type == Asc.c_oAscFill.FILL_TYPE_SOLID)
		{
			_color    = unifill.getRGBAColor();
			color = AscCommon.CreateAscColor(unifill.fill.color);
			color.asc_putR(_color.R);
			color.asc_putG(_color.G);
			color.asc_putB(_color.B);
			this.sendEvent("asc_onTextColor", color);
		}
		else if (unifill.fill.type == Asc.c_oAscFill.FILL_TYPE_GRAD)
		{
			_color    = unifill.getRGBAColor();
			if(unifill.fill.colors[0] && unifill.fill.colors[0].color)
			{
				color = AscCommon.CreateAscColor(unifill.fill.colors[0].color);
			}
			else
			{
				color = new Asc.asc_CColor();
			}
			color.asc_putR(_color.R);
			color.asc_putG(_color.G);
			color.asc_putB(_color.B);
			this.sendEvent("asc_onTextColor", color);
		}
		else
		{
			_color    = unifill.getRGBAColor();
			color = new Asc.asc_CColor();
			color.asc_putR(_color.R);
			color.asc_putG(_color.G);
			color.asc_putB(_color.B);
			this.sendEvent("asc_onTextColor", color);
		}
	};
	PDFEditorApi.prototype.asc_RotatePage = function(nAngle) {
		let oDoc = this.getPDFDoc();
		let oThumbnails = oDoc.Viewer.Thumbnails;

		let aPages = oThumbnails.getSelectedPages();
		oDoc.DoAction(function() {
			oDoc.RotatePages(aPages, nAngle);
			oThumbnails.keepSelectedPages = true;
		}, AscDFH.historydescription_Pdf_RotatePage, this, aPages);
	};
	PDFEditorApi.prototype.asc_GetPageRotate = function(nPage) {
		let oViewer = this.getDocumentRenderer();
		if (!oViewer) {
			return;
		}

		return oViewer.getPageRotate(nPage);
	};
	PDFEditorApi.prototype.asc_getAnchorPosition = function()
	{
		let oViewer		= editor.getDocumentRenderer();
		let pageObject	= oViewer.getPageByCoords(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
		let nPage		= pageObject ? pageObject.index : oViewer.currentPage;

		let nNativeW		= oViewer.file.pages[nPage].W;
		let nNativeH		= oViewer.file.pages[nPage].H;
		let nPageRotate		= oViewer.getPageRotate(nPage);
		let nCommentWidth	= 20;
		let nCommentHeight	= 20;
		let oDoc			= oViewer.getPDFDoc();

		let oBasePos = {
			x: 10,
			y: 10
		}

		switch (nPageRotate) {
			case 90: {
				oBasePos.y = nNativeH - nCommentHeight;
				break;
			}
			case 180: {
				oBasePos.x = nNativeW - nCommentWidth;
				oBasePos.y = nNativeH - nCommentHeight;
				break;
			}
			case 270: {
				oBasePos.x = nNativeW - nCommentWidth;
				break;
			}
		}

		if (pageObject && (pageObject.x > nNativeW || pageObject.x < 0 || pageObject.y > nNativeH || pageObject.y < 0)) {
			pageObject = null;
		}

		if (!pageObject) {
			let oPos = AscPDF.GetGlobalCoordsByPageCoords(oBasePos.x, oBasePos.y, nPage);
			oDoc.anchorPositionToAdd = {
				x: oBasePos.x,
				y: oBasePos.y
			};
			return new AscCommon.asc_CRect(oPos["X"] + nCommentWidth, oPos["Y"] + nCommentHeight / 2, 0, 0);
		}

		switch (nPageRotate) {
			case 0: {
				oDoc.anchorPositionToAdd = {
					x: pageObject.x,
					y: pageObject.y
				};
				break;
			}
			case 90: {
				oDoc.anchorPositionToAdd = {
					x: pageObject.x,
					y: pageObject.y - nCommentHeight / 4
				};
				break;
			}
			case 180: {
				oDoc.anchorPositionToAdd = {
					x: pageObject.x - nCommentWidth / 4,
					y: pageObject.y - nCommentHeight / 4
				};
				break;
			}
			case 270: {
				oDoc.anchorPositionToAdd = {
					x: pageObject.x - nCommentWidth / 4,
					y: pageObject.y
				};
				break;
			}
		}

		

		if (oDoc.mouseDownAnnot) {
			let aRect = oDoc.mouseDownAnnot.GetOrigRect();
			let oPos = AscPDF.GetGlobalCoordsByPageCoords(aRect[2], aRect[1] + (aRect[3] - aRect[1]) / 2, nPage);
			return new AscCommon.asc_CRect(oPos["X"], oPos["Y"], 0, 0);
		}
		
		let oPos = AscPDF.GetGlobalCoordsByPageCoords(pageObject.x, pageObject.y, nPage);

		return new AscCommon.asc_CRect(oPos["X"], oPos["Y"], 0, 0);
	};
	PDFEditorApi.prototype.asc_removeComment = function(Id)
	{
		let oDoc = this.getPDFDoc();
		if (!oDoc)
			return;

		oDoc.RemoveComment(Id);
	};
	PDFEditorApi.prototype.asc_remove = function() {
		let oDoc = this.getPDFDoc();
		oDoc.DoAction(function() {
			oDoc.Remove(1, false);
		}, AscDFH.historydescription_Pdf_ContextMenuRemove);
	};
	PDFEditorApi.prototype.asc_changeComment = function(Id, AscCommentData)
	{
		var oDoc = this.getDocumentRenderer().getPDFDoc();
		if (!oDoc)
			return;

		var CommentData = new AscCommon.CCommentData();
		CommentData.Read_FromAscCommentData(AscCommentData);

		oDoc.DoAction(function() {
			oDoc.EditComment(Id, CommentData);
		}, AscDFH.historydescription_Document_ChangeComment, null, Id);
	};
	PDFEditorApi.prototype.asc_selectComment = function(Id)
	{
		this.getPDFDoc().GoToAnnot(Id);
	};

	PDFEditorApi.prototype.asc_EditSelectAll = function()
	{
		let oViewer			= this.getDocumentRenderer();
		let oDoc			= oViewer.getPDFDoc();
		let oActiveForm		= oDoc.activeForm;
        let oActiveAnnot	= oDoc.mouseDownAnnot;

		if (oActiveForm && oActiveForm.IsCanEditText()) {
			oActiveForm.SelectAllText();
		}
		else if (oActiveAnnot && oActiveAnnot.IsFreeText() && oActiveAnnot.IsInTextBox()) {
            oActiveAnnot.SelectAllText();
		}
		else {
			oViewer.file.selectAll();
		}
		
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

		oViewer.file.removeSelection();

		oDoc.BlurActiveObject();
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
	PDFEditorApi.prototype.UpdateInterfaceState = function() {
		let oDoc = this.getPDFDoc();
		if (oDoc)
			oDoc.UpdateInterface();
	};
	PDFEditorApi.prototype.asc_ShowDocumentOutline = function() {
	};
	PDFEditorApi.prototype.asc_HideDocumentOutline = function() {
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

			documentRenderer.isDocumentContentReady = true;
			_t._openDocumentEndCallback();
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

		oViewer.file.removeSelection()
	};
	PDFEditorApi.prototype.SetStrikeout = function(r, g, b, opacity) {
		let oViewer	= this.getDocumentRenderer();
		let oDoc	= this.getPDFDoc();
		oDoc.SetStrikeout(r, g, b, opacity);

		oViewer.file.removeSelection();
	};
	PDFEditorApi.prototype.SetUnderline = function(r, g, b, opacity) {
		let oViewer	= this.getDocumentRenderer();
		let oDoc	= this.getPDFDoc();
		oDoc.SetUnderline(r, g, b, opacity);

		oViewer.file.removeSelection();
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
		
		let oDoc = this.getPDFDoc();
		oDoc.SearchEngine.Show = isShow;
		this.DocumentRenderer.onUpdateOverlay();
	};
	PDFEditorApi.prototype.canEnterText = function()
	{
		return this.canEdit();
	};
	PDFEditorApi.prototype.asc_GetTableOfContentsPr = function()
	{
		return null;
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

		if (window["AscDesktopEditor"])
		{
			let isCloud = !this.isLocalMode() && window["AscDesktopEditor"]["emulateCloudPrinting"];
			if (isCloud)
				window["AscDesktopEditor"]["emulateCloudPrinting"](true);

			let changes = viewer.Save();

			if (isCloud)
				window["AscDesktopEditor"]["emulateCloudPrinting"](false);

			window["AscDesktopEditor"]["Print"](JSON.stringify(desktopOptions), viewer.savedPassword ? viewer.savedPassword : "",
				changes ? AscCommon.Base64.encode(changes) : "", this.DocumentUrl);
		}

		return true;
	};
	PDFEditorApi.prototype.asyncImagesDocumentEndLoaded = function() {
		this.ImageLoader.bIsLoadDocumentFirst = false;
		var _bIsOldPaste                      = this.isPasteFonts_Images;

		// на методе _openDocumentEndCallback может поменяться this.EndActionLoadImages
		if (this.EndActionLoadImages == 1) {
			this.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadDocumentImages);
		}
		else if (this.EndActionLoadImages == 2) {
			if (_bIsOldPaste)
				this.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadImage);
			else
				this.sync_EndAction(Asc.c_oAscAsyncActionType.Information, Asc.c_oAscAsyncAction.LoadImage);
		}
		this.EndActionLoadImages = 0;

		// размораживаем меню... и начинаем считать документ
		if (false === this.isPasteFonts_Images && false === this.isSaveFonts_Images && false === this.isLoadImagesCustom) {
			this.ServerImagesWaitComplete = true;
			this._openDocumentEndCallback();
		}
		else {
			if (this.isPasteFonts_Images) {
				this.isPasteFonts_Images = false;
				this.pasteImageMap       = null;
				this.pasteCallback();
				this.pasteCallback            = null;
				this.decrementCounterLongAction();
			}
			else if (this.isSaveFonts_Images) {
				this.isSaveFonts_Images = false;
				this.saveImageMap       = null;
				this.pre_SaveCallback();

				if (this.bInit_word_control === false)
				{
					this.bInit_word_control = true;
					this.onDocumentContentReady();
				}
			}
			else if (this.isLoadImagesCustom) {
				this.isLoadImagesCustom = false;
				this.loadCustomImageMap = null;

				if (!this.ImageLoader.bIsAsyncLoadDocumentImages)
					this.SyncLoadImages_callback();
			}
		}
	};
	PDFEditorApi.prototype.onDocumentContentReady = function() {
		AscCommon.DocumentEditorApi.prototype.onDocumentContentReady.call(this);

		let thumbnailsDivId = "thumbnails-list";
		if (document.getElementById(thumbnailsDivId)) {
			this.DocumentRenderer.Thumbnails = new AscCommon.ThumbnailsControl(thumbnailsDivId);
			this.DocumentRenderer.setThumbnailsControl(this.DocumentRenderer.Thumbnails);
			
			this.DocumentRenderer.Thumbnails.registerEvent("onZoomChanged", function (value) {
				this.sendEvent("asc_onViewerThumbnailsZoomUpdate", value);
			});
		}
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

	PDFEditorApi.prototype.getSelectionState = function()
	{
		return null;
	};
	PDFEditorApi.prototype.getSpeechDescription = function(prevState, action)
	{
		return null;
	};
	PDFEditorApi.prototype.GenerateStyles = function() {};
	
	PDFEditorApi.prototype.Resize = function() {
		if (!this.DocumentRenderer)
			return;
		this.DocumentRenderer.resize();
	};
	PDFEditorApi.prototype._openDocumentEndCallback = function() {
		if (this.isDocumentLoadComplete || !this.ServerImagesWaitComplete || !this.ServerIdWaitComplete || !this.WordControl || !this.WordControl.m_oLogicDocument || !this.WordControl.m_oDrawingDocument.m_oDocumentRenderer.isDocumentContentReady)
			return;

		this.sendMathToMenu();
		this.sendStandartTextures();
		//выставляем тип copypaste
		this.isDocumentEditor = false;
		AscCommon.PasteElementsId.g_bIsDocumentCopyPaste = false;
		AscCommon.PasteElementsId.g_bIsPDFCopyPaste = true;
		
		if (this.isApplyChangesOnOpenEnabled)
		{
			if (AscCommon.EncryptionWorker)
			{
				AscCommon.EncryptionWorker.init();
				if (!AscCommon.EncryptionWorker.isChangesHandled)
					return AscCommon.EncryptionWorker.handleChanges(AscCommon.CollaborativeEditing.m_aChanges, this, this._openDocumentEndCallback);
			}
			this.bInit_word_control = true;
			this.isApplyChangesOnOpenEnabled = false;
			this._applyPreOpenLocks();

			// TODO: onDocumentContentReady вызываем в конце загрузки всех изменений (и объектов для этих изменений)
			let oThis = this;
			
			// Принимаем изменения на открытии только если это редактор, либо LiveViewer (т.е. включена быстрая совместка)
			if (this.isLiveViewer() || !this.isPdfViewer)
			{
				let perfStart    = performance.now();
				let OtherChanges = AscCommon.CollaborativeEditing.Have_OtherChanges();
				AscCommon.CollaborativeEditing.Apply_Changes(function()
				{
					let perfEnd = performance.now();
					if (OtherChanges)
					{
						AscCommon.sendClientLog("debug", AscCommon.getClientInfoString("onApplyChanges", perfEnd - perfStart), oThis);
					}
					oThis.onDocumentContentReady();
				});
				AscCommon.CollaborativeEditing.Release_Locks();
			}
			else
			{
				this.onDocumentContentReady();
			}

			this.isApplyChangesOnOpen = true;
		}
	};
	PDFEditorApi.prototype._canSyncCollaborativeChanges = function(isFirstLoad)
	{
		return (!this.isPdfViewer && (!isFirstLoad || this.isApplyChangesOnOpen));
	};
	PDFEditorApi.prototype.sync_CollaborativeChanges = function()
	{
		if (AscCommon.CollaborativeEditing.Is_Fast())
			return;
		
		this.sendEvent("asc_onCollaborativeChanges");
	};
	PDFEditorApi.prototype.sync_ContextMenuCallback = function(Data) {
		this.sendEvent("asc_onContextMenu", new CPdfContextMenuData(Data));
	};
	PDFEditorApi.prototype._finalizeAction = function() {
		let oDoc = this.getPDFDoc();
		if (!oDoc){
			return;
		}
		oDoc.FinalizeAction(true);
	};

	PDFEditorApi.prototype._waitPrint = function(actionType, options)
	{
		return false;
	};
	PDFEditorApi.prototype.asc_setPdfViewer = function(isPdfViewer) {
		this.isPdfViewer = isPdfViewer;
	};
	PDFEditorApi.prototype.isLiveViewer = function() {
		return this.isPdfViewer && AscCommon.CollaborativeEditing.Is_Fast() && !this.VersionHistory;
	};
	
	function CPdfContextMenuData(obj) {
		if (obj) {
			this.Type  		= ( undefined != obj.Type ) ? obj.Type : Asc.c_oAscPdfContextMenuTypes.Common;
			this.X_abs 		= ( undefined != obj.X_abs ) ? obj.X_abs : 0;
			this.Y_abs 		= ( undefined != obj.Y_abs ) ? obj.Y_abs : 0;
			this.PageNum	= ( undefined != obj.PageNum ) ? obj.PageNum : 0;
		}
		else {
			this.Type	= Asc.c_oAscPdfContextMenuTypes.Common;
			this.X_abs	= 0;
			this.Y_abs	= 0;
		}
	}

	CPdfContextMenuData.prototype.get_Type = function() {
		return this.Type;
	};
	CPdfContextMenuData.prototype.get_X = function() {
		return this.X_abs;
	};
	CPdfContextMenuData.prototype.get_Y = function() {
		return this.Y_abs;
	};
	CPdfContextMenuData.prototype.get_PageNum = function() {
		return this.PageNum;
	};
	CPdfContextMenuData.prototype.get_IsPageSelect = function() {
		return this.PageNum != -1 && this.PageNum != undefined;
	};

	/** @enum {number} */
	let c_oAscPdfContextMenuTypes = {
		Common       	: 0,	// Обычное контекстное меню
		Thumbnails		: 1		// контекстное меню тамбнейлов
	};

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Export
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	window['Asc']['PDFEditorApi'] = PDFEditorApi;
	AscCommon.PDFEditorApi        = PDFEditorApi;
	
	prot = window['Asc']['c_oAscPdfContextMenuTypes'] = window['Asc'].c_oAscPdfContextMenuTypes = c_oAscPdfContextMenuTypes;
	prot['Common']		= c_oAscPdfContextMenuTypes.Common;
	prot['Thumbnails']	= c_oAscPdfContextMenuTypes.Thumbnails;

	CPdfContextMenuData.prototype['get_Type']			= CPdfContextMenuData.prototype.get_Type;
	CPdfContextMenuData.prototype['get_X']				= CPdfContextMenuData.prototype.get_X;
	CPdfContextMenuData.prototype['get_Y']				= CPdfContextMenuData.prototype.get_Y;
	CPdfContextMenuData.prototype['get_PageNum']		= CPdfContextMenuData.prototype.get_PageNum;
	CPdfContextMenuData.prototype['get_IsPageSelect']	= CPdfContextMenuData.prototype.get_IsPageSelect;

	PDFEditorApi.prototype['sync_ContextMenuCallback']		= PDFEditorApi.prototype.sync_ContextMenuCallback;
	PDFEditorApi.prototype['sync_CanUndoCallback']			= PDFEditorApi.prototype.sync_CanUndoCallback;
	PDFEditorApi.prototype['sync_CanRedoCallback']			= PDFEditorApi.prototype.sync_CanRedoCallback;
	PDFEditorApi.prototype['asc_setAdvancedOptions']		= PDFEditorApi.prototype.asc_setAdvancedOptions;
	PDFEditorApi.prototype['startGetDocInfo']				= PDFEditorApi.prototype.startGetDocInfo;
	PDFEditorApi.prototype['stopGetDocInfo']				= PDFEditorApi.prototype.stopGetDocInfo;
	PDFEditorApi.prototype['can_CopyCut']					= PDFEditorApi.prototype.can_CopyCut;
	PDFEditorApi.prototype['asc_searchEnabled']				= PDFEditorApi.prototype.asc_searchEnabled;
	PDFEditorApi.prototype['asc_findText']					= PDFEditorApi.prototype.asc_findText;
	PDFEditorApi.prototype['asc_endFindText']				= PDFEditorApi.prototype.asc_endFindText;
	PDFEditorApi.prototype['asc_isSelectSearchingResults']	= PDFEditorApi.prototype.asc_isSelectSearchingResults;
	PDFEditorApi.prototype['asc_StartTextAroundSearch']		= PDFEditorApi.prototype.asc_StartTextAroundSearch;
	PDFEditorApi.prototype['asc_SelectSearchElement']		= PDFEditorApi.prototype.asc_SelectSearchElement;
	PDFEditorApi.prototype['ContentToHTML']					= PDFEditorApi.prototype.ContentToHTML;
	PDFEditorApi.prototype['goToPage']						= PDFEditorApi.prototype.goToPage;
	PDFEditorApi.prototype['getCountPages']					= PDFEditorApi.prototype.getCountPages;
	PDFEditorApi.prototype['getCurrentPage']				= PDFEditorApi.prototype.getCurrentPage;
	PDFEditorApi.prototype['asc_getPdfProps']				= PDFEditorApi.prototype.asc_getPdfProps;
	PDFEditorApi.prototype['asc_enterText']					= PDFEditorApi.prototype.asc_enterText;
	PDFEditorApi.prototype['asc_correctEnterText']			= PDFEditorApi.prototype.asc_correctEnterText;
	PDFEditorApi.prototype['asc_GetSelectedText']			= PDFEditorApi.prototype.asc_GetSelectedText;
	PDFEditorApi.prototype['asc_SelectPDFFormListItem']		= PDFEditorApi.prototype.asc_SelectPDFFormListItem;
	PDFEditorApi.prototype['asc_SetTextFormDatePickerDate']	= PDFEditorApi.prototype.asc_SetTextFormDatePickerDate;
	PDFEditorApi.prototype['asc_getHeaderFooterProperties']	= PDFEditorApi.prototype.asc_getHeaderFooterProperties;
	PDFEditorApi.prototype['ChangeReaderMode']				= PDFEditorApi.prototype.ChangeReaderMode;
	PDFEditorApi.prototype['asc_Save']						= PDFEditorApi.prototype.asc_Save;
	PDFEditorApi.prototype['asc_Print']						= PDFEditorApi.prototype.asc_Print;
	PDFEditorApi.prototype['asc_drawPrintPreview']			= PDFEditorApi.prototype.asc_drawPrintPreview;

	PDFEditorApi.prototype['CheckChangedDocument']		   = PDFEditorApi.prototype.CheckChangedDocument;
	PDFEditorApi.prototype['SetDrawingFreeze']             = PDFEditorApi.prototype.SetDrawingFreeze;
	PDFEditorApi.prototype['OnMouseUp']                    = PDFEditorApi.prototype.OnMouseUp;

	PDFEditorApi.prototype['asc_addComment']               = PDFEditorApi.prototype.asc_addComment;
	PDFEditorApi.prototype['can_AddQuotedComment']         = PDFEditorApi.prototype.can_AddQuotedComment;
	PDFEditorApi.prototype['asc_showComments']             = PDFEditorApi.prototype.asc_showComments;
	PDFEditorApi.prototype['asc_showComment']              = PDFEditorApi.prototype.asc_showComment;
	PDFEditorApi.prototype['asc_hideComments']             = PDFEditorApi.prototype.asc_hideComments;
	PDFEditorApi.prototype['asc_removeComment']            = PDFEditorApi.prototype.asc_removeComment;
	PDFEditorApi.prototype['asc_remove']            	   = PDFEditorApi.prototype.asc_remove;
	PDFEditorApi.prototype['asc_changeComment']            = PDFEditorApi.prototype.asc_changeComment;
	PDFEditorApi.prototype['asc_selectComment']            = PDFEditorApi.prototype.asc_selectComment;
	PDFEditorApi.prototype['asc_EditPage']                 = PDFEditorApi.prototype.asc_EditPage;
	PDFEditorApi.prototype['asc_AddPage']                  = PDFEditorApi.prototype.asc_AddPage;
	PDFEditorApi.prototype['asc_RemovePage']			   = PDFEditorApi.prototype.asc_RemovePage;
	PDFEditorApi.prototype['asc_createSmartArt']		   = PDFEditorApi.prototype.asc_createSmartArt;
	PDFEditorApi.prototype['asc_undoAllChanges']		   = PDFEditorApi.prototype.asc_undoAllChanges;

	PDFEditorApi.prototype['asc_setSkin']                  = PDFEditorApi.prototype.asc_setSkin;
	PDFEditorApi.prototype['asc_getAnchorPosition']        = PDFEditorApi.prototype.asc_getAnchorPosition;
	PDFEditorApi.prototype['asc_GetPageRotate']			   = PDFEditorApi.prototype.asc_GetPageRotate;
	PDFEditorApi.prototype['asc_RotatePage']          	   = PDFEditorApi.prototype.asc_RotatePage;
	PDFEditorApi.prototype['SetMarkerFormat']              = PDFEditorApi.prototype.SetMarkerFormat;
	PDFEditorApi.prototype['get_PageWidth']                = PDFEditorApi.prototype.get_PageWidth;
	PDFEditorApi.prototype['get_PageHeight']               = PDFEditorApi.prototype.get_PageHeight;
	PDFEditorApi.prototype['asc_EditSelectAll']            = PDFEditorApi.prototype.asc_EditSelectAll;
	PDFEditorApi.prototype['Undo']                         = PDFEditorApi.prototype.Undo;
	PDFEditorApi.prototype['Redo']                         = PDFEditorApi.prototype.Redo;
	PDFEditorApi.prototype['UpdateInterfaceState']         = PDFEditorApi.prototype.UpdateInterfaceState;
	PDFEditorApi.prototype['asc_SelectionCut']             = PDFEditorApi.prototype.asc_SelectionCut;
	PDFEditorApi.prototype['asc_CheckCopy']                = PDFEditorApi.prototype.asc_CheckCopy;
	PDFEditorApi.prototype['Paste']                        = PDFEditorApi.prototype.Paste;
	PDFEditorApi.prototype['asc_PasteData']                = PDFEditorApi.prototype.asc_PasteData;
	PDFEditorApi.prototype['asc_CloseFile']                = PDFEditorApi.prototype.asc_CloseFile;

	PDFEditorApi.prototype['getSelectionState']            = PDFEditorApi.prototype.getSelectionState;
	PDFEditorApi.prototype['getSpeechDescription']         = PDFEditorApi.prototype.getSpeechDescription;

	PDFEditorApi.prototype['asc_ShowDocumentOutline']      = PDFEditorApi.prototype.asc_ShowDocumentOutline;
	PDFEditorApi.prototype['asc_HideDocumentOutline']      = PDFEditorApi.prototype.asc_HideDocumentOutline;

	// text/para pr
	PDFEditorApi.prototype['put_TextPrBold']				= PDFEditorApi.prototype.put_TextPrBold;
	PDFEditorApi.prototype['put_TextPrItalic']				= PDFEditorApi.prototype.put_TextPrItalic;
	PDFEditorApi.prototype['put_TextPrUnderline']			= PDFEditorApi.prototype.put_TextPrUnderline;
	PDFEditorApi.prototype['put_TextPrStrikeout']			= PDFEditorApi.prototype.put_TextPrStrikeout;
	PDFEditorApi.prototype['put_PrLineSpacing']				= PDFEditorApi.prototype.put_PrLineSpacing;
	PDFEditorApi.prototype['put_LineSpacingBeforeAfter']	= PDFEditorApi.prototype.put_LineSpacingBeforeAfter;
	PDFEditorApi.prototype['FontSizeIn']					= PDFEditorApi.prototype.FontSizeIn;
	PDFEditorApi.prototype['FontSizeOut']					= PDFEditorApi.prototype.FontSizeOut;
	PDFEditorApi.prototype['put_TextPrBaseline']			= PDFEditorApi.prototype.put_TextPrBaseline;
	PDFEditorApi.prototype['put_TextPrFontSize']			= PDFEditorApi.prototype.put_TextPrFontSize;
	PDFEditorApi.prototype['put_TextPrFontName']			= PDFEditorApi.prototype.put_TextPrFontName;
	PDFEditorApi.prototype['put_TextColor']					= PDFEditorApi.prototype.put_TextColor;
	PDFEditorApi.prototype['asc_ChangeTextCase']			= PDFEditorApi.prototype.asc_ChangeTextCase;
	PDFEditorApi.prototype['put_PrAlign']					= PDFEditorApi.prototype.put_PrAlign;
	PDFEditorApi.prototype['setVerticalAlign']				= PDFEditorApi.prototype.setVerticalAlign;
	PDFEditorApi.prototype['IncreaseIndent']				= PDFEditorApi.prototype.IncreaseIndent;
	PDFEditorApi.prototype['DecreaseIndent']				= PDFEditorApi.prototype.DecreaseIndent;
	PDFEditorApi.prototype['ClearFormating']				= PDFEditorApi.prototype.ClearFormating;
	PDFEditorApi.prototype['UpdateParagraphProp']			= PDFEditorApi.prototype.UpdateParagraphProp;
	PDFEditorApi.prototype['paraApply']						= PDFEditorApi.prototype.paraApply;
	PDFEditorApi.prototype['sync_ListType']					= PDFEditorApi.prototype.sync_ListType;
	PDFEditorApi.prototype['put_ListType']					= PDFEditorApi.prototype.put_ListType;
	PDFEditorApi.prototype['asc_GetPossibleNumberingLanguage']= PDFEditorApi.prototype.asc_GetPossibleNumberingLanguage;

	// math
	PDFEditorApi.prototype['asc_AddMath2']			= PDFEditorApi.prototype.asc_AddMath2;
	PDFEditorApi.prototype['asc_ConvertMathView']	= PDFEditorApi.prototype.asc_ConvertMathView;

	// highlight annots
	PDFEditorApi.prototype['asc_SetFillColor']		= PDFEditorApi.prototype.asc_SetFillColor;
	PDFEditorApi.prototype['asc_GetFillColor']		= PDFEditorApi.prototype.asc_GetFillColor;
	PDFEditorApi.prototype['asc_SetStrokeColor']	= PDFEditorApi.prototype.asc_SetStrokeColor;
	PDFEditorApi.prototype['asc_GetStrokeColor']	= PDFEditorApi.prototype.asc_GetStrokeColor;
	PDFEditorApi.prototype['asc_SetOpacity']		= PDFEditorApi.prototype.asc_SetOpacity;
	PDFEditorApi.prototype['asc_GetOpacity']		= PDFEditorApi.prototype.asc_GetOpacity;
	// stamp
	PDFEditorApi.prototype['AddStampAnnot']					= PDFEditorApi.prototype.AddStampAnnot;
	PDFEditorApi.prototype['asc_getPropertyEditorStamps']	= PDFEditorApi.prototype.asc_getPropertyEditorStamps;

	// freetext
	PDFEditorApi.prototype['AddFreeTextAnnot']	= PDFEditorApi.prototype.AddFreeTextAnnot;

	// drawings
	PDFEditorApi.prototype['AddTextArt']							= PDFEditorApi.prototype.AddTextArt;
	PDFEditorApi.prototype['StartAddShape']							= PDFEditorApi.prototype.StartAddShape;
	PDFEditorApi.prototype['ShapeApply']							= PDFEditorApi.prototype.ShapeApply;
	PDFEditorApi.prototype['ChangeShapeType']						= PDFEditorApi.prototype.ChangeShapeType;
	PDFEditorApi.prototype['ImgApply']								= PDFEditorApi.prototype.ImgApply;
	PDFEditorApi.prototype['asc_FitImagesToPage']					= PDFEditorApi.prototype.asc_FitImagesToPage;
	PDFEditorApi.prototype['sync_shapePropCallback']				= PDFEditorApi.prototype.sync_shapePropCallback;
	PDFEditorApi.prototype['sync_annotPropCallback']				= PDFEditorApi.prototype.sync_annotPropCallback;
	PDFEditorApi.prototype['canUnGroup']							= PDFEditorApi.prototype.canUnGroup;
	PDFEditorApi.prototype['canGroup']								= PDFEditorApi.prototype.canGroup;
	PDFEditorApi.prototype['shapes_bringToFront']					= PDFEditorApi.prototype.shapes_bringToFront;
	PDFEditorApi.prototype['shapes_bringForward']					= PDFEditorApi.prototype.shapes_bringForward;
	PDFEditorApi.prototype['shapes_bringToBack']					= PDFEditorApi.prototype.shapes_bringToBack;
	PDFEditorApi.prototype['shapes_bringBackward']					= PDFEditorApi.prototype.shapes_bringBackward;
	PDFEditorApi.prototype['AddImageUrlAction']						= PDFEditorApi.prototype.AddImageUrlAction;
	PDFEditorApi.prototype['AddImageUrlActionCallback']				= PDFEditorApi.prototype.AddImageUrlActionCallback;
	PDFEditorApi.prototype['asc_addImage']							= PDFEditorApi.prototype.asc_addImage;
	PDFEditorApi.prototype['ChangeArtImageFromFile']				= PDFEditorApi.prototype.ChangeArtImageFromFile;
	PDFEditorApi.prototype['SetInterfaceDrawImagePlaceTextArt']		= PDFEditorApi.prototype.SetInterfaceDrawImagePlaceTextArt;
	PDFEditorApi.prototype['asc_setInterfaceDrawImagePlaceShape']	= PDFEditorApi.prototype.asc_setInterfaceDrawImagePlaceShape;
	PDFEditorApi.prototype['remove_Hyperlink']						= PDFEditorApi.prototype.remove_Hyperlink;
	PDFEditorApi.prototype['change_Hyperlink']						= PDFEditorApi.prototype.change_Hyperlink;
	PDFEditorApi.prototype['sync_HyperlinkClickCallback']			= PDFEditorApi.prototype.sync_HyperlinkClickCallback;
	PDFEditorApi.prototype['SetShowTextSelectPanel']				= PDFEditorApi.prototype.SetShowTextSelectPanel;
	PDFEditorApi.prototype['NeedShowTextSelectPanel']				= PDFEditorApi.prototype.NeedShowTextSelectPanel;

	// table
	PDFEditorApi.prototype['put_Table']						= PDFEditorApi.prototype.put_Table;
	PDFEditorApi.prototype['tblApply']						= PDFEditorApi.prototype.tblApply;
	PDFEditorApi.prototype['asc_DistributeTableCells']		= PDFEditorApi.prototype.asc_DistributeTableCells;
	PDFEditorApi.prototype['remColumn']						= PDFEditorApi.prototype.remColumn;
	PDFEditorApi.prototype['remTable']						= PDFEditorApi.prototype.remTable;
	PDFEditorApi.prototype['asc_getTableStylesPreviews']	= PDFEditorApi.prototype.asc_getTableStylesPreviews;
	PDFEditorApi.prototype['asc_GetSelectionBounds']		= PDFEditorApi.prototype.asc_GetSelectionBounds;
	PDFEditorApi.prototype['asc_setPdfViewer']		        = PDFEditorApi.prototype.asc_setPdfViewer;

	PDFEditorApi.prototype['asc_GetTableOfContentsPr']      = PDFEditorApi.prototype.asc_GetTableOfContentsPr;

})(window, window.document);
