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
	const c_oAscError             = Asc.c_oAscError;
	const c_oAscFileType          = Asc.c_oAscFileType;
	const c_oAscAsyncAction       = Asc.c_oAscAsyncAction;
	const c_oAscAsyncActionType   = Asc.c_oAscAsyncActionType;
	const c_oAscFontRenderingModeType = Asc.c_oAscFontRenderingModeType;

	/**
	 *
	 * @param config
	 * @constructor
	 * @extends {AscCommon.baseEditorsApi}
	 */
	function VisioEditorApi(config)
	{
		AscCommon.baseEditorsApi.call(this, config, AscCommon.c_oEditorId.Visio);

		this.WordControl = null;

		this.documentFormatSave = c_oAscFileType.VSDX;

		this.tmpIsFreeze   = null;
		this.tmpZoomType   = null;
		this.tmpDocumentUnits = null;
		this.tmpFontRenderingMode = null;
		/**
		 *
		 * @type {CVisioDocument}
		 */
		this.Document = null;

		/**		 * @type {HTMLDivElement}		*/
		this.HtmlElement = null;

		/**		 * @type {HTMLCanvasElement}	*/
		this.canvas = null;

		/**		 * @type {HTMLCanvasElement}	*/
		this.thumbnailsCanvas = null;
		this.locale = null;
		this.bInit_word_control = false;

		if (window.editor == undefined)
		{
			window.editor = this;
			window['editor'] = window.editor;
			Asc['editor'] = Asc.editor = this;

			if (window["NATIVE_EDITOR_ENJINE"])
				editor = window.editor;
		}

		this._init();
		return this;
	}

	VisioEditorApi.prototype = Object.create(AscCommon.baseEditorsApi.prototype);
	VisioEditorApi.prototype.constructor = VisioEditorApi;

	VisioEditorApi.prototype.initDefaultShortcuts = function()
	{
		// [[ActionType, KeyCode, Ctrl, Shift, Alt]]
		var aShortcuts =
			[
				[Asc.c_oAscDiagramShortcutType.Print, 80, true, false, false]
			];
		this.initShortcuts(aShortcuts, false)
	};
	VisioEditorApi.prototype.InitEditor = function(){
		this.Document = new AscVisio.CVisioDocument(this, this.WordControl.m_oDrawingDocument);

		this.WordControl.m_oLogicDocument = this.Document;
		this.WordControl.m_oDrawingDocument.m_oLogicDocument = this.WordControl.m_oLogicDocument;
	};
	VisioEditorApi.prototype._onEndLoadSdk  = function()
	{
		AscCommon.baseEditorsApi.prototype._onEndLoadSdk.call(this);

		AscFonts.g_fontApplication.Init();

		this.FontLoader  = AscCommon.g_font_loader;
		this.ImageLoader = AscCommon.g_image_loader;
		this.FontLoader.put_Api(this);
		this.ImageLoader.put_Api(this);

		this._loadSdkImages();

		this.WordControl      = new AscVisio.CEditorPage(this);
		this.WordControl.Name = this.HtmlElementName;
		this.CreateComponents();
		this.WordControl.Init();


		if (AscCommon.g_oTextMeasurer.SetParams)
		{
			AscCommon.g_oTextMeasurer.SetParams({ mode : "slide" });
		}

		if (this.tmpFontRenderingMode)
		{
			this.SetFontRenderingMode(this.tmpFontRenderingMode);
		}
		if (null !== this.tmpIsFreeze)
		{
			this.SetDrawingFreeze(this.tmpIsFreeze);
		}
		if (null !== this.tmpZoomType)
		{
			switch (this.tmpZoomType)
			{
				case AscCommon.c_oZoomType.FitToPage:
					this.zoomFitToPage();
					break;
				case AscCommon.c_oZoomType.FitToWidth:
					this.zoomFitToWidth();
					break;
				case AscCommon.c_oZoomType.CustomMode:
					this.zoomCustomMode();
					break;
			}
		}
		if (null != this.tmpDocumentUnits)
		{
			this.asc_SetDocumentUnits(this.tmpDocumentUnits);
			this.tmpDocumentUnits = null;
		}

		this.asc_setViewMode(this.isViewMode);
	};
	VisioEditorApi.prototype.CreateCSS = function()
	{
		var _head = document.getElementsByTagName('head')[0];

		var style0       = document.createElement('style');
		style0.type      = 'text/css';
		style0.innerHTML = ".block_elem { position:absolute;padding:0;margin:0; }";
		_head.appendChild(style0);
	};
	VisioEditorApi.prototype.CreateComponents = function()
	{
		this.asc_setSkin(this.skinObject);
		delete this.skinObject;

		this.CreateCSS();

		//stubs for html page
		window.CLayoutThumbnailDrawer = function () {};
		window.CMasterThumbnailDrawer = function () {};
		var _innerHTML = "<div id=\"id_panel_thumbnails\" class=\"block_elem\" style=\"touch-action:none;-webkit-touch-callout:none;background-color:" + AscCommon.GlobalSkin.BackgroundColorThumbnails + ";\">\
									<div id=\"id_panel_thumbnails_split\" class=\"block_elem\" style=\"pointer-events:none;background-color:" + AscCommon.GlobalSkin.BackgroundColorThumbnails + ";\"></div>\
		                            <canvas id=\"id_thumbnails_background\" class=\"block_elem\" style=\"-ms-touch-action: none;-webkit-user-select: none;z-index:1\"></canvas>\
		                            <canvas id=\"id_thumbnails\" class=\"block_elem\" style=\"-ms-touch-action: none;-webkit-user-select: none;z-index:2\"></canvas>\
		                            <div id=\"id_vertical_scroll_thmbnl\" style=\"left:0;top:0;width:1px;overflow:hidden;position:absolute;\">\
									</div>\
		                        </div>\
		                    <div id=\"id_main_parent\" class=\"block_elem\" style=\"width:100%;height:100%;touch-action:none;-ms-touch-action: none;-moz-user-select:none;-khtml-user-select:none;user-select:none;overflow:hidden;border-left-width: 1px;border-left-color:" + AscCommon.GlobalSkin.BorderSplitterColor + "; border-left-style: solid;\" UNSELECTABLE=\"on\">\
                            <div id=\"id_main\" class=\"block_elem\" style=\"width:100%;height:100%;z-index:5;-ms-touch-action: none;-moz-user-select:none;-khtml-user-select:none;user-select:none;background-color:" + AscCommon.GlobalSkin.BackgroundColor + ";overflow:hidden;\" UNSELECTABLE=\"on\">\
								<div id=\"id_panel_left\" class=\"block_elem\">\
									<canvas id=\"id_buttonTabs\" class=\"block_elem\"></canvas>\
									<canvas id=\"id_vert_ruler\" class=\"block_elem\"></canvas>\
								</div>\
                                <div id=\"id_panel_top\" class=\"block_elem\">\
									<canvas id=\"id_hor_ruler\" class=\"block_elem\"></canvas>\
                                </div>\
                                <div id=\"id_main_view\" class=\"block_elem\" style=\"width:100%;height:100%;overflow:hidden\">\
                                    <canvas id=\"id_viewer\" class=\"block_elem\" style=\"-ms-touch-action: none;-webkit-user-select: none;background-color:" + AscCommon.GlobalSkin.BackgroundColor + ";z-index:6\"></canvas>\
                                    <canvas id=\"id_viewer_overlay\" class=\"block_elem\" style=\"-ms-touch-action: none;-webkit-user-select: none;z-index:7\"></canvas>\
                                    <div id=\"id_target_cursor\" class=\"block_elem\" width=\"1\" height=\"1\" style=\"-ms-touch-action: none;-webkit-user-select: none;width:2px;height:13px;display:none;z-index:9;\"></div>\
                                </div>\
							    <div id=\"id_panel_right\" class=\"block_elem\" style=\"margin-right:1px;background-color:" + AscCommon.GlobalSkin.BackgroundColor + ";z-index:0;\">\
							        <div id=\"id_buttonRulers\" class=\"block_elem buttonRuler\"></div>\
								    <div id=\"id_vertical_scroll\" style=\"left:0;top:0;width:14px;overflow:hidden;position:absolute;\">\
								    </div>\
								    <div id=\"id_buttonPrevPage\" class=\"block_elem buttonPrevPage\"></div>\
								    <div id=\"id_buttonNextPage\" class=\"block_elem buttonNextPage\"></div>\
                                </div>\
                                <div id=\"id_horscrollpanel\" class=\"block_elem\" style=\"margin-bottom:1px;background-color:" + AscCommon.GlobalSkin.BackgroundColor + ";z-index:0;\">\
                                    <div id=\"id_horizontal_scroll\" style=\"left:0;top:0;height:14px;overflow:hidden;position:absolute;width:100%;\">\
                                    </div>\
                                </div>\
                            </div>";
		if (this.HtmlElement)
			_innerHTML += this.HtmlElement.innerHTML;

		if (this.HtmlElement != null)
		{
			this.HtmlElement.style.backgroundColor = AscCommon.GlobalSkin.BackgroundColor;
			this.HtmlElement.innerHTML = _innerHTML;
		}

		this.canvas = document.getElementById("id_viewer");
	};
	// работа с шрифтами
	VisioEditorApi.prototype.asyncFontsDocumentStartLoaded = function(blockType)
	{
		this.sync_StartAction(undefined === blockType ? Asc.c_oAscAsyncActionType.BlockInteraction : blockType, Asc.c_oAscAsyncAction.LoadDocumentFonts);
		var _progress         = this.OpenDocumentProgress;
		_progress.Type        = Asc.c_oAscAsyncAction.LoadDocumentFonts;
		_progress.FontsCount  = this.FontLoader.fonts_loading.length;
		_progress.CurrentFont = 0;

		var _loader_object = this.WordControl.m_oLogicDocument;
		var _count         = 0;
		if (_loader_object !== undefined && _loader_object != null)
		{
			for (var i in _loader_object.ImageMap)
			{
				++_count;
			}
		}

		_progress.ImagesCount  = _count;
		_progress.CurrentImage = 0;
	};
	VisioEditorApi.prototype.asyncFontsDocumentEndLoaded   = function(blockType) {
		this.sync_EndAction(undefined === blockType ? Asc.c_oAscAsyncActionType.BlockInteraction : blockType, Asc.c_oAscAsyncAction.LoadDocumentFonts);

		this.EndActionLoadImages = 0;

		if (null != this.WordControl.m_oLogicDocument)
		{
			//this.WordControl.m_oDrawingDocument.CheckGuiControlColors();
			// this.sendColorThemes(this.WordControl.m_oLogicDocument.themes[0]);
		}

		// открытие после загрузки документа

		var _loader_object = this.WordControl.m_oLogicDocument;
		if (null == _loader_object)
			_loader_object = this.WordControl.m_oDrawingDocument.m_oDocumentRenderer;

		var _count = 0;
		for (var i in _loader_object.ImageMap)
			++_count;

		if (_count > 0)
		{
			this.EndActionLoadImages = 1;
			this.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadDocumentImages);
		}

		this.ImageLoader.bIsLoadDocumentFirst = true;
		this.ImageLoader.LoadDocumentImages(_loader_object.ImageMap);
	};
	VisioEditorApi.prototype.asyncImagesDocumentEndLoaded = function()
	{
		this.ImageLoader.bIsLoadDocumentFirst = false;

		// на методе _openDocumentEndCallback может поменяться this.EndActionLoadImages
		if (this.EndActionLoadImages == 1)
		{
			this.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadDocumentImages);
		}
		this.EndActionLoadImages = 0;

		this.ServerImagesWaitComplete = true;
		this._openDocumentEndCallback();
	};
	VisioEditorApi.prototype._openDocumentEndCallback = function()
	{
		if (this.isDocumentLoadComplete || !this.ServerImagesWaitComplete || !this.ServerIdWaitComplete || !this.WordControl || !this.WordControl.m_oLogicDocument)
			return;

		if (this.isViewMode)
			this.asc_setViewMode(true);

		this.WordControl.m_oLogicDocument.Recalculate({Drawings : {All : true, Map : {}}});
		AscCommon.History.private_ClearRecalcData();
		this.WordControl.m_oLogicDocument.DrawingDocument.OnEndRecalculate();
		this.WordControl.GoToPage(0);

		this.LoadedObject       = null;
		this.bInit_word_control = true;
		this.onDocumentContentReady();

		this.WordControl.InitControl();

		if (this.isViewMode)
			this.asc_setViewMode(true);

		// Меняем тип состояния (на никакое)
		this.advancedOptionsAction = AscCommon.c_oAscAdvancedOptionsAction.None;

		this.WordControl.GoToPage(this.Document.getCurrentPage());
	};
	VisioEditorApi.prototype._coAuthoringInitEnd = function()
	{
		//todo other events
		var t                                        = this;
		this.CoAuthoringApi.onConnectionStateChanged = function(e)
		{
			// if (true === AscCommon.CollaborativeEditing.Is_Fast() && false === e['state'])
			// {
			// 	editor.WordControl.m_oLogicDocument.Remove_ForeignCursor(e['id']);
			// }
			t.sendEvent("asc_onConnectionStateChanged", e);
		};
	};

	VisioEditorApi.prototype.OpenDocumentFromZip = function(data)
	{
		this.DocumentType = 2;

		AscCommon.g_oIdCounter.Set_Load(true);
		let res = this.OpenDocumentFromZipNoInit(data);
		AscCommon.g_oIdCounter.Set_Load(false);
		this.WordControl.m_oLogicDocument.Set_FastCollaborativeEditing(true);

		this.LoadedObject = 1;
		AscFonts.IsCheckSymbols = false;

		this.Document.loadFonts();

		if (this.isMobileVersion)
		{
			AscCommon.AscBrowser.isSafariMacOs   = false;
			AscCommon.PasteElementsId.PASTE_ELEMENT_ID     = "wrd_pastebin";
			AscCommon.PasteElementsId.ELEMENT_DISPAY_STYLE = "none";
		}
		if (AscCommon.AscBrowser.isSafariMacOs)
			setInterval(AscCommon.SafariIntervalFocus, 10);
		return res;
	};
	VisioEditorApi.prototype.OpenDocumentFromZipNoInit = function(data)
	{
		if (!data) {
			return false;
		}
		let xmlParserContext = new AscCommon.XmlParserContext();

		let jsZlib = new AscCommon.ZLib();
		if (!jsZlib.open(data)) {
			return false;
		}

		this.InitEditor();
		this.Document.fromZip(jsZlib, xmlParserContext);

		// var context = reader.context;
		// this.WordControl.m_oLogicDocument.ImageMap = context.loadDataLinks();
		AscCommon.pptx_content_loader.Reader.ImageMapChecker = AscCommon.pptx_content_loader.ImageMapChecker;
		this.Document.ImageMap = xmlParserContext.loadDataLinks();
		this.ServerIdWaitComplete = true;

		jsZlib.close();
		return true;
	};
	VisioEditorApi.prototype.asc_CloseFile            = function()
	{
		AscCommon.History.Clear();
		AscCommon.g_oTableId.Clear();
		AscCommon.g_oIdCounter.Clear();
		this.isApplyChangesOnOpenEnabled = true;
		this.isDocumentLoadComplete = false;
		this.turnOffSpecialModes();
		AscCommon.pptx_content_loader.ImageMapChecker = {};
	};
	VisioEditorApi.prototype.asc_getAppProps = function()
	{
		return this.WordControl && this.WordControl.m_oLogicDocument && this.WordControl.m_oLogicDocument.app || null;
	};
	VisioEditorApi.prototype.getInternalCoreProps = function()
	{
		return this.WordControl && this.WordControl.m_oLogicDocument && this.WordControl.m_oLogicDocument.core;
	};
	VisioEditorApi.prototype.asc_CheckCopy = function()
	{
		//todo
		return;
	};
	VisioEditorApi.prototype.asc_PasteData = function(_format, data1, data2, text_data, useCurrentPoint, callback, checkLocks) {
		if (!this.canEdit())
			return;
		//todo
		return;
	}
	VisioEditorApi.prototype.onKeyDown = function(e)
	{
		return this.WordControl.onKeyDown(e);
	};
	VisioEditorApi.prototype.executeShortcut = function(type)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return false;

		return logicDocument.executeShortcut(type);
	};

	window["VisioEditorApi"]                                 = VisioEditorApi;
	window["VisioEditorApi"].prototype["asc_nativeOpenFile"] = function(base64File, version)
	{
		// this.SpellCheckUrl = '';

		this.User = new AscCommon.asc_CUser();
		this.User.setId("TM");
		this.User.setUserName("native");

		this.InitEditor();

		AscCommon.g_oIdCounter.Set_Load(true);

		this.InitEditor();

		this.isOpenOOXInBrowser = this["asc_isSupportFeature"]("ooxml") && AscCommon.checkOOXMLSignature(base64File);
		if (this.isOpenOOXInBrowser) {
			//slice because array contains garbage after end of function
			this.openOOXInBrowserZip = base64File.slice();
			this.OpenDocumentFromZipNoInit(base64File);
		}

		this.LoadedObject = 1;
		AscCommon.g_oIdCounter.Set_Load(false);
	};

	window["VisioEditorApi"].prototype["asc_nativeCalculateFile"] = function()
	{
		//todo
	};

	window["VisioEditorApi"].prototype["asc_nativeApplyChanges"] = function(changes)
	{
		//todo
	};

	window["VisioEditorApi"].prototype["asc_nativeApplyChanges2"] = function(data, isFull)
	{
		//todo
	};

	window["VisioEditorApi"].prototype["asc_nativeGetFile"] = function()
	{
		//todo
	};

	window["VisioEditorApi"].prototype.asc_nativeGetFile3 = function()
	{
		//todo
	};

	window["VisioEditorApi"].prototype["asc_nativeGetFileData"] = function()
	{
		if (this.isOpenOOXInBrowser && this.saveDocumentToZip) {
			let res;
			this.saveDocumentToZip(this.Document, this.editorId, function (data) {
				res = data;
			});
			if (res) {
				window["native"] && window["native"]["Save_End"] && window["native"]["Save_End"](";v10;", res.length);
				return res;
			}
			return new Uint8Array(0);
		}
	};

	window["VisioEditorApi"].prototype["asc_nativeCalculate"] = function()
	{
	};

	window["VisioEditorApi"].prototype["asc_nativePrint"] = function(_printer, _page, _options)
	{
		if (undefined === _printer && _page === undefined)
		{
			if (undefined !== window["AscDesktopEditor"])
			{
				var isSelection = (_options && _options["printOptions"] && _options["printOptions"]["selection"]) ? true : false;
				var _drawing_document = this.WordControl.m_oDrawingDocument;
				var pagescount        = _drawing_document.GetSlidesCount();
				if (isSelection)
					pagescount = this.WordControl.Thumbnails.GetSelectedArray().length;

				window["AscDesktopEditor"]["Print_Start"](this.DocumentUrl, pagescount, "", this.Document.getCurrentPage());

				var oDocRenderer                         = new AscCommon.CDocumentRenderer();
				oDocRenderer.InitPicker(AscCommon.g_oTextMeasurer.m_oManager);
				oDocRenderer.VectorMemoryForPrint        = new AscCommon.CMemory();
				var bOldShowMarks                        = this.ShowParaMarks;
				this.ShowParaMarks                       = false;
				oDocRenderer.IsNoDrawingEmptyPlaceholder = true;

				pagescount = _drawing_document.GetSlidesCount();
				for (var i = 0; i < pagescount; i++)
				{
					if (isSelection && !this.WordControl.Thumbnails.isSelectedPage(i))
						continue;

					oDocRenderer.Memory.Seek(0);
					oDocRenderer.VectorMemoryForPrint.ClearNoAttack();

					oDocRenderer.BeginPage(_drawing_document.m_oLogicDocument.GetWidthMM(), _drawing_document.m_oLogicDocument.GetHeightMM());
					this.WordControl.m_oLogicDocument.DrawPage(i, oDocRenderer);
					oDocRenderer.EndPage();

					window["AscDesktopEditor"]["Print_Page"](oDocRenderer.Memory.GetBase64Memory(), _drawing_document.m_oLogicDocument.GetWidthMM(), _drawing_document.m_oLogicDocument.GetHeightMM());
				}

				if (0 == pagescount)
				{
					oDocRenderer.BeginPage(_drawing_document.m_oLogicDocument.GetWidthMM(), _drawing_document.m_oLogicDocument.GetHeightMM());
					oDocRenderer.EndPage();

					window["AscDesktopEditor"]["Print_Page"](oDocRenderer.Memory.GetBase64Memory());
				}

				this.ShowParaMarks = bOldShowMarks;

				window["AscDesktopEditor"]["Print_End"]();
			}
			return;
		}

		let _logic_doc = this.WordControl.m_oLogicDocument;
		let pageW = _logic_doc.GetWidthMM();
		let pageH = _logic_doc.GetHeightMM();

		_printer.BeginPage(pageW, pageH);
		_logic_doc.DrawPage(_page, _printer);
		if (this.watermarkDraw)
			this.watermarkDraw.DrawOnRenderer(_printer, pageW, pageH);
		_printer.EndPage();
	};

	window["VisioEditorApi"].prototype["asc_nativePrintPagesCount"] = function()
	{
		return this.WordControl.GetSlidesCount();
	};

	window["VisioEditorApi"].prototype["asc_nativeGetPDF"] = function(options)
	{
		if (options && options["watermark"])
		{
			this.watermarkDraw = new AscCommon.CWatermarkOnDraw(options["watermark"], this);
			this.watermarkDraw.generateNative();
		}
		else
		{
			this.watermarkDraw = null;
		}

		var pagescount = this["asc_nativePrintPagesCount"]();
		if (options && options["printOptions"] && options["printOptions"]["onlyFirstPage"])
			pagescount = 1;

		var _renderer                         = new AscCommon.CDocumentRenderer();
		_renderer.InitPicker(AscCommon.g_oTextMeasurer.m_oManager);
		_renderer.VectorMemoryForPrint        = new AscCommon.CMemory();
		_renderer.DocInfo(this.asc_getCoreProps());
		var _bOldShowMarks                    = this.ShowParaMarks;
		this.ShowParaMarks                    = false;
		_renderer.IsNoDrawingEmptyPlaceholder = true;

		let nativeOptions = options ? options["nativeOptions"] : undefined;
		let pages = nativeOptions ? AscCommon.getNativePrintRanges(nativeOptions["pages"], nativeOptions["currentPage"], pagescount) : undefined;

		for (var i = 0; i < pagescount; i++)
		{
			if (pages !== undefined && !pages[i])
				continue;
			this["asc_nativePrint"](_renderer, i, options);
		}

		this.ShowParaMarks = _bOldShowMarks;

		window["native"]["Save_End"]("", _renderer.Memory.GetCurPosition());

		return _renderer.Memory.data;
	};
	VisioEditorApi.prototype.openDocument = function(file)
	{
		let perfStart = performance.now();
		// if (file.changes && this.VersionHistory)
		// {
		// 	this.VersionHistory.changes = file.changes;
		// 	this.VersionHistory.applyChanges(this);
		// }
		this.isOpenOOXInBrowser = this["asc_isSupportFeature"]("ooxml") && AscCommon.checkOOXMLSignature(file.data);
		if (this.isOpenOOXInBrowser) {
			this.openOOXInBrowserZip = file.data;
			this.OpenDocumentFromZip(file.data);
		} else {
			this.sendEvent("asc_onError", Asc.c_oAscError.ID.AccessDeny, Asc.c_oAscError.Level.Critical);
		}
		let perfEnd = performance.now();
		AscCommon.sendClientLog("debug", AscCommon.getClientInfoString("onOpenDocument", perfEnd - perfStart), this);
	};
	VisioEditorApi.prototype.isDocumentModified = function()
	{
		return false;
		if (!this.canSave)
		{
			// Пока идет сохранение, мы не закрываем документ
			return true;
		}
		return this.isDocumentModify;
	};
	VisioEditorApi.prototype.SetDrawingFreeze = function(bIsFreeze)
	{
		if (!this.isLoadFullApi)
		{
			this.tmpIsFreeze = bIsFreeze;
			return;
		}

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
	VisioEditorApi.prototype.asc_setSpellCheck = function(isOn)
	{
	};
	VisioEditorApi.prototype.asc_setSpellCheckSettings = function(oSettings)
	{
	};
	VisioEditorApi.prototype.zoomIn         = function()
	{
		this.WordControl.zoom_In();
	};
	VisioEditorApi.prototype.zoomOut        = function()
	{
		this.WordControl.zoom_Out();
	};
	VisioEditorApi.prototype.zoomFitToPage  = function()
	{
		if (!this.isLoadFullApi)
		{
			this.tmpZoomType = AscCommon.c_oZoomType.FitToPage;
			return;
		}
		this.WordControl.zoom_FitToPage();
	};
	VisioEditorApi.prototype.zoomFitToWidth = function()
	{
		if (!this.isLoadFullApi)
		{
			this.tmpZoomType = AscCommon.c_oZoomType.FitToWidth;
			return;
		}
		this.WordControl.zoom_FitToWidth();
	};
	VisioEditorApi.prototype.zoomCustomMode = function()
	{
		if (!this.isLoadFullApi)
		{
			this.tmpZoomType = AscCommon.c_oZoomType.CustomMode;
			return;
		}
		this.WordControl.m_nZoomType = 0;
		this.WordControl.zoom_Fire();
	};
	VisioEditorApi.prototype.zoom100        = function()
	{
		this.WordControl.m_nZoomValue = 100;
		this.WordControl.zoom_Fire();
	};
	VisioEditorApi.prototype.zoom           = function(percent)
	{
		this.WordControl.m_nZoomValue = percent;
		this.WordControl.zoom_Fire(0);
	};
	VisioEditorApi.prototype.goToPage       = function(number)
	{
		this.WordControl.GoToPage(number);
	};
	VisioEditorApi.prototype.SetFontRenderingMode         = function(mode)
	{
		if (!this.isLoadFullApi)
		{
			this.tmpFontRenderingMode = mode;
			return;
		}

		if (c_oAscFontRenderingModeType.noHinting === mode)
			AscCommon.g_fontManager.SetHintsProps(false, false);
		else if (c_oAscFontRenderingModeType.hinting === mode)
			AscCommon.g_fontManager.SetHintsProps(true, false);
		else if (c_oAscFontRenderingModeType.hintingAndSubpixeling === mode)
			AscCommon.g_fontManager.SetHintsProps(true, true);

		if (AscCommon.g_fontManager2 !== undefined && AscCommon.g_fontManager2 !== null)
			AscCommon.g_fontManager2.ClearFontsRasterCache();

		// this.WordControl.m_oDrawingDocument.ClearCachePages();
		//
		// if (this.bInit_word_control)
		// 	this.WordControl.OnScroll();
	}
	VisioEditorApi.prototype.asc_setLocale = function(val)
	{
		this.locale = val;
	};
	VisioEditorApi.prototype.asc_getLocale = function()
	{
		return this.locale;
	};
	VisioEditorApi.prototype.asc_SetDocumentUnits = function(_units)
	{
		//todo
		this.tmpDocumentUnits = _units;
	};
	VisioEditorApi.prototype.updateSkin = function()
	{
		var baseElem = document.getElementById(this.HtmlElementName);
		if (baseElem)
			baseElem.style.backgroundColor = AscCommon.GlobalSkin.BackgroundColor;

		var obj_id_main = document.getElementById("id_main");
		if (obj_id_main)
		{
			obj_id_main.style.backgroundColor = AscCommon.GlobalSkin.BackgroundColor;
			document.getElementById("id_panel_thumbnails").style.backgroundColor = AscCommon.GlobalSkin.BackgroundColorThumbnails;
			document.getElementById("id_panel_thumbnails_split").style.backgroundColor = AscCommon.GlobalSkin.BackgroundColorThumbnails;
			document.getElementById("id_main_parent").style.borderLeftColor = AscCommon.GlobalSkin.BorderSplitterColor;
			document.getElementById("id_viewer").style.backgroundColor = AscCommon.GlobalSkin.BackgroundColor;
			document.getElementById("id_panel_right").style.backgroundColor = AscCommon.GlobalSkin.ScrollBackgroundColor;
			document.getElementById("id_horscrollpanel").style.backgroundColor = AscCommon.GlobalSkin.ScrollBackgroundColor;
		}

		if (this.WordControl && this.WordControl.m_oBody)
			this.WordControl.OnResize(true);
	};
	VisioEditorApi.prototype.Resize = function() {
		if (false === this.bInit_word_control)
			return;
		this.WordControl.OnResize(false);
	};
	VisioEditorApi.prototype.sendEvent = function()
	{
		this.sendInternalEvent.apply(this, arguments);
		var name = arguments[0];
		if (_callbacks.hasOwnProperty(name))
		{
			for (var i = 0; i < _callbacks[name].length; ++i)
			{
				_callbacks[name][i].apply(this || window, Array.prototype.slice.call(arguments, 1));
			}
			return true;
		}
		return false;
	};
	var _callbacks = {};
	VisioEditorApi.prototype.asc_registerCallback = function(name, callback)
	{
		if (!_callbacks.hasOwnProperty(name))
			_callbacks[name] = [];
		_callbacks[name].push(callback);
	};
	VisioEditorApi.prototype.asc_unregisterCallback = function(name, callback)
	{
		if (_callbacks.hasOwnProperty(name))
		{
			for (var i = _callbacks[name].length - 1; i >= 0; --i)
			{
				if (_callbacks[name][i] == callback)
					_callbacks[name].splice(i, 1);
			}
		}
	};
	VisioEditorApi.prototype.asc_checkNeedCallback = function(name)
	{
		if (_callbacks.hasOwnProperty(name))
		{
			return true;
		}
		return false;
	};
	VisioEditorApi.prototype.asc_SetFastCollaborative = function(isOn)
	{
	};
	VisioEditorApi.prototype.getCountPages  = function()
	{
		return this.WordControl && this.WordControl.m_oLogicDocument && this.WordControl.m_oLogicDocument.getCountPages() || 0
	};
	VisioEditorApi.prototype.GetCurrentVisiblePage  = function()
	{
		return this.WordControl.m_oDrawingDocument.SlideCurrent;
	};
	VisioEditorApi.prototype.ShowThumbnails           = function(bIsShow)
	{
		if (bIsShow)
		{
			this.WordControl.Splitter1Pos = this.WordControl.OldSplitter1Pos;
			if (this.WordControl.Splitter1Pos == 0)
				this.WordControl.Splitter1Pos = 70;
			this.WordControl.OnResizeSplitter();
		}
		else
		{
			var old                       = this.WordControl.OldSplitter1Pos;
			this.WordControl.Splitter1Pos = 0;
			this.WordControl.OnResizeSplitter();
			this.WordControl.OldSplitter1Pos = old;
		}
	}
	VisioEditorApi.prototype["asc_setViewerTargetType"] = VisioEditorApi.prototype.asc_setViewerTargetType = function(type) {
		this.isHandMode = ("hand" === type);
		// this.WordControl.checkMouseHandMode();
		// this.WordControl.onMouseMove();
		this.sendEvent("asc_onChangeViewerTargetType", this.isHandMode);
	};
	VisioEditorApi.prototype.getLogicDocument = VisioEditorApi.prototype.private_GetLogicDocument = function() {
		return this.WordControl && this.WordControl.m_oLogicDocument || null;
	};

	VisioEditorApi.prototype.asc_DownloadAs = function(options)
	{
		if (this.isLongAction()) {
			return;
		}
		this.downloadAs(Asc.c_oAscAsyncAction.DownloadAs, options);
	};
	VisioEditorApi.prototype._downloadAs = function(actionType, options, oAdditionalData, dataContainer, downloadType)
	{
		var t = this;
		var fileType = options.fileType;

		if (this.isCloudSaveAsLocalToDrawingFormat(actionType, fileType))
		{
			this.localSaveToDrawingFormat(this.WordControl.m_oDrawingDocument.ToRendererPart(false, options.isPdfPrint), fileType);
			return true;
		}

		if (c_oAscFileType.PDF === fileType || c_oAscFileType.PDFA === fileType)
		{
			var isSelection = false;
			if (options.advancedOptions && options.advancedOptions && (Asc.c_oAscPrintType.Selection === options.advancedOptions.asc_getPrintType()))
				isSelection = true;

			var dd             = this.WordControl.m_oDrawingDocument;
			dataContainer.data = dd.ToRendererPart(oAdditionalData["nobase64"], isSelection);
		}
		else if(false && this.isOpenOOXInBrowser && this["asc_isSupportFeature"]("ooxml"))
		{
			var title = this.documentTitle;
			this.saveLogicDocumentToZip(undefined, undefined,
				function(data) {
					if (data) {
						if (c_oAscFileType.VSDX === fileType && !window.isCloudCryptoDownloadAs) {
							AscCommon.DownloadFileFromBytes(data, title, AscCommon.openXml.GetMimeType("vsdx"));
						} else {
							dataContainer.data = data;
							if (window.isCloudCryptoDownloadAs)
							{
								window["AscDesktopEditor"]["CryptoDownloadAs"](dataContainer.data, fileType);
								return true;
							}
							t._downloadAsUsingServer(actionType, options, oAdditionalData, dataContainer, downloadType);
							return;
						}
					} else {
						t.sendEvent("asc_onError", Asc.c_oAscError.ID.Unknown, Asc.c_oAscError.Level.NoCritical);
					}
					if (actionType)
					{
						t.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, actionType);
					}
				});
			return true;
		}

		if (window.isCloudCryptoDownloadAs)
		{
			window["AscDesktopEditor"]["CryptoDownloadAs"](dataContainer.data, fileType);
			return true;
		}
	};
	VisioEditorApi.prototype.asc_getPageName = function(index)
	{
		if (this.Document && this.Document.pages && index < this.Document.pages.page.length) {
			return this.Document.pages.page[index].name;
		}
		return "";
	};
	/*callbacks*/
	VisioEditorApi.prototype.sync_zoomChangeCallback  = function(percent, type)
	{	//c_oAscZoomType.Current, c_oAscZoomType.FitWidth, c_oAscZoomType.FitPage
		this.sendEvent("asc_onZoomChange", percent, type);
	};
	VisioEditorApi.prototype.sync_countPagesCallback  = function(count)
	{
		this.sendEvent("asc_onCountPages", count);
	};
	VisioEditorApi.prototype.sync_currentPageCallback = function(number)
	{
		this.sendEvent("asc_onCurrentPage", number);
	};
	VisioEditorApi.prototype.sync_ContextMenuCallback = function(Data)
	{
		this.sendEvent("asc_onContextMenu", Data);
	};
	VisioEditorApi.prototype.sync_EndAddShape = function()
	{
		editor.sendEvent("asc_onEndAddShape");
		if (this.WordControl.m_oDrawingDocument.m_sLockedCursorType == "crosshair")
		{
			this.WordControl.m_oDrawingDocument.UnlockCursorType();
		}
		if(this.WordControl.m_oLogicDocument.TurnOffInterfaceEvents) {
			this.WordControl.m_oLogicDocument.TurnOn_InterfaceEvents(false);
		}
	};
	VisioEditorApi.prototype.syncOnThumbnailsShow = function()
	{
		var bIsShow = true;
		if (0 == this.WordControl.Splitter1Pos)
			bIsShow = false;

		this.sendEvent("asc_onThumbnailsShow", bIsShow);
	};
	VisioEditorApi.prototype.OnMouseUp = function(x, y)
	{
		var _e = AscCommon.CreateMouseUpEventObject(x, y);
		AscCommon.Window_OnMouseUp(_e);

		//this.WordControl.onMouseUpExternal(x, y);
	};
	//temp stubs
	VisioEditorApi.prototype.getCountSlides = function()
	{
		return this.Document.getCountPages();
	};

	VisioEditorApi.prototype._printDesktop = function (options)
	{
		let desktopOptions = {};
		if (options && options.advancedOptions)
			desktopOptions["nativeOptions"] = options.advancedOptions.asc_getNativeOptions();

		window["AscDesktopEditor"]["Print"](JSON.stringify(desktopOptions));
		return true;
	};
	//-------------------------------------------------------------export---------------------------------------------------
	window['Asc']                                                       = window['Asc'] || {};
	window['Asc']['VisioEditorApi']                                       = VisioEditorApi;
	let prot															= VisioEditorApi.prototype;
	//todo apiBase
	prot['asc_GetFontThumbnailsPath']				= prot.asc_GetFontThumbnailsPath;
	prot["asc_coAuthoringDisconnect"]				= prot.asc_coAuthoringDisconnect;
	prot["asc_coAuthoringChatSendMessage"]			= prot.asc_coAuthoringChatSendMessage;
	prot["asc_coAuthoringChatGetMessages"]			= prot.asc_coAuthoringChatGetMessages;
	prot["asc_coAuthoringGetUsers"]					= prot.asc_coAuthoringGetUsers;
	prot["asc_isDocumentCanSave"]					= prot.asc_isDocumentCanSave;
	prot["asc_getCanUndo"]							= prot.asc_getCanUndo;
	prot["asc_getCanRedo"]							= prot.asc_getCanRedo;
	prot["asc_getEditorPermissions"]				= prot.asc_getEditorPermissions;
	prot['asc_setDocInfo']							= prot.asc_setDocInfo;
	prot['asc_changeDocInfo']                   	= prot.asc_changeDocInfo;
	prot["asc_LoadDocument"]						= prot.asc_LoadDocument;
	prot["asc_getDocumentName"] 					= prot.asc_getDocumentName;
	prot["asc_getCoreProps"]						= prot.asc_getCoreProps;
	prot["asc_Save"]								= prot.asc_Save;
	prot["forceSave"]								= prot.forceSave;
	prot["asc_setIsForceSaveOnUserSave"]			= prot.asc_setIsForceSaveOnUserSave;
	prot["asc_setAutoSaveGap"]						= prot.asc_setAutoSaveGap;
	prot["asc_decodeBuffer"]						= prot.asc_decodeBuffer;
	prot["asc_isOffline"]							= prot.asc_isOffline;
	prot["asc_getUrlType"]							= prot.asc_getUrlType;
	prot["asc_prepareUrl"]							= prot.asc_prepareUrl;
	prot["asc_getSessionToken"]						= prot.asc_getSessionToken;
	prot["asc_setInterfaceDrawImagePlaceShape"]		= prot.asc_setInterfaceDrawImagePlaceShape;
	prot["asc_nativeInitBuilder"]					= prot.asc_nativeInitBuilder;
	prot["asc_pluginsRegister"]						= prot.asc_pluginsRegister;
	prot["asc_pluginRun"]							= prot.asc_pluginRun;
	prot["asc_pluginStop"]							= prot.asc_pluginStop;
	prot["asc_pluginResize"]						= prot.asc_pluginResize;
	prot["asc_pluginButtonClick"]					= prot.asc_pluginButtonClick;
	prot["asc_pluginEnableMouseEvents"]				= prot.asc_pluginEnableMouseEvents;
	prot['asc_setViewMode']                     	= prot.asc_setViewMode;
	prot['asc_enableKeyEvents']                     = prot.asc_enableKeyEvents;

	// signatures
	prot["asc_addSignatureLine"] 					= prot.asc_addSignatureLine;
	prot["asc_CallSignatureDblClickEvent"] 			= prot.asc_CallSignatureDblClickEvent;
	prot["asc_getRequestSignatures"] 				= prot.asc_getRequestSignatures;
	prot["asc_AddSignatureLine2"]             		= prot.asc_AddSignatureLine2;
	prot["asc_Sign"]             					= prot.asc_Sign;
	prot["asc_RequestSign"]             			= prot.asc_RequestSign;
	prot["asc_ViewCertificate"] 					= prot.asc_ViewCertificate;
	prot["asc_SelectCertificate"] 					= prot.asc_SelectCertificate;
	prot["asc_GetDefaultCertificate"] 				= prot.asc_GetDefaultCertificate;
	prot["asc_getSignatures"] 						= prot.asc_getSignatures;
	prot["asc_isSignaturesSupport"] 				= prot.asc_isSignaturesSupport;
	prot["asc_isProtectionSupport"] 				= prot.asc_isProtectionSupport;
	prot["asc_isAnonymousSupport"] 					= prot.asc_isAnonymousSupport;
	prot["asc_RemoveSignature"] 					= prot.asc_RemoveSignature;
	prot["asc_RemoveAllSignatures"] 				= prot.asc_RemoveAllSignatures;
	prot["asc_gotoSignature"] 						= prot.asc_gotoSignature;
	prot["asc_getSignatureSetup"] 					= prot.asc_getSignatureSetup


	prot['asc_CloseFile']             				= prot.asc_CloseFile;
	prot['asc_getAppProps']             			= prot.asc_getAppProps;
	prot['asc_CheckCopy']             				= prot.asc_CheckCopy;
	prot['asc_PasteData']             				= prot.asc_PasteData;
	prot['asc_setSpellCheck']             			= prot.asc_setSpellCheck;
	prot['asc_setSpellCheckSettings']             	= prot.asc_setSpellCheckSettings;
	prot['asc_setLocale']             				= prot.asc_setLocale;
	prot['asc_getLocale']             				= prot.asc_getLocale;
	prot['asc_SetDocumentUnits']             		= prot.asc_SetDocumentUnits;
	prot['asc_registerCallback']             		= prot.asc_registerCallback;
	prot['asc_unregisterCallback']             		= prot.asc_unregisterCallback;
	prot['asc_checkNeedCallback']             		= prot.asc_checkNeedCallback;
	prot['asc_SetFastCollaborative']             	= prot.asc_SetFastCollaborative;
	prot['asc_DownloadAs']             				= prot.asc_DownloadAs;
	prot['asc_getPageName']             			= prot.asc_getPageName;
	prot['InitEditor']                          	= prot.InitEditor;
	prot['isDocumentModified']             			= prot.isDocumentModified;
	prot['SetDrawingFreeze']             			= prot.SetDrawingFreeze;
	prot['zoomIn']             						= prot.zoomIn;
	prot['zoomOut']             					= prot.zoomOut;
	prot['zoomFitToPage']             				= prot.zoomFitToPage;
	prot['zoomFitToWidth']             				= prot.zoomFitToWidth;
	prot['zoomCustomMode']             				= prot.zoomCustomMode;
	prot['zoom100']             					= prot.zoom100;
	prot['zoom']             						= prot.zoom;
	prot['goToPage']             					= prot.goToPage;
	prot['SetFontRenderingMode']             		= prot.SetFontRenderingMode;
	prot['Resize']             						= prot.Resize;
	prot['sendEvent']             					= prot.sendEvent;
	prot['getCountPages']             				= prot.getCountPages;
	prot['GetCurrentVisiblePage']             		= prot.GetCurrentVisiblePage;
	prot['ShowThumbnails']             				= prot.ShowThumbnails;
	prot['OnMouseUp']             					= prot.OnMouseUp;

})(window, window.document);
