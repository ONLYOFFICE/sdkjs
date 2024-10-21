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
	/**
	 *
	 * @param config
	 * @constructor
	 * @extends {AscCommon.baseEditorsApi}
	 */
	function asc_docs_api(config)
	{
		AscCommon.baseEditorsApi.call(this, config, AscCommon.c_oEditorId.Draw);

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

		this._init();
		return this;
	}

	asc_docs_api.prototype = Object.create(AscCommon.baseEditorsApi.prototype);
	asc_docs_api.prototype.constructor = asc_docs_api;

	asc_docs_api.prototype.InitEditor = function(){
		this.Document = new AscCommonDraw.CVisioDocument(this, this.WordControl.m_oDrawingDocument);

		this.WordControl.m_oLogicDocument = this.Document;
		this.WordControl.m_oDrawingDocument.m_oLogicDocument = this.WordControl.m_oLogicDocument;
	};
	asc_docs_api.prototype._onEndLoadSdk  = function()
	{
		AscCommon.baseEditorsApi.prototype._onEndLoadSdk.call(this);

		AscFonts.g_fontApplication.Init();

		this.FontLoader  = AscCommon.g_font_loader;
		this.ImageLoader = AscCommon.g_image_loader;
		this.FontLoader.put_Api(this);
		this.ImageLoader.put_Api(this);

		this._loadSdkImages();

		this.CreateComponents();
	};
	asc_docs_api.prototype.CreateComponents = function()
	{
		window.editor = this;
		//for CShapeDrawer.CheckDash
		window.Asc.editor = this;

		//stubs for html page
		window.CLayoutThumbnailDrawer = function () {};
		window.CMasterThumbnailDrawer = function () {};
		this.WordControl  = new AscCommonSlide.CEditorPage(this);
		if (this.HtmlElement != null) {
			let thumbsHtmlString = 	"<div id=\"id_panel_thumbnails\" class=\"block_elem\" style=\"touch-action: none; background-color: rgb(64, 64, 64); display: block; left: 0px; top: 0px; width: 20%; height:100%; overflow: scroll;\">" +
											"</div>";
			this.HtmlElement.innerHTML = ("<div id=\"id_main\" class=\"block_elem\" style=\"width:80%;height:100%;touch-action:none;-ms-touch-action: none;-moz-user-select:none;-khtml-user-select:none;user-select:none;background-color:" + AscCommon.GlobalSkin.BackgroundColor + ";overflow:hidden;\" UNSELECTABLE=\"on\">\
											<div id=\"id_main_view\" class=\"block_elem\" style=\"width:100%;height:100%;touch-action:none;overflow:hidden\">\
												<canvas id=\"id_viewer\" class=\"block_elem\" style=\"width:100%;height:100%;touch-action:none;-ms-touch-action: none;-webkit-user-select: none; background-color:" + AscCommon.GlobalSkin.BackgroundColor + ";z-index:1\"></canvas>\
												<canvas id=\"id_viewer_overlay\" class=\"block_elem\" style=\"touch-action:none;-ms-touch-action: none;-webkit-user-select: none; z-index:2\"></canvas>\
												<div id=\"id_target_cursor\" class=\"block_elem\" width=\"1\" height=\"1\" style=\"touch-action:none;-ms-touch-action: none;-webkit-user-select: none;width:2px;height:13px;z-index:4;\"></div>\
											</div>" + this.HtmlElement.innerHTML);
			this.HtmlElement.innerHTML = thumbsHtmlString + this.HtmlElement.innerHTML;
		}


		this.canvas = document.getElementById("id_viewer");
	};
	// работа с шрифтами
	asc_docs_api.prototype.asyncFontsDocumentStartLoaded = function(blockType)
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
				if (this.DocInfo && this.DocInfo.get_OfflineApp())
				{
					var localUrl = _loader_object.ImageMap[i];
					g_oDocumentUrls.addImageUrl(localUrl, this.documentUrl + 'media/' + localUrl);
				}
				++_count;
			}
		}

		_progress.ImagesCount  = _count;
		_progress.CurrentImage = 0;
	};
	asc_docs_api.prototype.asyncFontsDocumentEndLoaded   = function(blockType) {
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
	asc_docs_api.prototype.asyncImagesDocumentEndLoaded = function()
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
	asc_docs_api.prototype._openDocumentEndCallback = function()
	{
		if (this.isDocumentLoadComplete || !this.ServerImagesWaitComplete || !this.ServerIdWaitComplete || !this.WordControl || !this.WordControl.m_oLogicDocument)
			return;

		if (this.isViewMode)
			this.asc_setViewMode(true);

		this.onDocumentContentReady();

		// Меняем тип состояния (на никакое)
		this.advancedOptionsAction = AscCommon.c_oAscAdvancedOptionsAction.None;

		this.Document.draw(this.Document.zoom, undefined, this.Document.pageIndex);
	};

	asc_docs_api.prototype.OpenDocumentFromZip = function(data)
	{
		return this.OpenDocumentFromZipNoInit(data);
	};
	asc_docs_api.prototype.OpenDocumentFromZipNoInit = function(data)
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
		this.Document.imageMap = xmlParserContext.loadDataLinks();
		this.ServerIdWaitComplete = true;

		jsZlib.close();
		return true;
	};
	asc_docs_api.prototype.asc_CloseFile            = function()
	{
		AscCommon.History.Clear();
		AscCommon.g_oTableId.Clear();
		AscCommon.g_oIdCounter.Clear();
		this.isApplyChangesOnOpenEnabled = true;
		this.isDocumentLoadComplete = false;
		this.turnOffSpecialModes();
		AscCommon.pptx_content_loader.ImageMapChecker = {};
	};
	// Callbacks
	/* все имена callback'оф начинаются с On. Пока сделаны:
	 */
	var _callbacks = {};

	asc_docs_api.prototype.asc_registerCallback = function(name, callback)
	{
		if (!_callbacks.hasOwnProperty(name))
			_callbacks[name] = [];
		_callbacks[name].push(callback);
	};
	asc_docs_api.prototype.asc_nativeOpenFile = function(base64File, version)
	{
		// this.SpellCheckUrl = '';
		//
		// this.User = new AscCommon.asc_CUser();
		// this.User.setId("TM");
		// this.User.setUserName("native");

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

	asc_docs_api.prototype.asc_nativeApplyChanges2 = function(data, isFull)
	{
		// Чтобы заново созданные параграфы не отображались залоченными
		AscCommon.g_oIdCounter.Set_Load(true);
		//todo

		AscCommon.g_oIdCounter.Set_Load(false);
	};

	asc_docs_api.prototype.asc_nativeGetFileData = function()
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

	asc_docs_api.prototype.asc_nativeCalculateFile = function(options)
	{
	};
	asc_docs_api.prototype.asc_nativePrintPagesCount = function()
	{
		return 1;
	};
	asc_docs_api.prototype.asc_nativeGetPDF = function(options)
	{
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

		//todo ShowParaMarks
		this.ShowParaMarks = _bOldShowMarks;

		window["native"]["Save_End"]("", _renderer.Memory.GetCurPosition());

		return _renderer.Memory.data;
	};
	asc_docs_api.prototype.asc_nativePrint = function(_printer, _page, _options)
	{
		if (undefined === _printer && _page === undefined)
		{
			if (undefined !== window["AscDesktopEditor"]) {
				var isSelection = (_options && _options["printOptions"] && _options["printOptions"]["selection"]) ? true : false;
				//todo isSelection
				let pagescount = 1;
				var _logic_doc = this.Document;

				window["AscDesktopEditor"]["Print_Start"](this.DocumentUrl, pagescount, "", 0);

				var oDocRenderer = new AscCommon.CDocumentRenderer();
				oDocRenderer.InitPicker(AscCommon.g_oTextMeasurer.m_oManager);
				oDocRenderer.VectorMemoryForPrint = new AscCommon.CMemory();
				var bOldShowMarks = this.ShowParaMarks;
				this.ShowParaMarks = false;
				oDocRenderer.IsNoDrawingEmptyPlaceholder = true;

				for (var i = 0; i < pagescount; i++) {
					oDocRenderer.Memory.Seek(0);
					oDocRenderer.VectorMemoryForPrint.ClearNoAttack();

					oDocRenderer.BeginPage(_logic_doc.GetWidthMM(i), _logic_doc.GetHeightMM(i));
					_logic_doc.draw(100, oDocRenderer);
					oDocRenderer.EndPage();

					window["AscDesktopEditor"]["Print_Page"](oDocRenderer.Memory.GetBase64Memory(), _logic_doc.GetWidthMM(i), _logic_doc.GetHeightMM(i));
				}
				if (0 === pagescount) {
					oDocRenderer.BeginPage(_logic_doc.GetWidthMM(i), _logic_doc.GetHeightMM(i));
					oDocRenderer.EndPage();

					window["AscDesktopEditor"]["Print_Page"](oDocRenderer.Memory.GetBase64Memory());
				}

				this.ShowParaMarks = bOldShowMarks;

				window["AscDesktopEditor"]["Print_End"]();
			}
		} else {
			let _logic_doc = this.Document;
			_printer.BeginPage(_logic_doc.GetWidthMM(_page), _logic_doc.GetHeightMM(_page));
			_logic_doc.draw(100, _printer);
			_printer.EndPage();
		}
	};

	//-------------------------------------------------------------export---------------------------------------------------
	window['Asc']                                                       = window['Asc'] || {};
	window['Asc']['asc_docs_api']                                       = asc_docs_api;
	asc_docs_api.prototype['OpenDocumentFromZip']             			= asc_docs_api.prototype.OpenDocumentFromZip;
	asc_docs_api.prototype['asc_nativeOpenFile']             			= asc_docs_api.prototype.asc_nativeOpenFile;
	asc_docs_api.prototype['asc_nativeApplyChanges2']             		= asc_docs_api.prototype.asc_nativeApplyChanges2;
	asc_docs_api.prototype['asc_nativeGetFileData']             		= asc_docs_api.prototype.asc_nativeGetFileData;
	asc_docs_api.prototype['asc_nativeCalculateFile']             		= asc_docs_api.prototype.asc_nativeCalculateFile;
	asc_docs_api.prototype['asc_nativePrintPagesCount']             	= asc_docs_api.prototype.asc_nativePrintPagesCount;
	asc_docs_api.prototype['asc_nativeGetPDF']             				= asc_docs_api.prototype.asc_nativeGetPDF;
	asc_docs_api.prototype['asc_nativePrint']             				= asc_docs_api.prototype.asc_nativePrint;

})(window, window.document);
