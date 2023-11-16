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

(function(){

	function CCacheManager()
	{
		this.images = [];

		this.lock = function(w, h)
		{
			for (let i = 0; i < this.images.length; i++)
			{
				if (this.images[i].locked)
					continue;

				let canvas = this.images[i].canvas;
				let testW = canvas.width;
				let testH = canvas.height;
				if (w > testW || h > testH || ((4 * w * h) < (testW * testH)))
				{
					this.images.splice(i, 1);
					continue;	
				}

				this.images[i].locked = true;
				return canvas;
			}

			let newImage = { canvas : document.createElement("canvas"), locked : true };
			newImage.canvas.width = w + 100;
			newImage.canvas.height = h + 100;
			this.images.push(newImage);
			return newImage.canvas;
		};

		this.unlock = function(canvas)
		{
			for (let i = 0, len = this.images.length; i < len; i++)
			{
				if (this.images[i].canvas === canvas)
				{
					this.images[i].locked = false;
					return;
				}
			}
		};

		this.clear = function()
		{
			this.images = [];
		};
	};

	// wasm/asmjs module state
	var ModuleState = {
		None : 0,
		Loading : 1,
		Loaded : 2
	};

	// zoom mode
	var ZoomMode = {
		Custom : 0,
		Width : 1,
		Page : 2
	};

	// класс страницы.
	// isPainted - значит она когда-либо рисовалась и дорисовалась до конца (шрифты загружены)
	// links - гиперссылки. они запрашиваются ТОЛЬКО у страниц на экране и у отрисованных страниц.
	// так как нет смысла запрашивать ссылки у невидимых страниц и у страниц, которые мы в данный момент не можем отрисовать
	// text - текстовые команды. они запрашиваются всегда, если есть какая-то страница без текстовых команд
	// страницы на экране в приоритете.
	function CPageInfo()
	{
		this.isPainted = false;
		this.links = null;
		this.fields = null;
		this.annots = null;
	}
	
	function CDocumentPagesInfo()
	{
		this.pages = [];

		// все страницы ДО this.countCurrentPage должны иметь текстовые команды
		this.countTextPages = 0;
	}
	CDocumentPagesInfo.prototype.setCount = function(count)
	{
		this.pages = new Array(count);
		for (var i = 0; i < count; i++)
		{
			this.pages[i] = new CPageInfo();
		}
		this.countTextPages = 0;
	};
	CDocumentPagesInfo.prototype.setPainted = function(index)
	{
		this.pages[index].isPainted = true;
	};

	function CHtmlPage(id, api)
	{
		this.Api = api;
		this.parent = document.getElementById(id);
		this.thumbnails = null;
		
		this.bCachedMarkupAnnnots = false;

		this.offsetTop = 0;

		this.x = 0;
		this.y = 0;
		this.width 	= 0;
		this.height = 0;

		this.documentWidth  = 0;
		this.documentHeight = 0;

		this.scrollY = 0;
		this.scrollMaxY = 0;
		this.scrollX = 0;
		this.scrollMaxX = 0;

		this.zoomMode = ZoomMode.Custom;
		this.zoom 	= 1;
		this.zoomCoordinate = null;
		this.skipClearZoomCoord = false;
		
		this.drawingPages = [];
		this.isRepaint = false;
		
		this.canvas = null;
		this.canvasOverlay = null;

		this.Selection = null;

		this.file = null;
		this.isStarted = false;
		this.isCMapLoading = false;
		this.savedPassword = "";

		this.scrollWidth = this.Api.isMobileVersion ? 0 : 14;
		this.isVisibleHorScroll = false;

		this.m_oScrollHorApi = null;
		this.m_oScrollVerApi = null;

		this.backgroundColor = "#E6E6E6";
		this.backgroundPageColor = "#FFFFFF";
		this.outlinePageColor = "#000000";

		this.betweenPages = 20;

		this.moduleState = ModuleState.None;

		this.structure = null;
		this.currentPage = -1;

		this.startVisiblePage = -1;
		this.endVisiblePage = -1;
		this.pagesInfo = new CDocumentPagesInfo();

		this.statistics = {
			paragraph : 0,
			words : 0,
			symbols : 0,
			spaces : 0,
			process : false
		};

		this.handlers = {};

		this.overlay = null;
		this.timerScrollSelect = -1;

		this.SearchResults = null;
		this.isClearPages = false;

		this.isFullText = false;
		this.isFullTextMessage = false;
		this.fullTextMessageCallback = null;
		this.fullTextMessageCallbackArgs = null;

		this.isMouseDown = false;
		this.isMouseMoveBetweenDownUp = false;
		this.mouseMoveEpsilon = 5;
		this.mouseDownCoords = { X : 0, Y : 0 };
		this.mouseDownLinkObject = null;

		this.isFocusOnThumbnails = false;
		this.isDocumentContentReady = false;

		this.doc = new AscPDF.CPDFDoc(this);
		if (typeof CGraphicObjects !== "undefined") {
			this.DrawingObjects = new CGraphicObjects(this.doc, editor.WordControl.m_oDrawingDocument, this.Api);
			this.DrawingObjects.saveDocumentState = null;
		}

		this.isXP = ((AscCommon.AscBrowser.userAgent.indexOf("windowsxp") > -1) || (AscCommon.AscBrowser.userAgent.indexOf("chrome/49") > -1)) ? true : false;
		if (!this.isXP && AscCommon.AscBrowser.isIE && !AscCommon.AscBrowser.isIeEdge)
			this.isXP = true;

		if (this.isXP)
		{
			AscCommon.g_oHtmlCursor.register(AscCommon.Cursors.Grab, "7 8", "pointer");
			AscCommon.g_oHtmlCursor.register(AscCommon.Cursors.Grabbing, "6 6", "pointer");
		}

		var oThis = this;

		this.updateSkin = function()
		{
			this.backgroundColor = AscCommon.GlobalSkin.BackgroundColor;
			this.backgroundPageColor = AscCommon.GlobalSkin.Type === "dark" ? AscCommon.GlobalSkin.BackgroundColor : "#FFFFFF";
			this.outlinePageColor = AscCommon.GlobalSkin.PageOutline;

			if (this.canvas)
				this.canvas.style.backgroundColor = this.backgroundColor;

			if (this.thumbnails)
				this.thumbnails.updateSkin();

			if (this.resize)
				this.resize();
		};

		this.updateDarkMode = function()
		{
			this.isClearPages = true;

			if (this.thumbnails)
			{
				this.thumbnails.updateSkin();
				this.thumbnails.clearCachePages();
			}

			if (this.resize)
				this.resize();
		};

		this.setThumbnailsControl = function(thumbnails)
		{
			this.thumbnails = thumbnails;
			this.thumbnails.viewer = this;
			this.thumbnails.checkPageEmptyStyle();
			if (this.isStarted)
			{
				this.thumbnails.init();
			}
		};

		// events
		this.registerEvent = function(name, handler)
		{
			if (this.handlers[name] === undefined)
				this.handlers[name] = [];
			this.handlers[name].push(handler);
		};
		this.sendEvent = function()
		{
			var name = arguments[0];
			if (this.handlers.hasOwnProperty(name))
			{
				for (var i = 0; i < this.handlers[name].length; ++i)
				{
					this.handlers[name][i].apply(this || window, Array.prototype.slice.call(arguments, 1));
				}
				return true;
			}
		};

		/*
			[TIMER START]
		 */
		this.UseRequestAnimationFrame = AscCommon.AscBrowser.isChrome;
		this.RequestAnimationFrame    = (function()
		{
			return window.requestAnimationFrame ||
				window.webkitRequestAnimationFrame ||
				window.mozRequestAnimationFrame ||
				window.oRequestAnimationFrame ||
				window.msRequestAnimationFrame || null;
		})();
		this.CancelAnimationFrame     = (function()
		{
			return window.cancelRequestAnimationFrame ||
				window.webkitCancelAnimationFrame ||
				window.webkitCancelRequestAnimationFrame ||
				window.mozCancelRequestAnimationFrame ||
				window.oCancelRequestAnimationFrame ||
				window.msCancelRequestAnimationFrame || null;
		})();
		if (this.UseRequestAnimationFrame)
		{
			if (null == this.RequestAnimationFrame)
				this.UseRequestAnimationFrame = false;
		}
		this.RequestAnimationOldTime = -1;

		this.startTimer = function()
		{
			this.isStarted = true;
			if (this.UseRequestAnimationFrame)
				this.timerAnimation();
			else
				this.timer();
		};
		/*
			[TIMER END]
		*/

		this.log = function(message)
		{
			//console.log(message);
		};

		this.timerAnimation = function()
		{
			var now = Date.now();
			if (-1 == oThis.RequestAnimationOldTime || (now >= (oThis.RequestAnimationOldTime + 40)) || (now < oThis.RequestAnimationOldTime))
			{
				oThis.RequestAnimationOldTime = now;
				oThis.timer();
			}
			oThis.RequestAnimationFrame.call(window, oThis.timerAnimation);
		};

		this.timer = function()
		{
			// в порядке важности

			// 1) отрисовка
			// 2) гиперссылки для видимых (и уже отрисованных!) страниц
			// 3) табнейлы (если надо)
			// 4) текстовые команды

			var isViewerTask = oThis.isRepaint;
			if (oThis.isRepaint)
			{
				oThis._paint();
				oThis.onUpdateOverlay();
				oThis.isRepaint = false;
			}
			else if (oThis.checkPagesLinks())
			{
				isViewerTask = true;
			}

			if (oThis.thumbnails)
			{
				isViewerTask = oThis.thumbnails.checkTasks(isViewerTask);
			}

			if (!isViewerTask && !oThis.Api.WordControl.NoneRepaintPages)
			{
				oThis.checkPagesText();

				if (this.isFullTextMessage)
				{
					var countSync = 10;
					while ((countSync > 0) && !this.isFullText)
					{
						oThis.checkPagesText();
						--countSync;
					}
				}
			}

			if (!oThis.UseRequestAnimationFrame)
			{
				setTimeout(oThis.timer, 40);
			}
		};

		this.timerSync = function()
		{
			this.timer();
		};

		this.CreateScrollSettings = function()
		{
			var settings = new AscCommon.ScrollSettings();
			settings.screenW = this.width;
			settings.screenH = this.height;
			settings.vscrollStep = 45;
			settings.hscrollStep = 45;
			settings.isNeedInvertOnActive = GlobalSkin.isNeedInvertOnActive;

			settings.scrollBackgroundColor = GlobalSkin.ScrollBackgroundColor;
			settings.scrollBackgroundColorHover = GlobalSkin.ScrollBackgroundColor;
			settings.scrollBackgroundColorActive = GlobalSkin.ScrollBackgroundColor;

			settings.scrollerColor = GlobalSkin.ScrollerColor;
			settings.scrollerHoverColor = GlobalSkin.ScrollerHoverColor;
			settings.scrollerActiveColor = GlobalSkin.ScrollerActiveColor;

			settings.arrowColor = GlobalSkin.ScrollArrowColor;
			settings.arrowHoverColor = GlobalSkin.ScrollArrowHoverColor;
			settings.arrowActiveColor = GlobalSkin.ScrollArrowActiveColor;

			settings.strokeStyleNone = GlobalSkin.ScrollOutlineColor;
			settings.strokeStyleOver = GlobalSkin.ScrollOutlineHoverColor;
			settings.strokeStyleActive = GlobalSkin.ScrollOutlineActiveColor;

			settings.targetColor = GlobalSkin.ScrollerTargetColor;
			settings.targetHoverColor = GlobalSkin.ScrollerTargetHoverColor;
			settings.targetActiveColor = GlobalSkin.ScrollerTargetActiveColor;
			return settings;
		};

		this.scrollHorizontal = function(pos, maxPos)
		{
			this.scrollX = pos;
			this.scrollMaxX = maxPos;
			if (this.Api.WordControl.MobileTouchManager && this.Api.WordControl.MobileTouchManager.iScroll)
				this.Api.WordControl.MobileTouchManager.iScroll.x = - Math.max(0, Math.min(pos, maxPos));

			this.UpdateDrDocDrawingPages();

			if (this.disabledPaintOnScroll != true)
				this.paint();
				
		};
		this.UpdateDrDocDrawingPages = function() {
			let oThis = this;
			let xCenter = this.width >> 1;
			let yPos = this.scrollY >> 0;
			if (this.documentWidth > this.width)
				xCenter = (this.documentWidth >> 1) - (this.scrollX) >> 0;

			this.getPDFDoc().GetDrawingDocument().m_arrPages = this.drawingPages.map(function(page, index) {
				
				let w = (page.W) >> 0;
				let h = (page.H) >> 0;

				let x = ((xCenter) >> 0) - (w >> 1);
				let y = ((page.Y - yPos)) >> 0;

				return {
					width_mm: page.W / oThis.zoom * g_dKoef_pix_to_mm,
					height_mm: page.H / oThis.zoom * g_dKoef_pix_to_mm,
					drawingPage: {
						left: x,
						top: y,
						right: x + page.W,
						bottom: y + page.H
					}
				}
			});
		};
		this.scrollVertical = function(pos, maxPos)
		{
			this.scrollY = pos;
			this.scrollMaxY = maxPos;
			if (this.Api.WordControl.MobileTouchManager && this.Api.WordControl.MobileTouchManager.iScroll)
				this.Api.WordControl.MobileTouchManager.iScroll.y = - Math.max(0, Math.min(pos, maxPos));

			this.UpdateDrDocDrawingPages();

			if (this.disabledPaintOnScroll != true)
				this.paint();
		};

		this.onLoadModule = function()
		{
			this.moduleState = ModuleState.Loaded;
			window["AscViewer"]["InitializeFonts"](this.Api.baseFontsPath !== undefined ? this.Api.baseFontsPath : undefined);

			if (this._fileData != null)
			{
				this.open(this._fileData);
				delete this._fileData;
			}
		};

		this.checkModule = function()
		{
			if (this.moduleState == ModuleState.Loaded)
			{
				// все загружено - ок
				return true;
			}
			
			if (this.moduleState == ModuleState.Loading)
			{
				// загружается
				return false;
			}

			this.moduleState = ModuleState.Loading;

			var scriptElem = document.createElement('script');
			scriptElem.onerror = function()
			{
				// TODO: пробуем грузить несколько раз
			};

			var _t = this;
			window["AscViewer"]["onLoadModule"] = function() {
				_t.onLoadModule();
			};
			
			var basePath = window["AscViewer"]["baseEngineUrl"];
			
			var useWasm = false;
			var webAsmObj = window["WebAssembly"];
			if (typeof webAsmObj === "object")
			{
				if (typeof webAsmObj["Memory"] === "function")
				{
					if ((typeof webAsmObj["instantiateStreaming"] === "function") || (typeof webAsmObj["instantiate"] === "function"))
						useWasm = true;
				}
			}

			var src = basePath;
			if (useWasm)
				src += "drawingfile.js";
			else
				src += "drawingfile_ie.js";

			scriptElem.setAttribute('src', src);
			scriptElem.setAttribute('type','text/javascript');
			document.getElementsByTagName('head')[0].appendChild(scriptElem);

			return false;
		};

		this.onUpdatePages = function(pages) 
		{
			if (this.startVisiblePage < 0 || this.endVisiblePage < 0)
				return false;

			var isRepaint = false;
			for (var i = 0, len = pages.length; i < len; i++)
			{
				if (pages[i] >= this.startVisiblePage && pages[i] <= this.endVisiblePage)
				{
					isRepaint = true;
					break;
				}
			}

			this.paint();

			if (this.Api && this.Api.printPreview)
				this.Api.printPreview.update();
		};

		this.onUpdateStatistics = function(countParagraph, countWord, countSymbol, countSpace)
		{
			this.statistics.paragraph += countParagraph;
			this.statistics.words += countWord;
			this.statistics.symbols += countSymbol;
			this.statistics.spaces += countSpace;

			if (this.statistics.process)
			{
				this.Api.sync_DocInfoCallback({
					PageCount: this.getPagesCount(),
					WordsCount: this.statistics.words,
					ParagraphCount: this.statistics.paragraph,
					SymbolsCount: this.statistics.symbols,
					SymbolsWSCount: (this.statistics.symbols + this.statistics.spaces)
				});
			}
		};

		this.startStatistics = function()
		{
			this.statistics.process = true;
		};

		this.endStatistics = function()
		{
			this.statistics.process = false;
		};

		this.checkLoadCMap = function()
		{
			if (false === this.isCMapLoading)
			{
				if (!this.file.isNeedCMap())
				{
					this.onDocumentReady();
					return;
				}

				this.isCMapLoading = true;

				this.cmap_load_index = 0;
				this.cmap_load_max = 3;
			}

			var xhr = new XMLHttpRequest();
			let urlCmap = "../../../../sdkjs/pdf/src/engine/cmap.bin";
			if (this.Api.isSeparateModule === true)
				urlCmap = window["AscViewer"]["baseEngineUrl"] + "cmap.bin";

			xhr.open('GET', urlCmap, true);
			xhr.responseType = 'arraybuffer';

			if (xhr.overrideMimeType)
				xhr.overrideMimeType('text/plain; charset=x-user-defined');
			else
				xhr.setRequestHeader('Accept-Charset', 'x-user-defined');

			var _t = this;
			xhr.onload = function()
			{
				if (this.status === 200 || location.href.indexOf("file:") == 0)
				{
					_t.isCMapLoading = false;
					_t.file.setCMap(new Uint8Array(this.response));
					_t.onDocumentReady();
				}
			};
			xhr.onerror = function()
			{
				_t.cmap_load_index++;
				if (_t.cmap_load_index < _t.cmap_load_max)
				{
					_t.checkLoadCMap();
					return;
				}

				// error!
				_t.isCMapLoading = false;
				_t.onDocumentReady();
			};

			xhr.send(null);
		};

		this.onDocumentReady = function()
		{
			var _t = this;
			// в интерфейсе есть проблема - нужно посылать onDocumentContentReady после setAdvancedOptions
			setTimeout(function(){

				if (!_t.isStarted)
				{
					AscCommon.addMouseEvent(_t.canvasForms, "down", _t.onMouseDown);
					AscCommon.addMouseEvent(_t.canvasForms, "move", _t.onMouseMove);
					AscCommon.addMouseEvent(_t.canvasForms, "up", _t.onMouseUp);

					_t.parent.onmousewheel = _t.onMouseWhell;
					if (_t.parent.addEventListener)
						_t.parent.addEventListener("DOMMouseScroll", _t.onMouseWhell, false);

					_t.startTimer();
				}

				_t.sendEvent("onFileOpened");

				_t.sendEvent("onPagesCount", _t.file.pages.length);
				_t.sendEvent("onCurrentPageChanged", 0);

				_t.sendEvent("onStructure", _t.structure);
			}, 0);

			this.file.onRepaintPages = this.onUpdatePages.bind(this);
			this.file.onUpdateStatistics = this.onUpdateStatistics.bind(this);
			this.currentPage = -1;
			this.structure = this.file.getStructure();

			this.resize(true);
			
			this.openForms();
			this.openAnnots();
			
			if (this.thumbnails)
				this.thumbnails.init(this);

			this.setMouseLockMode(true);
		};

		this.open = function(data, password)
		{
			if (!this.checkModule())
			{
				this._fileData = data;
				return;
			}

			if (undefined !== password)
			{
				if (!this.file)
				{
					this.file = window["AscViewer"].createFile(data);

					if (this.file)
					{
						this.SearchResults = this.file.SearchResults;
						this.file.viewer = this;
					}
				}

				if (this.file && this.file.isNeedPassword())
				{
					window["AscViewer"].setFilePassword(this.file, password);
					this.Api.currentPassword = password;
				}
			}
			else
			{
				if (this.file)
					this.close();

				this.file = window["AscViewer"].createFile(data);
				
				if (this.file)
				{
					this.SearchResults = this.file.SearchResults;
					this.file.viewer = this;
				}
			}

			if (!this.file)
			{
				this.Api.sendEvent("asc_onError", Asc.c_oAscError.ID.ConvertationOpenError, Asc.c_oAscError.Level.Critical);
				return;
			}

			var _t = this;
			if (this.file.isNeedPassword())
			{
				// при повторном вводе пароля - проблемы в интерфейсе, если синхронно
				setTimeout(function(){
					_t.sendEvent("onNeedPassword");
				}, 100);
				return;
			}

			if (window["AscDesktopEditor"])
				this.savedPassword = password;

			this.pagesInfo.setCount(this.file.pages.length);
            this.getPDFDoc().GetDrawingDocument().m_lPagesCount = this.file.pages.length;

			for (let i = 0; i < this.file.pages.length; i++) {
				this.DrawingObjects.mergeDrawings(i);
			}
			
			this.checkLoadCMap();
			
			AscCommon.g_oIdCounter.Set_Load(false); // to do возможно не тут стоит выключать флаг
		};
		this.close = function()
		{
			if (!this.file || !this.file.isValid())
				return;

			this.file.close();

			this.structure = null;
			this.currentPage = -1;

			this.startVisiblePage = -1;
			this.endVisiblePage = -1;
			this.pagesInfo = new CDocumentPagesInfo();
			this.drawingPages = [];

			this.statistics = {
				paragraph : 0,
				words : 0,
				symbols : 0,
				spaces : 0,
				process : false
			};

			this._paint();
		};

		this.getFileNativeBinary = function()
		{
			if (!this.file || !this.file.isValid())
				return null;
			return this.file.getFileBinary();
		};

		this.openForms = function() {
			let oThis = this;

			this.scrollCount = 0;
			this.IsOpenFormsInProgress = true;

			function ExtractActions(oPanentAction) {
				let aActions = [];
				const propToRemove = 'Next';
			  
				const keys = Object.keys(oPanentAction).filter(function(key) {
					return key !== propToRemove;
				});
			  
				const tempObject = {};
			  
				for (let i = 0; i < keys.length; i++) {
					const key = keys[i];
					tempObject[key] = oPanentAction[key];
				}
			  
				aActions.push(tempObject);
			  
				if (oPanentAction["Next"]) {
					const nextActions = ExtractActions(oPanentAction["Next"]);
					aActions = aActions.concat(nextActions);
				}
			  
				return aActions;
			}
			
			let aActionsToCorrect = []; // параметры поля в actions указаны как ссылки на ap, после того, как все формы будут созданы, заменим их на ссылки на сами поля. 
			let aFormsInfo = this.file.nativeFile["getInteractiveFormsInfo"]();
			
			let oFormInfo, oForm, oRect;
			if (aFormsInfo["Fields"] == null) {
				this.IsOpenFormsInProgress = false;
				return;
			}
				
			for (let i = 0; i < aFormsInfo["Fields"].length; i++)
			{
				oFormInfo	= aFormsInfo["Fields"][i];
				oRect		= oFormInfo["rect"];

				oForm = this.doc.AddField(oFormInfo["name"], oFormInfo["type"], oFormInfo["page"], [oRect["x1"], oRect["y1"], oRect["x2"], oRect["y2"]]);
				
				if (!oForm) {
					// console.log("Error while reading form, index " + oFormInfo["AP"]["i"]);
					continue;
				}

				oForm.SetOriginPage(oFormInfo["page"]);

				if (oFormInfo["Parent"] != null)
				{
					oForm.AddToChildsMap(oFormInfo["Parent"]);
				}

				// appearance
				if (oFormInfo["border"] != null)
				{
					oForm.SetBorderStyle(oFormInfo["border"]);
				}

				if (oFormInfo["borderWidth"] != null)
				{
					oForm.SetBorderWidth(oFormInfo["borderWidth"]);
				}
				if (oFormInfo["BC"] != null)
				{
					oForm.SetBorderColor(oFormInfo["BC"]);
				}
				if (oFormInfo["BG"] != null)
					oForm.SetBackgroundColor(oFormInfo["BG"]);

				if (oFormInfo["font"]) {
					if (oFormInfo["font"]["color"] != null)
						oForm.SetTextColor(oFormInfo["font"]["color"]);
					// if (oFormInfo["font"]["name"] != null)
					// 	oForm.SetTextFont(oFormInfo["font"]["name"]);
					if (oFormInfo["font"]["size"] != null)
						oForm.SetTextSize(oFormInfo["font"]["size"]);
				}

				if (oFormInfo["AP"] != null) {
					oForm.SetApIdx(oFormInfo["AP"]["i"]);
					oForm.SetDrawFromStream(Boolean(oFormInfo["AP"]["have"]));
				}

				// text form
				if (oFormInfo["multiline"] != null)
				{
					oForm.SetMultiline(Boolean(oFormInfo["multiline"]));
				}
				if (oFormInfo["comb"])
				{
					oForm.SetComb(Boolean(oFormInfo["comb"]));
				}
				if (oFormInfo["richText"])
				{
					// to do
					oForm.SetRichText(Boolean(oFormInfo["richText"]));
				}
				if (oFormInfo["password"])
				{
					// to do
					oForm.SetPassword(Boolean(oFormInfo["password"]));
				}

				// button
				if (oFormInfo["position"] != null) {
					oForm.SetButtonPosition(oFormInfo["position"]);
				}
				if (oFormInfo["caption"] != null && oForm["type"] == AscPDF.FIELD_TYPES.button) {
					oForm.SetCaption(oFormInfo["caption"]);
				}
				if (oFormInfo["alternateCaption"] != null && oForm["type"] == AscPDF.FIELD_TYPES.button) {
					oForm.SetCaption(oFormInfo["alternateCaption"], AscPDF.CAPTION_TYPES.mouseDown);
				}
				if (oFormInfo["rolloverCaption"] != null && oForm["type"] == AscPDF.FIELD_TYPES.button) {
					oForm.SetCaption(oFormInfo["rolloverCaption"], AscPDF.CAPTION_TYPES.rollover);
				}
				
				if (oFormInfo["highlight"] != null) {
					oForm.SetHighlight(oFormInfo["highlight"]);
				}
				if (oFormInfo["IF"] != null) {
					if (oFormInfo["IF"]["FB"] != null)
						oForm.SetButtonFitBounds(Boolean(oFormInfo["IF"]["FB"]));
					if (oFormInfo["IF"]["SW"] != null)
						oForm.SetScaleWhen(oFormInfo["IF"]["SW"]);
					if (oFormInfo["IF"]["A"] != null)
						oForm.SetIconPosition(oFormInfo["IF"]["A"][0], oFormInfo["IF"]["A"][1]);
					if (oFormInfo["IF"]["S"] != null)
						oForm.SetScaleHow(oFormInfo["IF"]["S"]);
				}

				// combobox - listbox
				if (oFormInfo["editable"])
				{
					oForm.SetEditable(Boolean(oFormInfo["editable"]));
				}
				if (oFormInfo["commitOnSelChange"])
				{
					// to do
					oForm.SetCommitOnSelChange(Boolean(oFormInfo["commitOnSelChange"]));
				}
				if (oFormInfo["multipleSelection"])
				{
					oForm.SetMultipleSelection(Boolean(oFormInfo["multipleSelection"]));
				}
				if (oFormInfo["opt"])
				{
					oForm.SetOptions(oFormInfo["opt"]);
				}

				// checkbox - radiobutton
				if (oFormInfo["ExportValue"])
				{
					oForm.SetExportValue(oFormInfo["ExportValue"]);
				}
				if (oFormInfo["radiosInUnison"])
				{
					oForm.SetRadiosInUnison(Boolean(oFormInfo["radiosInUnison"]));
				}
				if (oFormInfo["NoToggleToOff"] != null && oFormInfo["type"] != AscPDF.FIELD_TYPES.button)
				{
					oForm.SetNoToggleToOff(Boolean(oFormInfo["NoToggleToOff"]));
				}
				if (oFormInfo["style"] != null)
				{
					oForm.SetStyle(oFormInfo["style"]);
				}

				// signature
				if (oFormInfo["Sig"] != null)
				{
					oForm.SetFilled(Boolean(oFormInfo["Sig"]));
				}

				// common
				if (oFormInfo["alignment"] != null && [AscPDF.FIELD_TYPES.combobox, AscPDF.FIELD_TYPES.text].includes(oFormInfo["type"]))
				{
					oForm.SetAlign(oFormInfo["alignment"]);
				}
				if (oFormInfo["maxLen"] != null)
				{
					oForm.SetCharLimit(oFormInfo["maxLen"]);
				}
				if (oFormInfo["doNotScroll"] != null)
				{
					oForm.SetDoNotScroll(Boolean(oFormInfo["doNotScroll"]));
				}
				if (oFormInfo["doNotSpellCheck"] != null && [AscPDF.FIELD_TYPES.combobox, AscPDF.FIELD_TYPES.text].includes(oFormInfo["type"]))
				{
					// to do
					oForm.SetDoNotSpellCheck(Boolean(oFormInfo["doNotSpellCheck"]));
				}
				if (oFormInfo["fileSelect"])
				{
					// to do
					oForm.SetFileSelect(Boolean(oFormInfo["fileSelect"]));
				}
				if (oFormInfo["flag"] != null)
				{
					// to do
				}
				if (oFormInfo["noexport"])
				{
					// to do
				}
				if (oFormInfo["readonly"])
				{
					// to do
					oForm.SetReadOnly(Boolean(oFormInfo["readonly"]));
				}
				if (oFormInfo["required"])
				{
					// to do
					oForm.SetRequired(Boolean(oFormInfo["required"]));
				}
				if (oFormInfo["value"] != null && oForm.GetType() != AscPDF.FIELD_TYPES.button)
				{
					oForm.SetValue(oFormInfo["value"], true);
				}
				if (oFormInfo["display"])
				{
					// to do
					oForm.SetDisplay(oFormInfo["display"]);
				}
				
				if (oFormInfo["sort"] != null) {
					// to do sort
				}
				if (oFormInfo["defaultValue"] != null) {
					oForm.SetDefaultValue(oFormInfo["defaultValue"]);
				}

				// actions
				if (oFormInfo["AA"] != null)
				{
					// mouseup 0
					if (oFormInfo["AA"]["A"])
					{
						aActionsToCorrect = aActionsToCorrect.concat(oForm.SetActionsOnOpen(AscPDF.FORMS_TRIGGERS_TYPES.MouseUp, ExtractActions(oFormInfo["AA"]["A"])));
					}
					// mousedown 1
					if (oFormInfo["AA"]["D"])
					{
						aActionsToCorrect = aActionsToCorrect.concat(oForm.SetActionsOnOpen(AscPDF.FORMS_TRIGGERS_TYPES.MouseDown, ExtractActions(oFormInfo["AA"]["D"])));
					}
					// mouseenter 2
					if (oFormInfo["AA"]["E"])
					{
						aActionsToCorrect = aActionsToCorrect.concat(oForm.SetActionsOnOpen(AscPDF.FORMS_TRIGGERS_TYPES.MouseEnter, ExtractActions(oFormInfo["AA"]["E"])));
					}
					// mouseexit 3
					if (oFormInfo["AA"]["X"])
					{
						aActionsToCorrect = aActionsToCorrect.concat(oForm.SetActionsOnOpen(AscPDF.FORMS_TRIGGERS_TYPES.MouseExit, ExtractActions(oFormInfo["AA"]["X"])));
					}
					// onFocus 4
					if (oFormInfo["AA"]["Fo"])
					{
						aActionsToCorrect = aActionsToCorrect.concat(oForm.SetActionsOnOpen(AscPDF.FORMS_TRIGGERS_TYPES.OnFocus, ExtractActions(oFormInfo["AA"]["Fo"])));
					}
					// onBlur 5
					if (oFormInfo["AA"]["Bl"])
					{
						aActionsToCorrect = aActionsToCorrect.concat(oForm.SetActionsOnOpen(AscPDF.FORMS_TRIGGERS_TYPES.OnBlur, ExtractActions(oFormInfo["AA"]["Bl"])));
					}

					// keystroke 6
					if (oFormInfo["AA"]["K"])
					{
						aActionsToCorrect = aActionsToCorrect.concat(oForm.SetActionsOnOpen(AscPDF.FORMS_TRIGGERS_TYPES.Keystroke, ExtractActions(oFormInfo["AA"]["K"])));
					}

					// Validate 7
					if (oFormInfo["AA"]["V"])
					{
						aActionsToCorrect = aActionsToCorrect.concat(oForm.SetActionsOnOpen(AscPDF.FORMS_TRIGGERS_TYPES.Validate, ExtractActions(oFormInfo["AA"]["V"])));
					}

					// Calculate 8
					if (oFormInfo["AA"]["C"])
					{
						aActionsToCorrect = aActionsToCorrect.concat(oForm.SetActionsOnOpen(AscPDF.FORMS_TRIGGERS_TYPES.Calculate, ExtractActions(oFormInfo["AA"]["C"])));
					}

					// format 9
					if (oFormInfo["AA"]["F"])
					{
						aActionsToCorrect = aActionsToCorrect.concat(oForm.SetActionsOnOpen(AscPDF.FORMS_TRIGGERS_TYPES.Format, ExtractActions(oFormInfo["AA"]["F"])));
					}
				}
			}
			
			if (aFormsInfo["Parents"]) {
				this.doc.FillFormsParents(aFormsInfo["Parents"]);
				this.doc.OnAfterFillFormsParents();
			}
			
			this.doc.FillButtonsIconsOnOpen();
			
			if (Array.isArray(aFormsInfo["CO"]) && aFormsInfo["CO"].length > 0)
				this.doc.GetCalculateInfo().SetCalculateOrder(aFormsInfo["CO"]);
			
			// после открытия всех форм, заменяем apIdx в Actions на ссылки на сами поля.
			aActionsToCorrect.forEach(function(action) {
				if (action.fields) {
					let aNewArrFields =  oThis.doc.widgets.filter(function(field) {
						return action.fields.includes(field._apIdx);
					});
	
					action.fields = aNewArrFields;
				}
			});

			this.IsOpenFormsInProgress = false;
		};
		this.openAnnots = function() {
			let oThis = this;

			let oAnnotsMap = {};
			let oDoc = this.getPDFDoc();
			let aAnnotsInfo = this.file.nativeFile["getAnnotationsInfo"]();
			let nMaxIdx		= this.file.nativeFile["getStartID"]();

			this.IsOpenAnnotsInProgress = true;

			let oAnnotInfo, oAnnot, aRect;
			for (let i = 0; i < aAnnotsInfo.length; i++) {
				oAnnotInfo = aAnnotsInfo[i];

				if (oAnnotInfo["Type"] == AscPDF.ANNOTATIONS_TYPES.Popup) {
					continue;
				}

				aRect = [oAnnotInfo["rect"]["x1"], oAnnotInfo["rect"]["y1"], oAnnotInfo["rect"]["x2"], oAnnotInfo["rect"]["y2"]];

				if (oAnnotInfo["RefTo"] == null || oAnnotInfo["Type"] != AscPDF.ANNOTATIONS_TYPES.Text) {
					oAnnot = oDoc.AddAnnot({
						page:			oAnnotInfo["page"],
						name:			oAnnotInfo["UniqueName"], 
						creationDate:	oAnnotInfo["CreationDate"] ? AscPDF.ParsePDFDate(oAnnotInfo["CreationDate"]).getTime() : undefined,
						modDate:		oAnnotInfo["LastModified"] ? AscPDF.ParsePDFDate(oAnnotInfo["LastModified"]).getTime() : undefined,
						contents:		oAnnotInfo["Contents"],
						author:			oAnnotInfo["User"],
						rect:			aRect,
						type:			oAnnotInfo["Type"],
						apIdx:			oAnnotInfo["AP"]["i"]
					});
	
					oAnnot.SetDrawFromStream(Boolean(oAnnotInfo["AP"]["have"]));
					oAnnot.SetOriginPage(oAnnotInfo["page"]);

					if (oAnnotInfo["RefTo"] == null)
						oAnnotsMap[oAnnotInfo["AP"]["i"]] = oAnnot;

					if (oAnnotInfo["InkList"]) {
						oAnnot.SetInkPoints(oAnnotInfo["InkList"]);
					}
					else if (oAnnotInfo["L"]) {
						oAnnot.SetLinePoints(oAnnotInfo["L"]);
					}
					else if (oAnnotInfo["Vertices"]) {
						oAnnot.SetVertices(oAnnotInfo["Vertices"]);
					}

					if (oAnnotInfo["LE"] != null) {
						if (Array.isArray(oAnnotInfo["LE"])) {
							oAnnot.SetLineStart(oAnnotInfo["LE"][0]);
							oAnnot.SetLineEnd(oAnnotInfo["LE"][1]);
						}
						else
							oAnnot.SetLineEnd(oAnnotInfo["LE"]);
					}

					if (oAnnotInfo["RefToReason"] != null)
						oAnnot.SetRefType(oAnnotInfo["RefToReason"]);
					if (oAnnotInfo["Popup"] != null)
						oAnnot.SetPopupIdx(oAnnotInfo["Popup"]);
					if (oAnnotInfo["Subj"])
						oAnnot.SetSubject(oAnnotInfo["Subj"]);
					if (oAnnotInfo["CL"])
						oAnnot.SetCallout(oAnnotInfo["CL"]);
					if (oAnnotInfo["RC"])
						oAnnot.SetRichContents(oAnnotInfo["RC"]);
					if (oAnnotInfo["RD"])
						oAnnot.SetReqtangleDiff(oAnnotInfo["RD"]);
					if (oAnnotInfo["display"])
						oAnnot.SetDisplay(oAnnotInfo["display"]);
					if (oAnnotInfo["locked"] != null)
						oAnnot.SetLock(Boolean(oAnnotInfo["locked"]));
					if (oAnnotInfo["lockedC"] != null)
						oAnnot.SetLockContent(Boolean(oAnnotInfo["lockedC"]));
					if (oAnnotInfo["IC"] != null)
						oAnnot.SetFillColor(oAnnotInfo["IC"]);
					if (oAnnotInfo["dashed"] != null)
						oAnnot.SetDash(oAnnotInfo["dashed"]);
					if (oAnnotInfo["border"] != null)
						oAnnot.SetBorder(oAnnotInfo["border"]);
					
					if (oAnnotInfo["noRotate"] != null)
						oAnnot.SetNoRotate(Boolean(oAnnotInfo["noRotate"]));
					if (oAnnotInfo["noZoom"] != null)
						oAnnot.SetNoZoom(Boolean(oAnnotInfo["noZoom"]));
					if (oAnnotInfo["Sy"] != null)
						oAnnot.SetCaretSymbol(oAnnotInfo["Sy"]);
					if (oAnnotInfo["LL"] != null)
						oAnnot.SetLeaderLength(oAnnotInfo["LL"]);
					if (oAnnotInfo["LLE"] != null)
						oAnnot.SetLeaderExtend(oAnnotInfo["LLE"]);
					if (oAnnotInfo["Cap"] != null)
						oAnnot.SetDoCaption(Boolean(oAnnotInfo["Cap"]));

					// FreeText/Redact
					if (oAnnotInfo["alignment"] != null)
						oAnnot.SetAlign(oAnnotInfo["alignment"]);

					// FreeText
					if (oAnnotInfo["defaultStyle"] != null)
						oAnnot.SetDefaultStyle(oAnnotInfo["defaultStyle"]);
					
					// border effect
					if (oAnnotInfo["BE"] != null) {
						if (oAnnotInfo["BE"]["I"] != null)
							oAnnot.SetBorderEffectIntensity(oAnnotInfo["BE"]["I"]);
						if (oAnnotInfo["BE"]["S"] != null)
							oAnnot.SetBorderEffectStyle(oAnnotInfo["BE"]["S"]);
					}
						
					oAnnot.SetStrokeColor(oAnnotInfo["C"]);
					
					if (oAnnotInfo["CA"] != null) {
						oAnnot.SetOpacity(oAnnotInfo["CA"]);
					}
					if (oAnnotInfo["borderWidth"] != null) {
						oAnnot.SetWidth(oAnnotInfo["borderWidth"]);
					}
					if (oAnnotInfo["QuadPoints"] != null) {
						let aSepQuads = [];
						for (let i = 0; i < oAnnotInfo["QuadPoints"].length; i+=8)
							aSepQuads.push(oAnnotInfo["QuadPoints"].slice(i, i+8));

						oAnnot.SetQuads(aSepQuads);
					}
				}
				else {
					if (oAnnotInfo["StateModel"] != AscPDF.TEXT_ANNOT_STATE_MODEL.Review && oAnnotsMap[oAnnotInfo["RefTo"]] && oAnnotsMap[oAnnotInfo["RefTo"]]._AddReplyOnOpen)
						oAnnotsMap[oAnnotInfo["RefTo"]]._AddReplyOnOpen(oAnnotInfo);
				}
			}

			for (let apIdx in oAnnotsMap) {
				if (oAnnotsMap[apIdx] instanceof AscPDF.CAnnotationText || oAnnotsMap[apIdx].GetReply(0) instanceof AscPDF.CAnnotationText)
					oAnnotsMap[apIdx]._OnAfterSetReply();
			}
			this.IsOpenAnnotsInProgress = false;

			oDoc.UpdateApIdx(nMaxIdx);
		};
		this.setZoom = function(value, isDisablePaint)
		{
			this.zoom = value;
			this.zoomMode = ZoomMode.Custom;
			this.sendEvent("onZoom", this.zoom);

			if (this.Api.watermarkDraw)
			{
				this.Api.watermarkDraw.zoom = this.zoom;
				this.Api.watermarkDraw.Generate();
			}

			this.resize(isDisablePaint);
		};
		this.setZoomMode = function(value)
		{
			this.zoomMode = value;
			this.resize();
		};
		this.calculateZoomToWidth = function()
		{
			if (!this.file || !this.file.isValid())
				return 1;

			var maxWidth = 0;
			for (let i = 0, len = this.file.pages.length; i < len; i++)
			{
				var pageW = (this.file.pages[i].W * 96 / this.file.pages[i].Dpi);
				if (pageW > maxWidth)
					maxWidth = pageW;
			}

			if (maxWidth < 1)
				return 1;

			return (this.width - 2 * this.betweenPages) / maxWidth;
		};
		this.calculateZoomToHeight = function()
		{
			if (!this.file || !this.file.isValid())
				return 1;

			var maxHeight = 0;
			var maxWidth = 0;
			for (let i = 0, len = this.file.pages.length; i < len; i++)
			{
				var pageW = (this.file.pages[i].W * 96 / this.file.pages[i].Dpi);
				var pageH = (this.file.pages[i].H * 96 / this.file.pages[i].Dpi);
				if (pageW > maxWidth)
					maxWidth = pageW;
				if (pageH > maxHeight)
					maxHeight = pageH;
			}

			if (maxWidth < 1 || maxHeight < 1)
				return 1;

			var zoom1 = (this.width - 2 * this.betweenPages) / maxWidth;
			var zoom2 = (this.height - 2 * this.betweenPages) / maxHeight;

			return Math.min(zoom1, zoom2);
		};
		this.fixZoomCoord = function(x, y)
		{
			if (this.Api.isMobileVersion)
			{
				x -= this.x;
				y -= this.y;
			}
			this.zoomCoordinate = this.getPageByCoords2(x, y);
			if (this.zoomCoordinate)
			{
				this.zoomCoordinate.xShift = x;
				this.zoomCoordinate.yShift = y;
			}
		};

		this.clearZoomCoord = function()
		{
			// нужно очищать, чтобы при любом ресайзе мы не скролились к последней сохранённой точке
			this.zoomCoordinate = null;
		};

		this.getFirstPagePosition = function()
		{
			let lPagesCount = this.drawingPages.length;
			for (let i = 0; i < lPagesCount; i++)
			{
				let page = this.drawingPages[i];
				if ((page.Y + page.H) > this.scrollY)
				{
					return {
						page : i,
						x : page.X,
						y : page.Y,
						scrollX : this.scrollX,
						scrollY : this.scrollY
					};
				}
			}
			return null;
		};

		this.setMouseLockMode = function(isEnabled)
		{
			this.MouseHandObject = isEnabled ? {} : null;
		};

		this.getPagesCount = function()
		{
			if (!this.file || !this.file.isValid())
				return 0;
			return this.file.pages.length;
		};

		this.getDocumentInfo = function()
		{
			if (!this.file || !this.file.isValid())
				return 0;
			return this.file.getDocumentInfo();
		};

		this.navigate = function(id)
		{
			var item = this.structure[id];
			if (!item)
				return;

			var pageIndex = item["page"];
			var drawingPage = this.drawingPages[pageIndex];
			if (!drawingPage)
				return;

			var posY = drawingPage.Y;
			posY -= this.betweenPages;

			var yOffset = item["y"];
			if (yOffset)
			{
				yOffset *= (drawingPage.H / this.file.pages[pageIndex].H);
				yOffset = yOffset >> 0;
				posY += yOffset;
			}

			if (posY > this.scrollMaxY)
				posY = this.scrollMaxY;
			this.m_oScrollVerApi.scrollToY(posY);
		};

		this.navigateToPage = function(pageNum, yOffset, xOffset)
		{
			var drawingPage = this.drawingPages[pageNum];
			if (!drawingPage)
				return;

			var posY = drawingPage.Y;
			posY -= this.betweenPages;

			if (yOffset)
			{
				yOffset *= (drawingPage.H / this.file.pages[pageNum].H);
				yOffset = yOffset >> 0;
				posY += yOffset;
			}

			if (posY > this.scrollMaxY)
				posY = this.scrollMaxY;

			var posX = 0;

			if (xOffset)
			{
				xOffset *= (drawingPage.W / this.file.pages[pageNum].W);
				xOffset = xOffset >> 0;
				posX += xOffset;
			}

			if (posX > this.scrollMaxX)
				posX = this.scrollMaxX;

			let oDoc = this.getPDFDoc();
			// выход из формы если вышли со страницы, где находится активная форма.
			if (this.disabledPaintOnScroll == false && oDoc.activeForm && this.pageDetector.pages.map(function(item) {
				return item.num;
			}).includes(oDoc.activeForm.GetPage()) == false) {
				if (oDoc.activeForm.IsChanged() == false) {
					oDoc.activeForm.SetDrawFromStream(true);
				}
				oDoc.activeForm.SetDrawHighlight(true);
				oDoc.activeForm = null;
			}

			this.m_oScrollVerApi.scrollToY(posY);
			this.m_oScrollHorApi.scrollToX(posX);
		};

		this.navigateToLink = function(link)
		{
			if ("" === link["link"])
				return;

			if ("#" === link["link"].charAt(0))
			{
				this.navigateToPage(parseInt(link["link"].substring(1)), link["dest"]);
			}
			else
			{
				var url = link["link"];
				var typeUrl = AscCommon.getUrlType(url);
				url = AscCommon.prepareUrl(url, typeUrl);
				this.sendEvent("onHyperlinkClick", url);
			}

			//console.log(link["link"]);
		};

		this.setTargetType = function(type)
		{
			this.setMouseLockMode(type == "hand");
		};

		this.updateCurrentPage = function(pageObject)
		{
			if (this.currentPage != pageObject.num)
			{
				this.currentPage = pageObject.num;
				this.Api.WordControl.m_oDrawingDocument.m_lCurrentPage = pageObject.num;
				this.sendEvent("onCurrentPageChanged", this.currentPage);
			}

			if (this.thumbnails)
				this.thumbnails.updateCurrentPage(pageObject);
		};

		this.recalculatePlaces = function()
		{
			if (!this.file || !this.file.isValid())
				return;

			// здесь картинки не обнуляем
			for (let i = 0, len = this.file.pages.length; i < len; i++)
			{
				if (!this.drawingPages[i])
				{
					this.drawingPages[i] = {
						X : 0,
						Y : 0,
						W : (this.file.pages[i].W * 96 * this.zoom / this.file.pages[i].Dpi) >> 0,
						H : (this.file.pages[i].H * 96 * this.zoom / this.file.pages[i].Dpi) >> 0,
						Image : undefined
					};
				}
				else
				{
					this.drawingPages[i].X = 0;
					this.drawingPages[i].Y = 0;
					this.drawingPages[i].W = (this.file.pages[i].W * 96 * this.zoom / this.file.pages[i].Dpi) >> 0;
					this.drawingPages[i].H = (this.file.pages[i].H * 96 * this.zoom / this.file.pages[i].Dpi) >> 0;
				}

				if (this.getPageRotate(i) & 1)
				{
					let tmp = this.drawingPages[i].W;
					this.drawingPages[i].W = this.drawingPages[i].H;
					this.drawingPages[i].H = tmp;
				}
			}

			this.documentWidth = 0;
			for (let i = 0, len = this.drawingPages.length; i < len; i++)
			{
				if (this.drawingPages[i].W > this.documentWidth)
					this.documentWidth = this.drawingPages[i].W;
			}
			// прибавим немного
			this.documentWidth += (4 * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

			var curTop = this.betweenPages + this.offsetTop;
			for (let i = 0, len = this.drawingPages.length; i < len; i++)
			{
				this.drawingPages[i].X = (this.documentWidth - this.drawingPages[i].W) >> 1;
				this.drawingPages[i].Y = curTop;

				curTop += this.drawingPages[i].H;
				curTop += this.betweenPages;
			}

			this.documentHeight = curTop;

			let xCenter = this.width >> 1;
			let yPos = this.scrollY >> 0;
			if (this.documentWidth > this.width)
				xCenter = (this.documentWidth >> 1) - (this.scrollX) >> 0;

			let oThis = this;
			this.getPDFDoc().GetDrawingDocument().m_arrPages = this.drawingPages.map(function(page, index) {
				
				let w = (page.W) >> 0;
				let h = (page.H) >> 0;

				let x = ((xCenter) >> 0) - (w >> 1);
				let y = ((page.Y - yPos)) >> 0;

				return {
					width_mm: page.W / oThis.zoom * g_dKoef_pix_to_mm,
					height_mm: page.H / oThis.zoom * g_dKoef_pix_to_mm,
					drawingPage: {
						left: x,
						top: y,
						right: x + page.W,
						bottom: y + page.H
					}
				}
			});
			
			this.paint();
		};

		this.setCursorType = function(cursor)
		{
			if (this.Api.isDrawInkMode()) {
				this.id_main.style.cursor = "";
				return;
			}

			if (this.isXP)
			{
				this.id_main.style.cursor = AscCommon.g_oHtmlCursor.value(cursor);
				return;
			}

			this.id_main.style.cursor = cursor;
		};

		this.getPageLinkByMouse = function()
		{
			var pageObject = this.getPageByCoords(AscCommon.global_mouseEvent.X - this.x, AscCommon.global_mouseEvent.Y - this.y);
			if (!pageObject)
				return null;

			var pageLinks = this.pagesInfo.pages[pageObject.index];
			if (pageLinks.links)
			{
				for (var i = 0, len = pageLinks.links.length; i < len; i++)
				{
					if (pageObject.x >= pageLinks.links[i]["x"] && pageObject.x <= (pageLinks.links[i]["x"] + pageLinks.links[i]["w"]) &&
						pageObject.y >= pageLinks.links[i]["y"] && pageObject.y <= (pageLinks.links[i]["y"] + pageLinks.links[i]["h"]))
					{
						return pageLinks.links[i];
					}
				}
			}
			return null;
		};
		this.getPageFieldByMouse = function(bGetHidden)
		{
			var pageObject = this.getPageByCoords(AscCommon.global_mouseEvent.X - this.x, AscCommon.global_mouseEvent.Y - this.y);
			if (!pageObject)
				return null;

			var pageFields = this.pagesInfo.pages[pageObject.index];
			if (pageFields.fields)
			{
				for (var i = 0, len = pageFields.fields.length; i < len; i++)
				{
					if (pageObject.x >= pageFields.fields[i]._origRect[0] && pageObject.x <= pageFields.fields[i]._origRect[2] &&
						pageObject.y >= pageFields.fields[i]._origRect[1] && pageObject.y <= pageFields.fields[i]._origRect[3])
					{
						if (bGetHidden)
							return pageFields.fields[i];
						else
							return pageFields.fields[i].IsHidden() == false ? pageFields.fields[i] : null;
						
					}
				}
			}
			return null;
		};
		this.getPageAnnotByMouse = function(bGetHidden)
		{
			let oDoc = this.getPDFDoc();
			let oDrDoc = oDoc.GetDrawingDocument();
			var pageObject = this.getPageByCoords(AscCommon.global_mouseEvent.X - this.x, AscCommon.global_mouseEvent.Y - this.y);
			if (!pageObject)
				return null;

			var page = this.pagesInfo.pages[pageObject.index];
			
			// если есть заселекченная аннотация под мышкой, то возвращаем её, а не первую попавшуюся
			if (oDoc.mouseDownAnnot) {
				let oPos	= oDrDoc.ConvertCoordsFromCursor2(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
				let X       = oPos.X;
				let Y       = oPos.Y;

				if (oDoc.mouseDownAnnot.hitInPath && oDoc.mouseDownAnnot.hitInPath(X, Y))
					return oDoc.mouseDownAnnot;
			}

			if (page.annots)
			{
				// сначала ищем text annot (sticky note)
				for (var i = page.annots.length -1; i >= 0; i--)
				{
					let oAnnot = page.annots[i];
					let nAnnotWidth		= Math.max(oAnnot._origRect[2] - oAnnot._origRect[0], 32) / (this.zoom * AscCommon.AscBrowser.retinaPixelRatio);
					let nAnnotHeight	= Math.max(oAnnot._origRect[3] - oAnnot._origRect[1], 32) / (this.zoom * AscCommon.AscBrowser.retinaPixelRatio);
					
					if (true !== bGetHidden && oAnnot.IsHidden() == true || false == oAnnot.IsComment())
						continue;
					
					if (pageObject.x >= oAnnot._origRect[0] && pageObject.x <= oAnnot._origRect[0] + nAnnotWidth &&
						pageObject.y >= oAnnot._origRect[1] && pageObject.y <= oAnnot._origRect[1] + nAnnotHeight)
					{
						if (bGetHidden)
							return oAnnot;
						else
							return oAnnot.IsHidden() == false ? oAnnot : null;
					}
				}

				for (var i = page.annots.length -1; i >= 0; i--)
				{
					let oAnnot = page.annots[i];
					let nAnnotWidth		= (oAnnot._origRect[2] - oAnnot._origRect[0]);
					let nAnnotHeight	= (oAnnot._origRect[3] - oAnnot._origRect[1]);
					
					if (true !== bGetHidden && oAnnot.IsHidden() == true || oAnnot.IsComment())
						continue;
					
					if (pageObject.x >= oAnnot._origRect[0] && pageObject.x <= oAnnot._origRect[0] + nAnnotWidth &&
						pageObject.y >= oAnnot._origRect[1] && pageObject.y <= oAnnot._origRect[1] + nAnnotHeight)
					{
						// у маркап аннотаций ищем по quads (т.к. rect too wide)
						if (oAnnot.IsTextMarkup())
						{
							if (oAnnot.IsInQuads(pageObject.x, pageObject.y))
								return oAnnot;
						}
						// у draw аннотаций ищем по path
						else if (oAnnot.IsInk() || oAnnot.IsLine())
						{
							let oPos	= oDrDoc.ConvertCoordsFromCursor2(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
							let X       = oPos.X;
        					let Y       = oPos.Y;

							if (oAnnot.hitInPath(X, Y))
								return oAnnot;
						}
						else
						{
							if (bGetHidden)
								return oAnnot;
							else
								return oAnnot.IsHidden() == false ? oAnnot : null;
						}
					}
				}
			}
			return null;
		};

		this.onMouseDown = function(e)
		{
			oThis.isFocusOnThumbnails = false;
			AscCommon.stopEvent(e);

			oThis.getPDFDoc().HideComments();
			
			var mouseButton = AscCommon.getMouseButton(e || {});
			if (mouseButton !== 0)
			{
				if (2 === mouseButton)
				{
					var posX = e.pageX || e.clientX;
					var posY = e.pageY || e.clientY;

					var x = posX - oThis.x;
					var y = posY - oThis.y;

					var isInSelection = false;
					if (oThis.overlay.m_oContext)
					{
						var pixX = AscCommon.AscBrowser.convertToRetinaValue(x, true);
						var pixY = AscCommon.AscBrowser.convertToRetinaValue(y, true);

						if (pixX >= 0 && pixY >= 0 && pixX < oThis.canvasOverlay.width && pixY < oThis.canvasOverlay.height)
						{
							var pixelOnOverlay = oThis.overlay.m_oContext.getImageData(pixX, pixY, 1, 1);
							if (Math.abs(pixelOnOverlay.data[0] - 51) < 10 &&
								Math.abs(pixelOnOverlay.data[1] - 102) < 10 &&
								Math.abs(pixelOnOverlay.data[2] - 204) < 10)
							{
								isInSelection = true;
							}
						}
					}

					if (isInSelection)
					{
						oThis.Api.sync_BeginCatchSelectedElements();
						oThis.Api.sync_ChangeLastSelectedElement(Asc.c_oAscTypeSelectElement.Text, undefined);
						oThis.Api.sync_EndCatchSelectedElements();

						oThis.Api.sync_ContextMenuCallback({
							Type: Asc.c_oAscContextMenuTypes.Common,
							X_abs: x,
							Y_abs: y
						});
					}
					else
					{
						oThis.Api.sync_BeginCatchSelectedElements();
						oThis.Api.sync_EndCatchSelectedElements();
						oThis.removeSelection();
						oThis.Api.sync_ContextMenuCallback({
							Type  : Asc.c_oAscContextMenuTypes.ChangeHdrFtr,
							X_abs : x,
							Y_abs : y
						});
					}
				}
				return;
			}

			oThis.isMouseDown = true;

			if (!oThis.file || !oThis.file.isValid())
				return;

			AscCommon.check_MouseDownEvent(e, true);
			AscCommon.global_mouseEvent.LockMouse();

			oThis.mouseDownCoords.X = AscCommon.global_mouseEvent.X;
			oThis.mouseDownCoords.Y = AscCommon.global_mouseEvent.Y;

			oThis.isMouseMoveBetweenDownUp = false;
			oThis.getPDFDoc().OnMouseDown(e);
		};

		this.onMouseDownEpsilon = function()
		{
			if (oThis.MouseHandObject)
			{
				if (oThis.mouseDownLinkObject || oThis.mouseDownField)
				{
					// если нажали на ссылке - то не зажимаем лапу
					oThis.setCursorType("pointer");
					return;
				}
				// режим лапы. просто начинаем режим Active - зажимаем лапу
				oThis.setCursorType(AscCommon.Cursors.Grabbing);
				oThis.MouseHandObject.X = oThis.mouseDownCoords.X;
				oThis.MouseHandObject.Y = oThis.mouseDownCoords.Y;
				oThis.MouseHandObject.Active = true;
				oThis.MouseHandObject.ScrollX = oThis.scrollX;
				oThis.MouseHandObject.ScrollY = oThis.scrollY;
				return;
			}

			var pageObjectLogic = this.getPageByCoords2(oThis.mouseDownCoords.X - oThis.x, oThis.mouseDownCoords.Y - oThis.y);
			this.file.onMouseDown(pageObjectLogic.index, pageObjectLogic.x, pageObjectLogic.y);

			if (-1 === this.timerScrollSelect && AscCommon.global_mouseEvent.IsLocked)
			{
				this.timerScrollSelect = setInterval(this.selectWheel, 20);
			}
		};

		this.onMouseUp = function(e)
		{
			oThis.isFocusOnThumbnails = false;
			//if (e && e.preventDefault)
			//	e.preventDefault();

			var mouseButton = AscCommon.getMouseButton(e || {});
			if (mouseButton !== 0)
				return;

			oThis.isMouseDown = false;

			if (!oThis.file || !oThis.file.isValid())
				return;

			if (!e)
			{
				// здесь - имитируем моус мув ---------------------------
				e   = {};
				e.pageX = AscCommon.global_mouseEvent.X;
				e.pageY = AscCommon.global_mouseEvent.Y;

				e.clientX = AscCommon.global_mouseEvent.X;
				e.clientY = AscCommon.global_mouseEvent.Y;

				e.altKey   = AscCommon.global_mouseEvent.AltKey;
				e.shiftKey = AscCommon.global_mouseEvent.ShiftKey;
				e.ctrlKey  = AscCommon.global_mouseEvent.CtrlKey;
				e.metaKey  = AscCommon.global_mouseEvent.CtrlKey;

				e.srcElement = AscCommon.global_mouseEvent.Sender;
				// ------------------------------------------------------

				AscCommon.Window_OnMouseUp(e);
			}

			AscCommon.check_MouseUpEvent(e);

			let oDoc = oThis.getPDFDoc();
			oDoc.OnMouseUp(e);

			if (false == oThis.Api.isInkDrawerOn())
			{
				if (!oThis.MouseHandObject && global_mouseEvent.ClickCount == 2 && !oDoc.mouseDownAnnot && !oDoc.mouseDownField)
				{
					var pageObjectLogic = oThis.getPageByCoords2(oThis.mouseDownCoords.X - oThis.x, oThis.mouseDownCoords.Y - oThis.y);
					oThis.file.selectWholeWord(pageObjectLogic.index, pageObjectLogic.x, pageObjectLogic.y);
				}
				else if (!oThis.MouseHandObject && global_mouseEvent.ClickCount == 3 && !oDoc.mouseDownAnnot && !oDoc.mouseDownField)
				{
					var pageObjectLogic = oThis.getPageByCoords2(oThis.mouseDownCoords.X - oThis.x, oThis.mouseDownCoords.Y - oThis.y);
					oThis.file.selectWholeRow(pageObjectLogic.index, pageObjectLogic.x, pageObjectLogic.y);
				}
			}

			if (oThis.mouseDownLinkObject)
			{
				// значит не уходили с ссылки
				// проверим - остались ли на ней
				var mouseUpLinkObject = oThis.getPageLinkByMouse();
				if (mouseUpLinkObject === oThis.mouseDownLinkObject)
				{
					oThis.navigateToLink(mouseUpLinkObject);
				}
			}

			// если было нажатие - то отжимаем
			if (oThis.isMouseMoveBetweenDownUp)
				oThis.file.onMouseUp();

			if (oThis.MouseHandObject)
				oThis.MouseHandObject.Active = false;
			oThis.isMouseMoveBetweenDownUp = false;
			oThis.mouseDownLinkObject = null;

			if (-1 !== oThis.timerScrollSelect)
			{
				clearInterval(oThis.timerScrollSelect);
				oThis.timerScrollSelect = -1;
			}
		};

		this.onMouseMove = function(e)
		{
			if (!oThis.file || !oThis.file.isValid())
				return;

			let oDoc = oThis.getPDFDoc();
			AscCommon.check_MouseMoveEvent(e);
			if (e && e.preventDefault)
				e.preventDefault();

			// если мышка нажата и еще не вышли за eps - то проверяем, модет вышли сейчас?
			// и, если вышли - то эмулируем
			if (oThis.isMouseDown && !oThis.isMouseMoveBetweenDownUp && !oDoc.mouseDownAnnot && !oDoc.mouseDownField && !oThis.Api.isInkDrawerOn())
			{
				var offX = Math.abs(oThis.mouseDownCoords.X - AscCommon.global_mouseEvent.X);
				var offY = Math.abs(oThis.mouseDownCoords.Y - AscCommon.global_mouseEvent.Y);

				if (offX > oThis.mouseMoveEpsilon || offY > oThis.mouseMoveEpsilon)
				{
					oThis.isMouseMoveBetweenDownUp = true;
					oThis.onMouseDownEpsilon();
				}
			}

			if (oThis.MouseHandObject)
			{
				if (oThis.MouseHandObject.Active && !oDoc.mouseDownAnnot && !oThis.Api.isInkDrawerOn())
				{
					// двигаем рукой
					oThis.setCursorType(AscCommon.Cursors.Grabbing);

					var scrollX = AscCommon.global_mouseEvent.X - oThis.MouseHandObject.X;
					var scrollY = AscCommon.global_mouseEvent.Y - oThis.MouseHandObject.Y;

					if (0 != scrollX && oThis.isVisibleHorScroll)
					{
						var pos = oThis.MouseHandObject.ScrollX - scrollX;
						if (pos < 0) pos = 0;
						if (pos > oThis.scrollMaxX) pos = oThis.scrollMaxX;
						oThis.m_oScrollHorApi.scrollToX(pos);
					}
					if (0 != scrollY)
					{
						var pos = oThis.MouseHandObject.ScrollY - scrollY;
						if (pos < 0) pos = 0;
						if (pos > oThis.scrollMaxY) pos = oThis.scrollMaxY;
						oThis.m_oScrollVerApi.scrollToY(pos);
					}

					return;
				}
				else
				{
					if (false == editor.isEmbedVersion)
						oDoc.OnMouseMove(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y, e);
				}
				return;
			}
			else
			{
				if (oThis.mouseDownLinkObject)
				{
					// селект начат на ссылке. смотрим, нужно ли начать реально селект
					if (oThis.isMouseMoveBetweenDownUp)
					{
						// вышли за eps
						oThis.mouseDownLinkObject = null;
						oThis.setCursorType("default");
					}
					else
					{
						oThis.setCursorType("pointer");
					}
				}

				if (oThis.isMouseDown)
				{
					if (oThis.isMouseMoveBetweenDownUp && !oDoc.activeForm && (!oDoc.mouseDownAnnot || (oDoc.mouseDownAnnot && oDoc.mouseDownAnnot.IsTextMarkup() == true)) && !oThis.Api.isInkDrawerOn())
					{
						// нажатая мышка - курсор всегда default (так как за eps вышли)
						oThis.setCursorType("default");

						var pageObjectLogic = oThis.getPageByCoords2(AscCommon.global_mouseEvent.X - oThis.x, AscCommon.global_mouseEvent.Y - oThis.y);
						oThis.file.onMouseMove(pageObjectLogic.index, pageObjectLogic.x, pageObjectLogic.y);
					}
					else
					{
						if (false == editor.isEmbedVersion)
							oDoc.OnMouseMove(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y, e);
					}
				}
				else
				{
					oThis.getPDFDoc().OnMouseMove(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y, e);
				}
			}
			return false;
		};

		this.onMouseWhell = function(e)
		{
			if (!oThis.file || !oThis.file.isValid())
				return;

			if (oThis.MouseHandObject && oThis.MouseHandObject.IsActive)
				return;

			var _ctrl = false;
			if (e.metaKey !== undefined)
				_ctrl = e.ctrlKey || e.metaKey;
			else
				_ctrl = e.ctrlKey;

			AscCommon.stopEvent(e);

			if (true === _ctrl)
			{
				return false;
			}

			var delta  = 0;
			var deltaX = 0;
			var deltaY = 0;

			if (undefined != e.wheelDelta && e.wheelDelta != 0)
			{
				//delta = (e.wheelDelta > 0) ? -45 : 45;
				delta = -45 * e.wheelDelta / 120;
			}
			else if (undefined != e.detail && e.detail != 0)
			{
				//delta = (e.detail > 0) ? 45 : -45;
				delta = 45 * e.detail / 3;
			}

			// New school multidimensional scroll (touchpads) deltas
			deltaY = delta;

			if (oThis.isVisibleHorScroll)
			{
				if (e.axis !== undefined && e.axis === e.HORIZONTAL_AXIS)
				{
					deltaY = 0;
					deltaX = delta;
				}

				// Webkit
				if (undefined !== e.wheelDeltaY && 0 !== e.wheelDeltaY)
				{
					//deltaY = (e.wheelDeltaY > 0) ? -45 : 45;
					deltaY = -45 * e.wheelDeltaY / 120;
				}
				if (undefined !== e.wheelDeltaX && 0 !== e.wheelDeltaX)
				{
					//deltaX = (e.wheelDeltaX > 0) ? -45 : 45;
					deltaX = -45 * e.wheelDeltaX / 120;
				}
			}

			deltaX >>= 0;
			deltaY >>= 0;

			if (0 != deltaX)
				oThis.m_oScrollHorApi.scrollBy(deltaX, 0, false);
			else if (0 != deltaY)
				oThis.m_oScrollVerApi.scrollBy(0, deltaY, false);

			// здесь - имитируем моус мув ---------------------------
			var _e   = {};
			_e.pageX = AscCommon.global_mouseEvent.X;
			_e.pageY = AscCommon.global_mouseEvent.Y;

			_e.clientX = AscCommon.global_mouseEvent.X;
			_e.clientY = AscCommon.global_mouseEvent.Y;

			_e.altKey   = AscCommon.global_mouseEvent.AltKey;
			_e.shiftKey = AscCommon.global_mouseEvent.ShiftKey;
			_e.ctrlKey  = AscCommon.global_mouseEvent.CtrlKey;
			_e.metaKey  = AscCommon.global_mouseEvent.CtrlKey;

			_e.srcElement = AscCommon.global_mouseEvent.Sender;

			oThis.onMouseMove(_e);
			// ------------------------------------------------------

			return false;
		};

		this.selectWheel = function()
		{
			if (!oThis.file || !oThis.file.isValid())
				return;

			if (oThis.MouseHandObject)
				return;

			var positionMinY = oThis.y;
			var positionMaxY = oThis.y + oThis.height;

			var scrollYVal = 0;
			if (AscCommon.global_mouseEvent.Y < positionMinY)
			{
				var delta = 30;
				if (20 > (positionMinY - AscCommon.global_mouseEvent.Y))
					delta = 10;

				scrollYVal = -delta;
			}
			else if (AscCommon.global_mouseEvent.Y > positionMaxY)
			{
				var delta = 30;
				if (20 > (AscCommon.global_mouseEvent.Y - positionMaxY))
					delta = 10;

				scrollYVal = delta;
			}

			var scrollXVal = 0;
			if (oThis.isVisibleHorScroll)
			{
				var positionMinX = oThis.x;
				var positionMaxX = oThis.x + oThis.width;

				if (AscCommon.global_mouseEvent.X < positionMinX)
				{
					var delta = 30;
					if (20 > (positionMinX - AscCommon.global_mouseEvent.X))
						delta = 10;

					scrollXVal = -delta;
				}
				else if (AscCommon.global_mouseEvent.X > positionMaxX)
				{
					var delta = 30;
					if (20 > (AscCommon.global_mouseEvent.X - positionMaxX))
						delta = 10;

					scrollXVal = delta;
				}
			}

			if (0 != scrollYVal)
				oThis.m_oScrollVerApi.scrollByY(scrollYVal, false);
			if (0 != scrollXVal)
				oThis.m_oScrollHorApi.scrollByX(scrollXVal, false);

			if (scrollXVal != 0 || scrollYVal != 0)
			{
				// здесь - имитируем моус мув ---------------------------
				var _e   = {};
				_e.pageX = AscCommon.global_mouseEvent.X;
				_e.pageY = AscCommon.global_mouseEvent.Y;

				_e.clientX = AscCommon.global_mouseEvent.X;
				_e.clientY = AscCommon.global_mouseEvent.Y;

				_e.altKey   = AscCommon.global_mouseEvent.AltKey;
				_e.shiftKey = AscCommon.global_mouseEvent.ShiftKey;
				_e.ctrlKey  = AscCommon.global_mouseEvent.CtrlKey;
				_e.metaKey  = AscCommon.global_mouseEvent.CtrlKey;

				_e.srcElement = AscCommon.global_mouseEvent.Sender;

				oThis.onMouseMove(_e);
				// ------------------------------------------------------
			}
		};

		this.paint = function()
		{
			this.isRepaint = true;
		};
		
		this.getStructure = function()
		{
			if (!this.file || !this.file.isValid())
				return null;
			var res = this.file.structure();
			return res;
		};

		this.drawSearchPlaces = function(dKoefX, dKoefY, xDst, yDst, places)
		{
			var rPR = 1;//AscCommon.AscBrowser.retinaPixelRatio;
			var len = places.length;

			var ctx = this.overlay.m_oContext;

			for (var i = 0; i < len; i++)
			{
				var place = places[i];
				if (undefined === place.Ex)
				{
					var _x = (rPR * (xDst + dKoefX * place.X)) >> 0;
					var _y = (rPR * (yDst + dKoefY * place.Y)) >> 0;

					var _w = (rPR * (dKoefX * place.W)) >> 0;
					var _h = (rPR * (dKoefY * place.H)) >> 0;

					if (_x < this.overlay.min_x)
						this.overlay.min_x = _x;
					if ((_x + _w) > this.overlay.max_x)
						this.overlay.max_x = _x + _w;

					if (_y < this.overlay.min_y)
						this.overlay.min_y = _y;
					if ((_y + _h) > this.overlay.max_y)
						this.overlay.max_y = _y + _h;

					ctx.rect(_x, _y, _w, _h);
				}
				else
				{
					var _x1 = (rPR * (xDst + dKoefX * place.X)) >> 0;
					var _y1 = (rPR * (yDst + dKoefY * place.Y)) >> 0;

					var x2 = place.X + place.W * place.Ex;
					var y2 = place.Y + place.W * place.Ey;
					var _x2 = (rPR * (xDst + dKoefX * x2)) >> 0;
					var _y2 = (rPR * (yDst + dKoefY * y2)) >> 0;

					var x3 = x2 - place.H * place.Ey;
					var y3 = y2 + place.H * place.Ex;
					var _x3 = (rPR * (xDst + dKoefX * x3)) >> 0;
					var _y3 = (rPR * (yDst + dKoefY * y3)) >> 0;

					var x4 = place.X - place.H * place.Ey;
					var y4 = place.Y + place.H * place.Ex;
					var _x4 = (rPR * (xDst + dKoefX * x4)) >> 0;
					var _y4 = (rPR * (yDst + dKoefY * y4)) >> 0;

					this.overlay.CheckPoint(_x1, _y1);
					this.overlay.CheckPoint(_x2, _y2);
					this.overlay.CheckPoint(_x3, _y3);
					this.overlay.CheckPoint(_x4, _y4);

					ctx.moveTo(_x1, _y1);
					ctx.lineTo(_x2, _y2);
					ctx.lineTo(_x3, _y3);
					ctx.lineTo(_x4, _y4);
					ctx.lineTo(_x1, _y1);
				}
			}

			ctx.fill();
			ctx.beginPath();
		};

		this.drawSearchCur = function(pageIndex, places)
		{
			var pageCoords = this.pageDetector.pages[pageIndex - this.startVisiblePage];
			if (!pageCoords)
				return;

			var scale = this.file.pages[pageIndex].Dpi / 25.4;
			var dKoefX = scale * pageCoords.w / this.file.pages[pageIndex].W;
			var dKoefY = scale * pageCoords.h / this.file.pages[pageIndex].H;

			var ctx = this.overlay.m_oContext;
			ctx.fillStyle = "rgba(51,102,204,255)";

			this.drawSearchPlaces(dKoefX, dKoefY, pageCoords.x, pageCoords.y, places);

			ctx.fill();
			ctx.beginPath();
		};

		this.drawSearch = function (pageIndex, searchingObj)
		{
			var pageCoords = this.pageDetector.pages[pageIndex - this.startVisiblePage];
			if (!pageCoords)
				return;

			var scale = this.file.pages[pageIndex].Dpi / 25.4;
			var dKoefX = scale * pageCoords.w / this.file.pages[pageIndex].W;
			var dKoefY = scale * pageCoords.h / this.file.pages[pageIndex].H;

			for (var i = 0; i < searchingObj.length; i++)
			{
				this.drawSearchPlaces(dKoefX, dKoefY, pageCoords.x, pageCoords.y, searchingObj[i]);
			}
		};

		this.onUpdateOverlay = function()
		{
			this.overlay.Clear()

			if (!this.file)
				return;

			if (this.startVisiblePage < 0 || this.endVisiblePage < 0)
				return;

			// seletion
			var ctx = this.overlay.m_oContext;
			ctx.globalAlpha = 0.2;

			if (this.file.SearchResults.IsSearch)
			{
				if (this.file.SearchResults.Show)
				{
					ctx.globalAlpha = 0.5;
					ctx.fillStyle = "rgba(255,200,0,1)";
					ctx.beginPath();

					for (let i = this.startVisiblePage; i <= this.endVisiblePage; i++)
					{
						var searchingObj = this.file.SearchResults.Pages[i];
						if (0 != searchingObj.length)
							this.drawSearch(i, searchingObj);
					}

					ctx.fill();
					ctx.globalAlpha = 0.2;
				}
				ctx.beginPath();

				if (this.CurrentSearchNavi && this.file.SearchResults.Show)
				{
					var pageNum  = this.CurrentSearchNavi[0].PageNum;
					ctx.fillStyle = "rgba(51,102,204,255)";
					if (pageNum >= this.startVisiblePage && pageNum <= this.endVisiblePage)
					{
						this.drawSearchCur(pageNum, this.CurrentSearchNavi);
					}
				}
			}

			let oDoc = this.getPDFDoc();
			//if (!this.MouseHandObject)
			{
				ctx.fillStyle = "rgba(51,102,204,255)";
				ctx.beginPath();

				for (let i = this.startVisiblePage; i <= this.endVisiblePage; i++)
				{
					var pageCoords = this.pageDetector.pages[i - this.startVisiblePage];
					ctx.beginPath();
					this.file.drawSelection(i, this.overlay, pageCoords.x, pageCoords.y, pageCoords.w, pageCoords.h);
					ctx.fill();
				}

				if (this.DrawingObjects.needUpdateOverlay())
				{
					this.DrawingObjects.drawingDocument.AutoShapesTrack.PageIndex = -1;
					this.DrawingObjects.drawOnOverlay(this.DrawingObjects.drawingDocument.AutoShapesTrack);
					this.DrawingObjects.drawingDocument.AutoShapesTrack.CorrectOverlayBounds();
				}
				else if (oDoc.mouseDownAnnot)
				{
					if (oDoc.mouseDownAnnot.IsTextMarkup())
					{
						oDoc.mouseDownAnnot.DrawSelected(this.overlay);
					}
					else if (oDoc.mouseDownAnnot.IsInk() || oDoc.mouseDownAnnot.IsLine())
					{
						let nPage = oDoc.mouseDownAnnot.GetPage();
						this.DrawingObjects.drawingDocument.AutoShapesTrack.PageIndex = nPage;
						this.DrawingObjects.drawSelect(nPage);

						if (oDoc.mouseDownAnnot.IsLine())
							oDoc.mouseDownAnnot.DrawSelected(this.overlay)
					}
				}
			}
			if (oDoc.activeForm && oDoc.activeForm.content && oDoc.activeForm.content.IsSelectionUse() && oDoc.activeForm.content.IsSelectionEmpty() == false)
			{
				ctx.beginPath();
				oDoc.activeForm.content.DrawSelectionOnPage(0);
				ctx.globalAlpha = 0.2;
				ctx.fill();
			}
			
			ctx.globalAlpha = 1.0;
		};

		this._paint = function()
		{
			if (!this.file || !this.file.isValid() || !this.canvas)
				return;

			this.canvas.width = this.canvas.width;
			let ctx = this.canvas.getContext("2d");
			ctx.strokeStyle = AscCommon.GlobalSkin.PageOutline;
			let lineW = AscCommon.AscBrowser.retinaPixelRatio >> 0;
			ctx.lineWidth = lineW;

			let yPos = this.scrollY >> 0;
			let yMax = yPos + this.height;
			let xCenter = this.width >> 1;
			if (this.documentWidth > this.width)
			{
				xCenter = (this.documentWidth >> 1) - (this.scrollX) >> 0;
			}

			let lStartPage = -1;
			let lEndPage = -1; 
			
			let oDoc = this.getPDFDoc();
			let lPagesCount = this.drawingPages.length;
			for (let i = 0; i < lPagesCount; i++)
			{
				let page = this.drawingPages[i];
				let pageT = page.Y;
				let pageB = page.Y + page.H;
				
				if (yPos < pageB && yMax > pageT)
				{
					// страница на экране

					if (-1 == lStartPage)
						lStartPage = i;
					lEndPage = i;
				}
				else
				{
					// страница не видна - выкидываем из кэша
					if (page.Image)
					{
						if (this.file.cacheManager)
							this.file.cacheManager.unlock(page.Image);
						
						delete page.Image;
						delete page.ImageForms;
						delete page.ImageAnnots;
						oDoc.ClearCache(i);
					}
				}
			}

			this.pageDetector = new CCurrentPageDetector(this.canvas.width, this.canvas.height);

			let oDrDoc = oDoc.GetDrawingDocument();
			oDrDoc.m_lDrawingFirst = lStartPage;
			oDrDoc.m_lDrawingEnd = lEndPage;
			this.startVisiblePage = lStartPage;
			this.endVisiblePage = lEndPage;

			var isStretchPaint = this.Api.WordControl.NoneRepaintPages;
			if (this.isClearPages)
				isStretchPaint = false;

			for (let i = lStartPage; i <= lEndPage; i++)
			{
				// отрисовываем страницу
				let page = this.drawingPages[i];
				if (!page)
					break;

				let w = (page.W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
				let h = (page.H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

				let rotateAngle = this.getPageRotate(i);
				let natW = w;
				let natH = h;
				if (rotateAngle & 1)
				{
					natW = h;
					natH = w;
				}

				if (!isStretchPaint)
				{
					if (!this.file.cacheManager)
					{
						if (this.bCachedMarkupAnnnots)
						{
							if (this.pagesInfo.pages[i].needRedrawHighlights || this.isClearPages || (page.Image && ((page.Image.requestWidth !== natW) || (page.Image.requestHeight !== natH))))
								delete page.Image;
						}
						else
						{
							if (this.isClearPages || (page.Image && ((page.Image.requestWidth !== natW) || (page.Image.requestHeight !== natH))))
								delete page.Image;
						}
					}
					else
					{
						if (this.bCachedMarkupAnnnots)
						{
							if (this.pagesInfo.pages[i].needRedrawHighlights || this.isClearPages || (page.Image && ((page.Image.requestWidth < natW) || (page.Image.requestHeight < natH))))
							{
								if (this.file.cacheManager)
									this.file.cacheManager.unlock(page.Image);

								delete page.Image;
							}
						}
						else
						{
							if (this.isClearPages || (page.Image && ((page.Image.requestWidth < natW) || (page.Image.requestHeight < natH))))
							{
								if (this.file.cacheManager)
									this.file.cacheManager.unlock(page.Image);

								delete page.Image;
							}
						}
						
					}
				}

				if (!page.Image && !isStretchPaint)
				{
					page.Image = this.file.getPage(i, natW, natH, undefined, this.Api.isDarkMode ? 0x3A3A3A : 0xFFFFFF);
					if (this.bCachedMarkupAnnnots)
						this._paintMarkupAnnotsOnPage(i, page.Image.getContext("2d"));

					// нельзя кэшировать с вотермарком - так как есть поворот
					//if (this.Api.watermarkDraw)
					//	this.Api.watermarkDraw.Draw(page.Image.getContext("2d"), w, h);
				}

				let x = ((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - (w >> 1);
				let y = ((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

				if (page.Image)
				{
					if (0 === rotateAngle)
					{
						ctx.drawImage(page.Image, 0, 0, page.Image.width, page.Image.height, x, y, w, h);
					}
					else
					{
						let cx = x + 0.5 * w;
						let cy = y + 0.5 * h;

						ctx.save();
						ctx.translate(cx, cy);
						ctx.rotate(rotateAngle * Math.PI / 2);
						ctx.drawImage(page.Image, -0.5 * natW, -0.5 * natH, natW, natH);
						ctx.restore();
					}
					this.pagesInfo.setPainted(i);
				}
				else
				{
					ctx.fillStyle = "#FFFFFF";
					ctx.fillRect(x, y, w, h);
				}
				ctx.strokeRect(x + lineW / 2, y + lineW / 2, w - lineW, h - lineW);

				if (this.Api.watermarkDraw)
					this.Api.watermarkDraw.Draw(ctx, x, y, w, h);

				this.pageDetector.addPage(i, x, y, w, h);

				if (false == this.bCachedMarkupAnnnots)
					this._paintMarkupAnnotsOnPage(i, ctx);
			}
			
			this.isClearPages = false;
			this.updateCurrentPage(this.pageDetector.getCurrentPage(this.currentPage));
			
			// выход из формы если вышли со страницы, где находится активная форма.
			if (oDoc.activeForm && this.pageDetector.pages.map(function(item) {
				return item.num;
			}).includes(oDoc.activeForm.GetPage()) == false) {
				if (oDoc.activeForm.IsChanged() == false) {
					oDoc.activeForm.SetDrawFromStream(true);
				}
				oDoc.activeForm.SetDrawHighlight(true);
				oDoc.activeForm = null;
			}

			this._paintAnnots();
			this._paintForms();
			this._paintFormsHighlight();
			this._paintComboboxesMarkers();
			oDoc.UpdateUndoRedo();
			oDoc.UpdateCommentPos();
		};
		this.Get_PageLimits = function() {
			let W = this.width;
			let H = this.height;
			let scaleCoef = this.zoom * AscCommon.AscBrowser.retinaPixelRatio;

			return {
				X: 0,
				Y: 0,
				XLimit: W * g_dKoef_pix_to_mm / scaleCoef,
				YLimit: H * g_dKoef_pix_to_mm / scaleCoef
			}
		};
		this.SelectNextField = function()
		{
			this.doc.SelectNextField();
		};
		this.SelectPrevField = function()
		{
			this.doc.SelectPrevField();
		};
		
		this.checkPagesLinks = function()
		{
			if (this.startVisiblePage < 0 || this.endVisiblePage < 0)
				return false;

			for (var i = this.startVisiblePage; i <= this.endVisiblePage; i++)
			{
				var page = this.pagesInfo.pages[i];
				if (page.isPainted && null === page.links)
				{
					page.links = this.file.getLinks(i);
					return true;
				}
			}

			return false;
		};

		this.checkPagesText = function()
		{
			if (this.startVisiblePage < 0 || this.endVisiblePage < 0)
				return false;

			if (this.isFullText)
				return;

			var pagesCount = this.file.pages.length;
			var isCommands = false;
			for (var i = this.startVisiblePage; i <= this.endVisiblePage; i++)
			{
				if (null == this.file.pages[i].text)
				{
					this.file.pages[i].text = this.file.getText(i);
					isCommands = true;
				}
			}

			if (!isCommands)
			{
				while (this.pagesInfo.countTextPages < pagesCount)
				{
					// мы могли уже получить команды, так как видимые страницы в приоритете
					if (null != this.file.pages[this.pagesInfo.countTextPages].text)
					{
						this.pagesInfo.countTextPages++;
						continue;
					}

					this.file.pages[this.pagesInfo.countTextPages].text = this.file.getText(this.pagesInfo.countTextPages);
					if (null != this.file.pages[this.pagesInfo.countTextPages].text)
					{
						this.pagesInfo.countTextPages++;
						isCommands = true;
					}

					break;
				}
			}

			if (this.pagesInfo.countTextPages === pagesCount)
			{
				this.file.destroyText();

				this.isFullText = true;
				if (this.isFullTextMessage)
					this.unshowTextMessage();

				if (this.statistics.process)
				{
					this.endStatistics();
					this.Api.sync_GetDocInfoEndCallback();
				}
			}

			return isCommands;
		};

		this.getPageByCoords = function(xInp, yInp)
		{
			if (this.startVisiblePage < 0 || this.endVisiblePage < 0)
				return null;

			var x = xInp * AscCommon.AscBrowser.retinaPixelRatio;
			var y = yInp * AscCommon.AscBrowser.retinaPixelRatio;
			for (var i = this.startVisiblePage; i <= this.endVisiblePage; i++)
			{
				var pageCoords = this.pageDetector.pages[i - this.startVisiblePage];
				if (!pageCoords)
					continue;
				if (x >= pageCoords.x && x <= (pageCoords.x + pageCoords.w) &&
					y >= pageCoords.y && y <= (pageCoords.y + pageCoords.h))
				{
					return {
						index : i,
						x : this.file.pages[i].W * (x - pageCoords.x) / pageCoords.w,
						y : this.file.pages[i].H * (y - pageCoords.y) / pageCoords.h
					};
				}
			}
			return null;
		};

		this.getPageByCoords2 = function(x, y)
		{
			if (this.startVisiblePage < 0 || this.endVisiblePage < 0)
				return null;

			var pageCoords = null;
			var pageIndex = 0;
			for (pageIndex = this.startVisiblePage; pageIndex <= this.endVisiblePage; pageIndex++)
			{
				pageCoords = this.pageDetector.pages[pageIndex - this.startVisiblePage];
				if ((pageCoords.y + pageCoords.h) > y)
					break;
			}
			if (pageIndex > this.endVisiblePage)
				pageIndex = this.endVisiblePage;

			if (!pageCoords)
				pageCoords = {x:0, y:0, w:1, h:1};

			var pixToMM = (25.4 / this.file.pages[pageIndex].Dpi);
			return {
				index : pageIndex,
				x : this.file.pages[pageIndex].W * pixToMM * (x * AscCommon.AscBrowser.retinaPixelRatio - pageCoords.x) / pageCoords.w,
				y : this.file.pages[pageIndex].H * pixToMM * (y * AscCommon.AscBrowser.retinaPixelRatio - pageCoords.y) / pageCoords.h
			};
		};
		this.getPageByCoords3 = function(xInp, yInp)
		{
			if (this.startVisiblePage < 0 || this.endVisiblePage < 0)
				return null;
			
			var x = xInp * AscCommon.AscBrowser.retinaPixelRatio;
			var y = yInp * AscCommon.AscBrowser.retinaPixelRatio;
			
			let pageIndex = this.endVisiblePage;
			let page = this.pageDetector.pages[pageIndex - this.startVisiblePage];
			for (var i = this.startVisiblePage; i <= this.endVisiblePage; i++)
			{
				let _page = this.pageDetector.pages[i - this.startVisiblePage];
				if (!_page)
					continue;
				
				if (_page.y + _page.h + this.betweenPages * AscCommon.AscBrowser.retinaPixelRatio > y)
				{
					pageIndex = i;
					page = _page;
					break;
				}
			}
			
			if (!page)
				return null;
			
			return {
				index : pageIndex,
				x : this.file.pages[pageIndex].W * (x - page.x) / page.w,
				y : this.file.pages[pageIndex].H * (y - page.y) / page.h
			};
		};

		this.getViewportPosition = function()
		{
			return {
				scrollY : this.scrollY >> 0,
				centerX : (this.documentWidth > this.width) ? ((this.documentWidth >> 1) - (this.scrollX) >> 0) : (this.width >> 1)
			};
		};

		this.ConvertCoordsToCursor = function(x, y, pageIndex)
		{
			var dKoef = (this.zoom * g_dKoef_mm_to_pix);
			var rPR = 1;//AscCommon.AscBrowser.retinaPixelRatio;
			let yPos = this.scrollY >> 0;
			let xCenter = this.width >> 1;
			if (this.documentWidth > this.width)
			{
				xCenter = (this.documentWidth >> 1) - (this.scrollX) >> 0;
			}

			let page = this.drawingPages[pageIndex];

			let _w = (page.W * rPR) >> 0;
			let _h = (page.H * rPR) >> 0;
			let _x = ( (xCenter * rPR) >> 0) - (_w >> 1);
			let _y = ( (page.Y - yPos) * rPR) >> 0;

			var x_pix = (_x + x * dKoef) >> 0;
			var y_pix = (_y + y * dKoef) >> 0;
			var w_pix = (_w * dKoef) >> 0;
			var h_pix = (_h * dKoef) >> 0

			return ( {x : x_pix, y : y_pix, w : w_pix, h: h_pix} );
		};

		this.getPageLikeDetector = function(pageIndex)
		{
			let curPosition = this.getViewportPosition();

			let page = this.drawingPages[pageIndex];
			let w = (page.W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
			let h = (page.H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

			return {
				x : ((curPosition.centerX * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - (w >> 1),
				y : ((page.Y - curPosition.scrollY) * AscCommon.AscBrowser.retinaPixelRatio) >> 0,
				w : w,
				h : h
			};
		};

		this.Copy = function(_text_format)
		{
			return this.file.copy(_text_format);
		};
		this.selectAll = function()
		{
			return this.file.selectAll();
		};
		this.removeSelection = function()
		{
			var pageObjectLogic = this.getPageByCoords2(AscCommon.global_mouseEvent.X - this.x, AscCommon.global_mouseEvent.Y - this.y);
			this.file.onMouseDown(pageObjectLogic.index, pageObjectLogic.x, pageObjectLogic.y);
			this.file.onMouseUp(pageObjectLogic.index, pageObjectLogic.x, pageObjectLogic.y);
		};

		this.isCanCopy = function()
		{
			// TODO: нужно прерываться после первого же символа
			var text_format = { Text : "" };
			this.Copy(text_format);
			text_format.Text = text_format.Text.replace(new RegExp("\n", 'g'), "");
			return (text_format.Text === "") ? false : true;
		};

		this.findText = function(text, isMachingCase, isWholeWords, isNext, callback)
		{
			if (!this.isFullText)
			{
				this.fullTextMessageCallbackArgs = [text, isMachingCase, isWholeWords, isNext, callback];
				this.fullTextMessageCallback = function() {
					this.file.findText(this.fullTextMessageCallbackArgs[0], this.fullTextMessageCallbackArgs[1], this.fullTextMessageCallbackArgs[2], this.fullTextMessageCallbackArgs[3]);
					this.onUpdateOverlay();

					if (this.fullTextMessageCallbackArgs[4])
						this.fullTextMessageCallbackArgs[4].call(this.Api, this.SearchResults.Current, this.SearchResults.Count);
				};
				this.showTextMessage();
				return true; // async
			}

			this.file.findText(text, isMachingCase, isWholeWords, isNext);
			this.onUpdateOverlay();
			return false;
		};

		this.ToSearchResult = function()
		{
			var naviG = this.CurrentSearchNavi;

			var navi = naviG[0];
			var x    = navi.X;
			var y    = navi.Y;

			if (navi.Transform)
			{
				var xx = navi.Transform.TransformPointX(x, y);
				var yy = navi.Transform.TransformPointY(x, y);

				x = xx;
				y = yy;
			}

			var drawingPage = this.drawingPages[navi.PageNum];
			if (!drawingPage)
				return;

			var offsetBorder = 30;

			var scale = this.file.pages[navi.PageNum].Dpi / 25.4;
			var dKoefX = scale * drawingPage.W / this.file.pages[navi.PageNum].W;
			var dKoefY = scale * drawingPage.H / this.file.pages[navi.PageNum].H;

			var nX = drawingPage.X + dKoefX * x;
			var nY = drawingPage.Y + dKoefY * y;
			var nY2 = drawingPage.Y + dKoefY * (y + navi.H);

			if (this.m_oScrollHorApi)
				nX -= this.m_oScrollHorApi.scrollHCurrentX;
			nY -= this.m_oScrollVerApi.scrollVCurrentY;
			nY2 -= this.m_oScrollVerApi.scrollVCurrentY;

			var boxX = 0;
			var boxY = 0;
			var boxR = this.width;
			var boxB = this.height;

			var nValueScrollHor = 0;
			if (nX < boxX)
			{
				nValueScrollHor = nX - boxX - offsetBorder;
			}
			if (nX > boxR)
			{
				nValueScrollHor = nX - boxR + offsetBorder;
			}

			var nValueScrollVer = 0;
			if (nY < boxY)
			{
				nValueScrollVer = nY - boxY - offsetBorder;
			}
			if (nY2 > boxB)
			{
				nValueScrollVer = nY2 - boxB + offsetBorder;
			}

			if (0 !== nValueScrollHor)
			{
				this.m_bIsUpdateTargetNoAttack = true;
				this.m_oScrollHorApi.scrollByX(nValueScrollHor);
			}
			if (0 !== nValueScrollVer)
			{
				this.m_oScrollVerApi.scrollByY(nValueScrollVer);
			}
		};

		this.SelectSearchElement = function(elmId)
		{
			var nSearchedId = 0, nPage;
			var nMatchesCount = 0;
			for (nPage = 0; nPage < this.SearchResults.Pages.length; nPage++)
			{
				for (var nMatch = 0; nMatch < this.SearchResults.Pages[nPage].length; nMatch++)
				{
					nMatchesCount++;

					if (nMatchesCount - 1 == elmId)
					{
						nSearchedId = nMatch;
						break;
					}
				}
				if (nMatchesCount - 1 == elmId)
				{
					nSearchedId = nMatch;
					break;
				}
			}

			this.CurrentSearchNavi = this.SearchResults.Pages[nPage][nSearchedId];
			this.SearchResults.CurrentPage = nPage;
			this.SearchResults.Current = nSearchedId;
			this.SearchResults.CurMatchIdx = elmId;
            this.ToSearchResult();
			this.onUpdateOverlay();
			this.Api.sync_setSearchCurrent(elmId, this.SearchResults.Count);
		};
		
		this.OnKeyDown = function(e)
		{
			var bRetValue = false;
			let oDoc = this.getPDFDoc();

			if (e.KeyCode === 8) // BackSpace
			{
				if (oDoc.activeForm && oDoc.activeForm.IsEditable())
				{
					oDoc.activeForm.Remove(-1, e.CtrlKey == true);
					if (oDoc.activeForm.IsNeedRecalc())
						this._paint();

					this.onUpdateOverlay();
					// сбрасываем счетчик до появления курсора
					if (true !== e.ShiftKey)
						oThis.Api.WordControl.m_oDrawingDocument.TargetStart();
					// Чтобы при зажатой клавише курсор не пропадал
					oThis.Api.WordControl.m_oDrawingDocument.showTarget(true);
					
					bRetValue = true;
				}
			}
			else if (e.KeyCode === 9) // Tab
			{
				window.event.preventDefault();
				if (true == e.ShiftKey)
					this.SelectPrevField();
				else
					this.SelectNextField();
			}
			else if (e.KeyCode === 13 && e.ShiftKey == false) // Enter
			{
				window.event.stopPropagation();
				if (this.doc.activeForm && this.doc.activeForm.IsEditable() && this.doc.activeForm.IsMultiline && this.doc.activeForm.IsMultiline())
					this.Api.asc_enterText([13]);
				else
					this.doc.EnterDownActiveField(oDoc.activeForm);
			}
			else if (e.KeyCode === 13 && e.ShiftKey == true) // Enter
			{
				window.event.stopPropagation();
				if (this.doc.activeForm && this.doc.activeForm.IsEditable() && this.doc.activeForm.IsMultiline && this.doc.activeForm.IsMultiline())
					this.Api.asc_enterText([13]);
				else
					this.doc.EnterDownActiveField(oDoc.activeForm);
			}
			else if (e.KeyCode === 27) // Esc
			{
				if (this.Api.isInkDrawerOn())
				{
					this.Api.stopInkDrawer();
				}
				else if (this.Api.isMarkerFormat)
				{
					this.Api.sync_MarkerFormatCallback(false);
				}
				
				oDoc.EscapeForm();
				
				editor.sync_HideComment();
			}
			else if (e.KeyCode === 32) // Space
			{
				if (oDoc.activeForm)
				{
					
				}
				// to do включить checkbox/radio
			}
			else if ( e.KeyCode == 33 ) // PgUp
			{
				if (e.AltKey == true)
				{
					var nextPage = -1;
					if (this.thumbnails)
						nextPage = this.currentPage - this.thumbnails.countPagesInBlock;
					if (nextPage < 0)
						nextPage = this.currentPage - 1;

					if (nextPage >= 0)
						this.navigateToPage(nextPage);
				}
				else {
					this.m_oScrollVerApi.scrollByY(-this.height, false);
					this.timerSync();
				}
			}
			else if ( e.KeyCode == 34 ) // PgDn
			{
				if (e.AltKey == true)
				{
					var pagesCount = this.getPagesCount();
					var nextPage = pagesCount;
					if (this.thumbnails)
					{
						nextPage = this.currentPage + this.thumbnails.countPagesInBlock;
						if (nextPage >= pagesCount)
							nextPage = pagesCount - 1;
					}
					if (nextPage >= pagesCount)
						nextPage = this.currentPage + 1;

					if (nextPage < pagesCount)
						this.navigateToPage(nextPage);
				}
				else {
					this.m_oScrollVerApi.scrollByY(this.height, false);
					this.timerSync();
				}
			}
			else if ( e.KeyCode == 35 ) // End
			{
				if ( true === e.CtrlKey ) // Ctrl + End
				{
					this.m_oScrollVerApi.scrollToY(this.m_oScrollVerApi.maxScrollY, false);
				}
				this.timerSync();
				bRetValue = true;
			}
			else if ( e.KeyCode == 36 ) // клавиша Home
			{
				if ( true === e.CtrlKey ) // Ctrl + Home
				{
					this.m_oScrollVerApi.scrollToY(0, false);
				}
				this.timerSync();
				bRetValue = true;
			}
			else if ( e.KeyCode == 37 ) // Left Arrow
			{
				if (oDoc.activeForm && (oDoc.activeForm.IsEditable() || oDoc.activeForm.GetType() == AscPDF.FIELD_TYPES.combobox))
				{
					// сбрасываем счетчик до появления курсора
					if (true !== e.ShiftKey)
						oThis.Api.WordControl.m_oDrawingDocument.TargetStart();
					// Чтобы при зажатой клавише курсор не пропадал
					oThis.Api.WordControl.m_oDrawingDocument.showTarget(true);

					let oFieldBounds = oDoc.activeForm.getFormRelRect();
					let oCurPos = oDoc.activeForm.MoveCursorLeft(true === e.ShiftKey, true === e.CtrlKey);
					if (oDoc.activeForm.content.IsSelectionUse())
						this.Api.WordControl.m_oDrawingDocument.TargetEnd();
						
					let nCursorH = g_oTextMeasurer.GetHeight();
					if ((oCurPos.X < oFieldBounds.X || oCurPos.Y - nCursorH * 0.75 < oFieldBounds.Y) && oDoc.activeForm._doNotScroll != true)
					{
						oDoc.activeForm.AddToRedraw();
						this._paint();
						if (oDoc.activeForm.UpdateScroll)
							oDoc.activeForm.UpdateScroll(true);
					}
					
					this.onUpdateOverlay();
				}
				else if (!this.isFocusOnThumbnails && this.isVisibleHorScroll)
				{
					this.m_oScrollHorApi.scrollByX(-40);
				}
				else if (this.isFocusOnThumbnails)
				{
					if (this.currentPage > 0)
						this.navigateToPage(this.currentPage - 1);
				}
				bRetValue = true;
			}
			else if ( e.KeyCode == 38 ) // Top Arrow
			{
				if (oDoc.activeForm && !oDoc.activeForm.IsNeedDrawHighlight())
				{
					switch (oDoc.activeForm.GetType())
					{
						case AscPDF.FIELD_TYPES.listbox:
							oDoc.activeForm.MoveSelectUp();
							break;
						case AscPDF.FIELD_TYPES.text:
							// сбрасываем счетчик до появления курсора
							if (true !== e.ShiftKey)
								oThis.Api.WordControl.m_oDrawingDocument.TargetStart();
							// Чтобы при зажатой клавише курсор не пропадал
							oThis.Api.WordControl.m_oDrawingDocument.showTarget(true);

							let oFieldBounds = oDoc.activeForm.getFormRelRect();
							let oCurPos = oDoc.activeForm.MoveCursorUp(true === e.ShiftKey, true === e.CtrlKey);
							if (oDoc.activeForm.content.IsSelectionUse())
								this.Api.WordControl.m_oDrawingDocument.TargetEnd();

							let nCursorH = g_oTextMeasurer.GetHeight();
							if (oCurPos.Y - nCursorH * 0.75 < oFieldBounds.Y && oDoc.activeForm._doNotScroll != true)
							{
								oDoc.activeForm.AddToRedraw();
								this._paint();
								if (oDoc.activeForm.UpdateScroll)
									oDoc.activeForm.UpdateScroll(true);
							}
														
							this.onUpdateOverlay();
							break;
					}
				}
				else if (!this.isFocusOnThumbnails)
				{
					this.m_oScrollVerApi.scrollByY(-40);
				}
				bRetValue = true;
			}
			else if ( e.KeyCode == 39 ) // Right Arrow
			{	
				if (oDoc.activeForm && (oDoc.activeForm.IsEditable() || oDoc.activeForm.GetType() == AscPDF.FIELD_TYPES.combobox))
				{
					// сбрасываем счетчик до появления курсора
					if (true !== e.ShiftKey)
						oThis.Api.WordControl.m_oDrawingDocument.TargetStart();
					// Чтобы при зажатой клавише курсор не пропадал
					oThis.Api.WordControl.m_oDrawingDocument.showTarget(true);

					let oFieldBounds = oDoc.activeForm.getFormRelRect();
					let oCurPos = oDoc.activeForm.MoveCursorRight(true === e.ShiftKey, true === e.CtrlKey);
					
					if (oDoc.activeForm.content.IsSelectionUse())
						this.Api.WordControl.m_oDrawingDocument.TargetEnd();

					if ((oCurPos.X > oFieldBounds.X + oFieldBounds.W || oCurPos.Y > oFieldBounds.Y + oFieldBounds.H) && oDoc.activeForm._doNotScroll != true)
					{
						oDoc.activeForm.AddToRedraw();
						this._paint();
						if (oDoc.activeForm.UpdateScroll)
							oDoc.activeForm.UpdateScroll(true);
					}

					this.onUpdateOverlay();
				}
				else if (!this.isFocusOnThumbnails && this.isVisibleHorScroll)
				{
					this.m_oScrollHorApi.scrollByX(40);
				}
				else if (this.isFocusOnThumbnails)
				{
					if (this.currentPage < (this.getPagesCount() - 1))
						this.navigateToPage(this.currentPage + 1);
				}
				bRetValue = true;
			}
			else if ( e.KeyCode == 40 ) // Bottom Arrow
			{
				if (oDoc.activeForm && !oDoc.activeForm.IsNeedDrawHighlight())
				{
					switch (oDoc.activeForm.GetType())
					{
						case AscPDF.FIELD_TYPES.listbox:
							oDoc.activeForm.MoveSelectDown();
							break;
						case AscPDF.FIELD_TYPES.text:
							// сбрасываем счетчик до появления курсора
							if (true !== e.ShiftKey)
								oThis.Api.WordControl.m_oDrawingDocument.TargetStart();
							// Чтобы при зажатой клавише курсор не пропадал
							oThis.Api.WordControl.m_oDrawingDocument.showTarget(true);

							let oFieldBounds = oDoc.activeForm.getFormRelRect();
							let oCurPos = oDoc.activeForm.MoveCursorDown(true === e.ShiftKey, true === e.CtrlKey);
							if (oDoc.activeForm.content.IsSelectionUse())
								this.Api.WordControl.m_oDrawingDocument.TargetEnd();
								
							if (oCurPos.Y > oFieldBounds.Y + oFieldBounds.H && oDoc.activeForm._doNotScroll != true)
							{
								oDoc.activeForm.AddToRedraw();
								this._paint();
								if (oDoc.activeForm.UpdateScroll)
									oDoc.activeForm.UpdateScroll(true);
							}

							this.onUpdateOverlay();
							break;
					}
					
				}
				else if (!this.isFocusOnThumbnails)
				{
					this.m_oScrollVerApi.scrollByY(40);
				}
				bRetValue = true;
			}
			else if (e.KeyCode === 46) // Delete
			{
				let oDoc = this.getPDFDoc();

				if (oDoc.activeForm && oDoc.activeForm.IsEditable())
				{
					oDoc.activeForm.Remove(1, e.CtrlKey == true);
					if (oDoc.activeForm._needRecalc)
						this._paint();

					this.onUpdateOverlay();
					// сбрасываем счетчик до появления курсора
					if (true !== e.ShiftKey)
						oThis.Api.WordControl.m_oDrawingDocument.TargetStart();
					// Чтобы при зажатой клавише курсор не пропадал
					oThis.Api.WordControl.m_oDrawingDocument.showTarget(true);
					bRetValue = true;
				}
				else if (oDoc.mouseDownAnnot && this.isMouseDown == false) {
					oDoc.RemoveAnnot(oDoc.mouseDownAnnot.GetId());
				}
			}
			else if ( e.KeyCode == 65 && true === e.CtrlKey ) // Ctrl + A
			{
				if (oDoc.activeForm && [AscPDF.FIELD_TYPES.text, AscPDF.FIELD_TYPES.combobox].includes(oDoc.activeForm.GetType()))
				{
					if (oDoc.activeForm.IsNeedDrawHighlight())
						return;
					oDoc.activeForm.content.SelectAll();
					if (oDoc.activeForm.content.IsSelectionUse())
						this.Api.WordControl.m_oDrawingDocument.TargetEnd();
					
					this.onUpdateOverlay();
					bRetValue = true;
				}
				else
				{
					bRetValue = true;
					if (!this.isFullTextMessage) {
						if (!this.isFullText)
						{
							this.fullTextMessageCallbackArgs = [];
							this.fullTextMessageCallback = function() {
								this.file.selectAll();
							};
							this.showTextMessage();
						}
						else
						{
							this.file.selectAll();
						}
					}
				} 
			}
			else if ( e.KeyCode == 80 && true === e.CtrlKey ) // Ctrl + P + ...
			{
				this.Api.onPrint();
				bRetValue = true;
			}
			else if ( e.KeyCode == 83 && true === e.CtrlKey ) // Ctrl + S + ...
			{
				// nothing
				bRetValue = true;
			}
			else if ( e.KeyCode == 89 && true === e.CtrlKey ) // Ctrl + Y
			{
				this.doc.DoRedo();
				bRetValue = true;
			}
			else if ( e.KeyCode == 90 && true === e.CtrlKey ) // Ctrl + Z
			{
				this.doc.DoUndo();
				bRetValue = true;
			}

			oDoc.UpdateCopyCutState();
			return bRetValue;
		};
		this.showTextMessage = function()
		{
			if (this.isFullTextMessage)
				return;

			this.isFullTextMessage = true;
			this.Api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Waiting);
		};

		this.unshowTextMessage = function()
		{
			this.isFullTextMessage = false;
			this.Api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Waiting);

			if (this.fullTextMessageCallback)
			{
				this.fullTextMessageCallback.apply(this, this.fullTextMessageCallbackArgs);
				this.fullTextMessageCallback = null;
				this.fullTextMessageCallbackArgs = null;
			}
		};

		this.getTextCommandsSize = function()
		{
			var result = 0;
			for (var i = 0; i < this.file.pages.length; i++)
			{
				if (this.file.pages[i].text)
					result += this.file.pages[i].text.length;
			}
			return result;
		};


		this.Get_AllFontNames = function()
		{
			return {
				// to do
			}
		};

		this.setRotatePage = function(pageNum, angle, ismultiply)
		{
			if (!this.file || !this.file.isValid())
				return;

			if (undefined === pageNum)
				pageNum = this.currentPage;

			let page = this.file.pages[pageNum];
			if (!page)
				return;

			angle = (angle / 90) >> 0;

			if (page.angle && ismultiply)
				page.angle += angle;
			else
				page.angle = angle;

			page.angle += 4;
			page.angle &= 0x03;

			if (0 === page.angle)
				delete page.angle;

			this.resize();
			this.thumbnails && this.thumbnails.resize();
		};

		this.getPageRotate = function(pageNum)
		{
			if (!this.file || !this.file.isValid())
				return 0;

			if (!this.file.pages[pageNum])
				return 0;

			let value = this.file.pages[pageNum].angle;
			return (undefined === value) ? 0 : value;
		};

		this.setOffsetTop = function(offset)
		{
			this.offsetTop = offset;
			this.resize();
		};
		this.createComponents();
	};
	CHtmlPage.prototype.getPDFDoc = function()
	{
		return this.doc;
	};
	CHtmlPage.prototype._paintForms = function()
	{
		const ctx = this.canvasForms.getContext('2d');
		ctx.globalAlpha = 1;
		
		let xCenter = this.width >> 1;
		let yPos = this.scrollY >> 0;
		if (this.documentWidth > this.width)
		{
			xCenter = (this.documentWidth >> 1) - (this.scrollX) >> 0;
		}
		
		for (let i = this.startVisiblePage; i <= this.endVisiblePage; i++)
		{
			let page = this.drawingPages[i];
			if (!page)
				break;

			let aForms = this.pagesInfo.pages[i].fields != null ? this.pagesInfo.pages[i].fields : null;
			if (this.pagesInfo.pages[i].graphics == null)
				this.pagesInfo.pages[i].graphics = {};
			
			if (!aForms)
				continue;

			
			let w = (page.W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
			let h = (page.H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

			let cachedImg = page.ImageForms;
			if (!cachedImg || this.pagesInfo.pages[i].needRedrawForms || cachedImg.width != w || cachedImg.height != h)
			{
				// рисуем на отдельном канвасе, кешируем
				let tmpCanvas		= page.ImageForms ? page.ImageForms : document.createElement('canvas');
				let tmpCanvasCtx	= tmpCanvas.getContext('2d');
				
				tmpCanvas.width		= w;
				tmpCanvas.height	= h;

				if (page.ImageForms)
					tmpCanvasCtx.clearRect(0, 0, w, h);

				let nScale			= AscCommon.AscBrowser.retinaPixelRatio * this.zoom;
				let widthPx			= this.canvas.width;
				let heightPx		= this.canvas.height;
				
				let oGraphicsPDF = new AscPDF.CPDFGraphics();
				this.pagesInfo.pages[i].graphics.pdf = oGraphicsPDF;
				oGraphicsPDF.Init(tmpCanvasCtx, widthPx * nScale, heightPx * nScale);
				oGraphicsPDF.SetCurPage(i);

				let oGraphicsWord = new AscCommon.CGraphics();
				this.pagesInfo.pages[i].graphics.word = oGraphicsWord;
				oGraphicsWord.init(tmpCanvasCtx, widthPx * nScale, heightPx * nScale, widthPx * g_dKoef_pix_to_mm, heightPx * g_dKoef_pix_to_mm);
				oGraphicsWord.m_oFontManager = AscCommon.g_fontManager;
				oGraphicsWord.endGlobalAlphaColor = [255, 255, 255];
				oGraphicsWord.transform(1, 0, 0, 1, 0, 0);
				
				if (this.pagesInfo.pages[i].fields != null) {
					this.pagesInfo.pages[i].fields.forEach(function(field) {
						field.DrawOnPage(oGraphicsPDF, oGraphicsWord, i);
					});
				}
				
				page.ImageForms = tmpCanvas;
				this.pagesInfo.pages[i].needRedrawForms = false;
			}
			
			let x = ((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - (w >> 1);
			let y = ((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
			
			ctx.drawImage(page.ImageForms, 0, 0, page.ImageForms.width, page.ImageForms.height, x, y, w, h);
		}
		
		let oDoc = this.getPDFDoc();
		if (oDoc.activeForm && oDoc.activeForm.UpdateScroll) {
			if (oDoc.activeForm.IsNeedDrawHighlight())
				oDoc.activeForm.UpdateScroll(false);
			else
				oDoc.activeForm.UpdateScroll(true);
		}
			
		if (oDoc.activeForm && [AscPDF.FIELD_TYPES.combobox, AscPDF.FIELD_TYPES.text].includes(oDoc.activeForm.GetType()))
			oDoc.activeForm.content.RecalculateCurPos();
	};
	CHtmlPage.prototype._paintAnnots = function()
	{
		const ctx = this.canvasForms.getContext('2d');
		ctx.clearRect(0, 0, this.canvasForms.width, this.canvasForms.height);
		ctx.globalAlpha = 1;
		
		let xCenter = this.width >> 1;
		let yPos = this.scrollY >> 0;
		if (this.documentWidth > this.width)
		{
			xCenter = (this.documentWidth >> 1) - (this.scrollX) >> 0;
		}
		
		//let time1 = performance.now();
		for (let i = this.startVisiblePage; i <= this.endVisiblePage; i++)
		{
			let page = this.drawingPages[i];
			if (!page)
				break;

			let aAnnots = this.pagesInfo.pages[i].annots != null ? this.pagesInfo.pages[i].annots : null;
			if (this.pagesInfo.pages[i].graphics == null)
				this.pagesInfo.pages[i].graphics = {};
			
			if (!aAnnots)
				continue;

			// рисуем на отдельном канвасе, кешируем
			let tmpCanvas = page.ImageAnnots ? page.ImageAnnots : document.createElement('canvas');
			
			let cachedImg = page.ImageAnnots;
			let w = (page.W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
			let h = (page.H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
			
			if (!cachedImg || this.pagesInfo.pages[i].needRedrawAnnots || cachedImg.width != w || cachedImg.height != h)
			{
				tmpCanvas.width = w;
				tmpCanvas.height = h;
				let tmpCanvasCtx = tmpCanvas.getContext('2d');
				
				let nScale		= AscCommon.AscBrowser.retinaPixelRatio * this.zoom;
				let widthPx		= this.canvas.width;
				let heightPx	= this.canvas.height;
				
				let oGraphicsWord = new AscCommon.CGraphics();
				this.pagesInfo.pages[i].graphics.word = oGraphicsWord;
				oGraphicsWord.init(tmpCanvasCtx, widthPx * nScale, heightPx * nScale, widthPx * g_dKoef_pix_to_mm, heightPx * g_dKoef_pix_to_mm);
				oGraphicsWord.m_oFontManager = AscCommon.g_fontManager;
				oGraphicsWord.endGlobalAlphaColor = [255, 255, 255];
				oGraphicsWord.transform(1, 0, 0, 1, 0, 0);
				
				let oGraphicsPDF = new AscPDF.CPDFGraphics();
				this.pagesInfo.pages[i].graphics.pdf = oGraphicsPDF;
				oGraphicsPDF.Init(tmpCanvasCtx, widthPx * nScale, heightPx * nScale);
				oGraphicsPDF.SetCurPage(i);
				
				if (this.pagesInfo.pages[i].annots != null) {
					this.pagesInfo.pages[i].annots.forEach(function(annot) {
						if (annot.IsTextMarkup() == false) {
							if (annot.IsNeedDrawFromStream() == false) {
								annot.Draw(oGraphicsPDF, oGraphicsWord);
							}
							else {
								annot.Recalculate();
								annot.DrawFromStream(oGraphicsPDF);
							}
						}
					});
				}
				
				page.ImageAnnots			= tmpCanvas;
				page.ImageAnnots.maxRect	= oGraphicsPDF.GetDrawedRect(true);

				this.pagesInfo.pages[i].needRedrawAnnots = false;
			}
			
			// if (this.pagesInfo.pages[i].annots != null) {
			// 	let bFromStream = this.pagesInfo.pages[i].annots.find(function(field) {
			// 		if (field.IsNeedDrawFromStream() == true)
			// 			return true;
			// 	});
				
			// 	if (bFromStream) {
			// 		this.pagesInfo.pages[i].annots.forEach(function(field) {
			// 			// если форма не менялась, рисуем внешний вид из потока
			// 			if (field.IsNeedDrawFromStream() == true)
			// 				field.DrawFromStream();
			// 		});
			// 	}
			// }
			
			let x = (((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0)) - (w >> 1);
			let y = (((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0);
			// let x = (((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0)) - (w >> 1) + page.ImageAnnots.maxRect.xMin;
			// let y = (((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0) + page.ImageAnnots.maxRect.yMin;
			
			ctx.drawImage(page.ImageAnnots, 0, 0, page.ImageAnnots.width, page.ImageAnnots.height, x, y, w, h);

			// let wCropped = page.ImageAnnots.maxRect.xMax - page.ImageAnnots.maxRect.xMin;
			// let hCropped = page.ImageAnnots.maxRect.yMax - page.ImageAnnots.maxRect.yMin;

			// ctx.drawImage(page.ImageAnnots, page.ImageAnnots.maxRect.xMin, page.ImageAnnots.maxRect.yMin, wCropped, hCropped, x, y, wCropped, hCropped);
			//let time2 = performance.now();
			// console.log("time: " + (time2 - time1));
		}
		
		// if (this.activeForm && this.activeForm.UpdateScroll)
		// 	this.activeForm.UpdateScroll(true);
		// if (this.activeForm && [AscPDF.FIELD_TYPES.combobox, AscPDF.FIELD_TYPES.text].includes(this.activeForm.GetType()))
		// 	this.activeForm.content.RecalculateCurPos();
	};
	CHtmlPage.prototype._paintMarkupAnnotsOnPage = function(pageIndex, ctx)
	{
		let xCenter = this.width >> 1;
		let yPos = this.scrollY >> 0;
		if (this.documentWidth > this.width)
		{
			xCenter = (this.documentWidth >> 1) - (this.scrollX) >> 0;
		}
		
        let aAnnots = this.pagesInfo.pages[pageIndex].annots != null ? this.pagesInfo.pages[pageIndex].annots : null;
		if (this.pagesInfo.pages[pageIndex].graphics == null)
			this.pagesInfo.pages[pageIndex].graphics = {};
		
		if (!aAnnots)
			return;
		
		let page = this.drawingPages[pageIndex];
		if (!page)
			return;

		let w = (page.W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
		let h = (page.H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
		
		let indLeft = ((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - (w >> 1);
        let indTop  = ((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

		if (this.pagesInfo.pages[pageIndex].needRedrawHighlights || false == this.bCachedMarkupAnnnots)
		{
			let nScale		= AscCommon.AscBrowser.retinaPixelRatio * this.zoom;
			let widthPx		= this.canvas.width;
			let heightPx	= this.canvas.height;
			
			let oGraphicsPDF = new AscPDF.CPDFGraphics();
			oGraphicsPDF.SetCurPage(pageIndex);
			
			this.pagesInfo.pages[pageIndex].graphics.pdf = oGraphicsPDF;
			oGraphicsPDF.Init(ctx, widthPx * nScale, heightPx * nScale);

			if (false == this.bCachedMarkupAnnnots) {
				ctx.save();
				ctx.beginPath();
				ctx.rect(indLeft, indTop, w, h);
				ctx.clip();
				ctx.setTransform(1, 0, 0, 1, indLeft, indTop);
			}
				
			
			if (this.pagesInfo.pages[pageIndex].annots != null) {
				this.pagesInfo.pages[pageIndex].annots.forEach(function(annot) {
					if (annot.IsTextMarkup()) {
						if (false == annot.IsNeedDrawFromStream())
							annot.Draw(oGraphicsPDF);
						else
							annot.DrawFromStream(oGraphicsPDF);
					}
				});
			}
			
			if (false == this.bCachedMarkupAnnnots) {
				ctx.restore();
			}

			this.pagesInfo.pages[pageIndex].needRedrawHighlights = false;
		}
	};
	CHtmlPage.prototype._paintFormsHighlight = function()
	{
		let oCtx = this.canvasForms.getContext("2d");
		for (let i = this.startVisiblePage; i <= this.endVisiblePage; i++)
		{
			let page = this.drawingPages[i];
			if (!page)
				break;

			let aForms = this.pagesInfo.pages[i].fields != null ? this.pagesInfo.pages[i].fields : null;
			
			if (!aForms)
				continue;
			
			if (this.pagesInfo.pages[i].fields != null) {
				this.pagesInfo.pages[i].fields.forEach(function(field) {
					if (field.IsNeedDrawHighlight())
						field.DrawHighlight(oCtx);
				});
			}
		}
	};
	CHtmlPage.prototype._paintComboboxesMarkers = function()
	{
		let oCtx = this.canvasForms.getContext("2d");
		for (let i = this.startVisiblePage; i <= this.endVisiblePage; i++)
		{
			let page = this.drawingPages[i];
			if (!page)
				break;

			let aForms = this.pagesInfo.pages[i].fields != null ? this.pagesInfo.pages[i].fields : null;
			
			if (!aForms)
				continue;
			
			if (this.pagesInfo.pages[i].fields != null) {
				this.pagesInfo.pages[i].fields.forEach(function(field) {
					if (field.GetType() == AscPDF.FIELD_TYPES.combobox)
						field.DrawMarker(oCtx);
				});
			}
		}
	};
	CHtmlPage.prototype.createComponents = function()
	{
		var elements = "<div id=\"id_main\" class=\"block_elem\" style=\"touch-action:none;-ms-touch-action: none;-moz-user-select:none;-khtml-user-select:none;user-select:none;background-color:" + AscCommon.GlobalSkin.BackgroundColor + ";overflow:hidden;\" UNSELECTABLE=\"on\">";
		elements += "<canvas id=\"id_viewer\" class=\"block_elem\" style=\"left:0px;top:0px;width:100;height:100;\"></canvas>";
		elements += "<canvas id=\"id_forms\" class=\"block_elem\" style=\"left:0px;top:0px;width:100;height:100;\"></canvas>";
		elements += "<div id=\"id_target_cursor\" class=\"block_elem\" width=\"1\" height=\"1\" style=\"touch-action:none;-ms-touch-action: none;-webkit-user-select: none;width:2px;height:13px;z-index:4;\"></div>"
		elements += "<canvas id=\"id_overlay\" class=\"block_elem\" style=\"left:0px;top:0px;width:100;height:100;\"></canvas>";
		elements += "</div>";

		elements += "<div id=\"id_vertical_scroll\" class=\"block_elem\" style=\"display:none;left:0px;top:0px;width:0px;height:0px;\"></div>";
		elements += "<div id=\"id_horizontal_scroll\" class=\"block_elem\" style=\"display:none;left:0px;top:0px;width:0px;height:0px;\"></div>";
		//this.parent.style.backgroundColor = this.backgroundColor; <= this color from theme
		this.parent.innerHTML = elements;
		
		let oControl = editor.WordControl.m_oBody.Controls.find(function(control) {
			return control.HtmlElement.id == "id_main";
		});
		oControl.HtmlElement = document.getElementById("id_main");
		
		this.id_main = oControl.HtmlElement;

		this.canvas = document.getElementById("id_viewer");
		this.canvas.style.backgroundColor = this.backgroundColor;
		
		this.canvasOverlay = document.getElementById("id_overlay");
		this.canvasOverlay.style.pointerEvents = "none";
		
		this.canvasForms = document.getElementById("id_forms");
		
		this.Api.WordControl.m_oDrawingDocument.TargetHtmlElement = document.getElementById('id_target_cursor');
		
		this.overlay = new AscCommon.COverlay();
		this.overlay.m_oControl = { HtmlElement : this.canvasOverlay };
		this.overlay.m_bIsShow = true;
		this.DrawingObjects.drawingDocument.AutoShapesTrack.m_oOverlay = this.overlay;
		this.overlay.Clear();
		this.DrawingObjects.drawingDocument.AutoShapesTrack.m_oContext = this.overlay.m_oContext;
		this.overlay.m_oHtmlPage = this.Api.WordControl;
		this.Api.WordControl.m_bIsRuler = false;
		this.updateSkin();
	};
	CHtmlPage.prototype.resize = function(isDisablePaint)
	{
		let oThis = this;
		this.isFocusOnThumbnails = false;
		
		var rect = this.canvas.getBoundingClientRect();
		this.x = rect.left;
		this.y = rect.top;
		
		var oldsize = {w: this.width, h: this.height};
		this.width = this.parent.offsetWidth - this.scrollWidth;
		this.height = this.parent.offsetHeight;
		
		if (this.zoomMode === ZoomMode.Width)
			this.zoom = this.calculateZoomToWidth();
		else if (this.zoomMode === ZoomMode.Page)
			this.zoom = this.calculateZoomToHeight();
		
		// в мобильной версии мы будем получать координаты от MobileTouchManager (до этого момента они уже должны быть) и не нужно их запоминать, так как мы перетрём нужные нам значения
		// ну а если их нет и зум произошёл не от тача, то запоминаем их как при обычном зуме
		if (!this.zoomCoordinate)
			this.fixZoomCoord( (this.width >> 1), (this.height >> 1) );
		
		this.sendEvent("onZoom", this.zoom, this.zoomMode);
		
		this.recalculatePlaces();
		
		this.isVisibleHorScroll = (this.documentWidth > this.width) ? true : false;
		if (this.isVisibleHorScroll)
			this.height -= this.scrollWidth;
		
		this.canvas.style.width = this.width + "px";
		this.canvas.style.height = this.height + "px";
		AscCommon.calculateCanvasSize(this.canvas);
		
		this.canvasOverlay.style.width = this.width + "px";
		this.canvasOverlay.style.height = this.height + "px";
		AscCommon.calculateCanvasSize(this.canvasOverlay);
		
		this.canvasForms.style.width = this.width + "px";
		this.canvasForms.style.height = this.height + "px";
		AscCommon.calculateCanvasSize(this.canvasForms);
		
		var scrollV = document.getElementById("id_vertical_scroll");
		scrollV.style.display = "block";
		scrollV.style.left = this.width + "px";
		scrollV.style.top = "0px";
		scrollV.style.width = this.scrollWidth + "px";
		scrollV.style.height = this.height + "px";
		
		var scrollH = document.getElementById("id_horizontal_scroll");
		scrollH.style.display = this.isVisibleHorScroll ? "block" : "none";
		scrollH.style.left = "0px";
		scrollH.style.top = this.height + "px";
		scrollH.style.width = this.width + "px";
		scrollH.style.height = this.scrollWidth + "px";
		
		var settings = this.CreateScrollSettings();
		settings.isHorizontalScroll = true;
		settings.isVerticalScroll = false;
		settings.contentW = this.documentWidth;
		
		if (this.m_oScrollHorApi)
			this.m_oScrollHorApi.Repos(settings, this.isVisibleHorScroll);
		else
		{
			this.m_oScrollHorApi = new AscCommon.ScrollObject("id_horizontal_scroll", settings);
			
			this.m_oScrollHorApi.onLockMouse  = function(evt) {
				AscCommon.check_MouseDownEvent(evt, true);
				AscCommon.global_mouseEvent.LockMouse();
			};
			this.m_oScrollHorApi.offLockMouse = function(evt) {
				AscCommon.check_MouseUpEvent(evt);
			};
			this.m_oScrollHorApi.bind("scrollhorizontal", function(evt) {
				oThis.scrollHorizontal(evt.scrollD, evt.maxScrollX);
			});
		}
		
		settings = this.CreateScrollSettings();
		settings.isHorizontalScroll = false;
		settings.isVerticalScroll = true;
		settings.contentH = this.documentHeight;
		if (this.m_oScrollVerApi)
			this.m_oScrollVerApi.Repos(settings, undefined, true);
		else
		{
			this.m_oScrollVerApi = new AscCommon.ScrollObject("id_vertical_scroll", settings);
			
			this.m_oScrollVerApi.onLockMouse  = function(evt) {
				AscCommon.check_MouseDownEvent(evt, true);
				AscCommon.global_mouseEvent.LockMouse();
			};
			this.m_oScrollVerApi.offLockMouse = function(evt) {
				AscCommon.check_MouseUpEvent(evt);
			};
			this.m_oScrollVerApi.bind("scrollvertical", function(evt) {
				oThis.scrollVertical(evt.scrollD, evt.maxScrollY);
			});
		}
		
		this.scrollMaxX = this.m_oScrollHorApi.getMaxScrolledX();
		this.scrollMaxY = this.m_oScrollVerApi.getMaxScrolledY();
		
		if (this.scrollX >= this.scrollMaxX)
			this.scrollX = this.scrollMaxX;
		if (this.scrollY >= this.scrollMaxY)
			this.scrollY = this.scrollMaxY;
		
		if (this.zoomCoordinate && this.isDocumentContentReady)
		{
			var newPoint = this.ConvertCoordsToCursor(this.zoomCoordinate.x, this.zoomCoordinate.y, this.zoomCoordinate.index);
			// oldsize используется чтобы при смене ориентации экрана был небольшой скролл
			var shiftX = this.Api.isMobileVersion ? ( (oldsize.w - this.width) >> 1) : 0;
			var shiftY = this.Api.isMobileVersion ? ( (oldsize.h - this.height) >> 1) : 0;
			var newScrollX = this.scrollX + newPoint.x - this.zoomCoordinate.xShift + shiftX;
			var newScrollY = this.scrollY + newPoint.y - this.zoomCoordinate.yShift + shiftY;
			newScrollX = Math.max(0, Math.min(newScrollX, this.scrollMaxX) );
			newScrollY = Math.max(0, Math.min(newScrollY, this.scrollMaxY) );
			if (this.scrollY == 0 && !this.Api.isMobileVersion)
				newScrollY = 0;
			
			this.m_oScrollVerApi.scrollToY(newScrollY);
			this.m_oScrollHorApi.scrollToX(newScrollX);
		}
		
		if (this.thumbnails)
			this.thumbnails.resize();
		
		if (true !== isDisablePaint)
			this.timerSync();
		
		if (this.Api.WordControl.MobileTouchManager)
			this.Api.WordControl.MobileTouchManager.Resize();
		
		if (!this.Api.isMobileVersion || !this.skipClearZoomCoord)
			this.clearZoomCoord();
	};
	CHtmlPage.prototype.repaintFormsOnPage = function(pageIndex)
	{
		if (this.pagesInfo.pages.length > pageIndex)
		{
			this.pagesInfo.pages[pageIndex].needRedrawForms = true;
			this.isRepaint = true;
		}
	};
	CHtmlPage.prototype.Save = function()
	{
		let memoryInitSize = 1024 * 500; // 500Kb
		let oMemory	= null;
		let aPages	= this.pagesInfo.pages;

		// по информации аннотаций определим какие были удалены
		let oDoc		= this.getPDFDoc();
		let aAnnotsInfo	= this.file.nativeFile["getAnnotationsInfo"]();
		let aDeleted	= [];
		aAnnotsInfo.forEach(function(oInfo) {
			if (oInfo["StateModel"] == AscPDF.TEXT_ANNOT_STATE_MODEL.Review)
				return;
			
			let isInDoc = oDoc.annots.find(function(annot) {
				return annot.GetApIdx() == oInfo["AP"]["i"] || annot._replies.find(function(reply) {
					return reply.GetApIdx() == oInfo["AP"]["i"];
				});
			});

			if (!isInDoc) {
				if (aDeleted[oInfo["page"]] == null) {
					aDeleted[oInfo["page"]] = [];
				}

				aDeleted[oInfo["page"]].push(oInfo["AP"]["i"]);
			}
		});

		for (let i = 0; i < aPages.length; i++)
		{
			if ((aPages[i].annots == null || aPages[i].annots.length === 0) && (aPages[i].fields == null || aPages[i].fields.length === 0) && !aDeleted[i])
				continue;

			if (!oMemory)
			{
				oMemory = new AscCommon.CMemory(true);
				oMemory.Init(memoryInitSize);
			}

			let nStartPos = oMemory.GetCurPosition();
			oMemory.Skip(4);
			oMemory.WriteByte(0); // Annotation
			oMemory.WriteLong(i);
			
			if (aPages[i].annots) {
				for (let nAnnot = 0; nAnnot < aPages[i].annots.length; nAnnot++) {
					aPages[i].annots[nAnnot].WriteToBinary && aPages[i].annots[nAnnot].IsChanged() && aPages[i].annots[nAnnot].WriteToBinary(oMemory);
				}
			}

			if (aDeleted[i]) {
				for (let j = 0; j < aDeleted[i].length; j++) {
					oMemory.WriteByte(AscCommon.CommandType.ctAnnotFieldDelete);
					oMemory.WriteLong(8);
					oMemory.WriteLong(aDeleted[i][j]);
				}
			}

			// forms
			if (aPages[i].fields) {
				for (let nForm = 0; nForm < aPages[i].fields.length; nForm++) {
					aPages[i].fields[nForm].WriteToBinary && aPages[i].fields[nForm].IsChanged() && aPages[i].fields[nForm].WriteToBinary(oMemory);
				}
			}

			let nEndPos = oMemory.GetCurPosition();
			oMemory.Seek(nStartPos);
			oMemory.WriteLong(nEndPos - nStartPos);
			oMemory.Seek(nEndPos);
		}

		if (oMemory)
			return new Uint8Array(oMemory.data.buffer, 0, oMemory.GetCurPosition());
		return null;
	};

	function CCurrentPageDetector(w, h)
	{
		// размеры окна
		this.width = w;
		this.height = h;

		this.pages = [];
	}
	CCurrentPageDetector.prototype.addPage = function(num, x, y, w, h)
	{
		this.pages.push({ num : num, x : x, y : y, w : w, h : h });
	};
	CCurrentPageDetector.prototype.getCurrentPage = function(currentPage)
	{
		var count = this.pages.length;
		var visibleH = 0;
		var page, currentVisibleH;
		var pageNum = 0;
		for (var i = 0; i < count; i++)
		{
			page = this.pages[i];
			currentVisibleH = Math.min(this.height, page.y + page.h) - Math.max(0, page.y);
			if (currentVisibleH == page.h)
			{
				// первая полностью видимая страница
				pageNum = i;
				break;
			}

			if (currentVisibleH > visibleH)
			{
				visibleH = currentVisibleH;
				pageNum = i;
			}
		}

		page = this.pages[pageNum];
		if (!page)
		{
			return {
				num : currentPage,
				x : 0,
				y : 0,
				r : 1,
				b : 1
			};
		}

		var x = 0;
		if (page.x < 0) 
			x = -page.x / page.w;

		var y = 0;
		if (page.y < 0)
			y = -page.y / page.h;

		var r = 1;
		if ((page.x + page.w) > this.width)
			r -= (page.x + page.w - this.width) / page.w;

		var b = 1;
		if ((page.y + page.h) > this.height)
			b -= (page.y + page.h - this.height) / page.h;
		
		return {
			num : page.num,
			x : x,
			y : y,
			r : r,
			b : b
		};
	};
	
	AscCommon.CViewer = CHtmlPage;
	AscCommon.ViewerZoomMode = ZoomMode;
	AscCommon.CCacheManager = CCacheManager;

	if (!window["AscPDF"])
	    window["AscPDF"] = {};

	window["AscPDF"].CPageInfo = CPageInfo;

})();
