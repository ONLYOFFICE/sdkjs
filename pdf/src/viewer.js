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

		this.isXP = ((AscCommon.AscBrowser.userAgent.indexOf("windowsxp") > -1) || (AscCommon.AscBrowser.userAgent.indexOf("chrome/49") > -1)) ? true : false;
		if (!this.isXP && AscCommon.AscBrowser.isIE && !AscCommon.AscBrowser.isIeEdge)
			this.isXP = true;

		if (this.isXP)
		{
			AscCommon.g_oHtmlCursor.register("grab", "grab", "7 8", "pointer");
			AscCommon.g_oHtmlCursor.register("grabbing", "grabbing", "6 6", "pointer");
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

		this.createComponents = function()
		{
			this.updateSkin();

			var elements = "";
			elements += "<canvas id=\"id_viewer\" class=\"block_elem\" style=\"left:0px;top:0px;width:100;height:100;\"></canvas>";
			elements += "<canvas id=\"id_overlay\" class=\"block_elem\" style=\"left:0px;top:0px;width:100;height:100;\"></canvas>";
			elements += "<div id=\"id_vertical_scroll\" class=\"block_elem\" style=\"display:none;left:0px;top:0px;width:0px;height:0px;\"></div>";
			elements += "<div id=\"id_horizontal_scroll\" class=\"block_elem\" style=\"display:none;left:0px;top:0px;width:0px;height:0px;\"></div>";
		
			//this.parent.style.backgroundColor = this.backgroundColor; <= this color from theme
			this.parent.innerHTML = elements;

			this.canvas = document.getElementById("id_viewer");
			this.canvas.style.backgroundColor = this.backgroundColor;

			this.canvasOverlay = document.getElementById("id_overlay");
			this.canvasOverlay.style.pointerEvents = "none";

			this.overlay = new AscCommon.COverlay();
			this.overlay.m_oControl = { HtmlElement : this.canvasOverlay };
			this.overlay.m_bIsShow = true;
		};
		this.createComponents();

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

			this.paint();
		};
		this.scrollVertical = function(pos, maxPos)
		{
			this.scrollY = pos;
			this.scrollMaxY = maxPos;
			if (this.Api.WordControl.MobileTouchManager && this.Api.WordControl.MobileTouchManager.iScroll)
				this.Api.WordControl.MobileTouchManager.iScroll.y = - Math.max(0, Math.min(pos, maxPos));

			this.paint();
		};

		this.resize = function(isDisablePaint)
		{
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

			settings = this.CreateScrollSettings();
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

		this.onLoadModule = function()
		{
			this.moduleState = ModuleState.Loaded;
			window["AscViewer"]["InitializeFonts"]();

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
				}
			}
			else
			{
				if (this.file)
					this.file.close();

				this.file = window["AscViewer"].createFile(data);

				if (this.file)
				{
					this.SearchResults = this.file.SearchResults;
					this.file.viewer = this;
				}
			}

			if (!this.file)
			{
				this.Api.sendEvent("asc_onError", c_oAscError.ID.ConvertationOpenError, c_oAscError.Level.Critical);
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

			this.pagesInfo.setCount(this.file.pages.length);

			// в интерфейсе есть проблема - нужно посылать onDocumentContentReady после setAdvancedOptions
			setTimeout(function(){

				if (!_t.isStarted)
				{
					AscCommon.addMouseEvent(_t.canvas, "down", _t.onMouseDown);
					AscCommon.addMouseEvent(_t.canvas, "move", _t.onMouseMove);
					AscCommon.addMouseEvent(_t.canvas, "up", _t.onMouseUp);

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

			if (this.thumbnails)
				this.thumbnails.init(this);

			this.setMouseLockMode(true);
		};

		this.setZoom = function(value)
		{
			var oldZoom = this.zoom;
			this.zoom = value;
			this.zoomMode = ZoomMode.Custom;
			this.sendEvent("onZoom", this.zoom);
			this.resize(oldZoom);
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

		this.navigateToPage = function(pageNum, yOffset)
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
			this.m_oScrollVerApi.scrollToY(posY);
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
			}

			this.documentWidth = 0;
			for (let i = 0, len = this.drawingPages.length; i < len; i++)
			{
				if (this.drawingPages[i].W > this.documentWidth)
					this.documentWidth = this.drawingPages[i].W;
			}
			// прибавим немного
			this.documentWidth += (4 * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

			var curTop = this.betweenPages;
			for (let i = 0, len = this.drawingPages.length; i < len; i++)
			{
				this.drawingPages[i].X = (this.documentWidth - this.drawingPages[i].W) >> 1;
				this.drawingPages[i].Y = curTop;

				curTop += this.drawingPages[i].H;
				curTop += this.betweenPages;
			}

			this.documentHeight = curTop;
			this.paint();			
		};

		this.setCursorType = function(cursor)
		{
			if (this.isXP)
			{
				this.canvas.style.cursor = AscCommon.g_oHtmlCursor.value(cursor);
				return;
			}

			this.canvas.style.cursor = cursor;
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

		this.onMouseDown = function(e)
		{
			oThis.isFocusOnThumbnails = false;
			AscCommon.stopEvent(e);

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
						oThis.Api.sendEvent("asc_onContextMenu", undefined);
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
			oThis.mouseDownLinkObject = oThis.getPageLinkByMouse();

			// нажали мышь - запомнили координаты и находимся ли на ссылке
			// при выходе за epsilon на mouseMove - сэмулируем нажатие
			// так что тут только курсор

			var cursorType;
			if (oThis.mouseDownLinkObject)
				cursorType = "pointer";
			else
			{
				if (oThis.MouseHandObject)
					cursorType = "grabbing";
				else
					cursorType = "default";
			}

			oThis.setCursorType(cursorType);

			if (!oThis.MouseHandObject && !oThis.mouseDownLinkObject)
			{
				// ждать смысла нет
				oThis.isMouseMoveBetweenDownUp = true;
				oThis.onMouseDownEpsilon();
			}
		};

		this.onMouseDownEpsilon = function()
		{
			if (oThis.MouseHandObject)
			{
				if (oThis.mouseDownLinkObject)
				{
					// если нажали на ссылке - то не зажимаем лапу
					oThis.setCursorType("pointer");
					return;
				}
				// режим лапы. просто начинаем режим Active - зажимаем лапу
				oThis.setCursorType("grabbing");
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
			if (e && e.preventDefault)
				e.preventDefault();

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

			if (oThis.MouseHandObject)
			{
				if (oThis.mouseDownLinkObject)
				{
					// смотрим - если совпало со ссылкой при нажатии - то переходим по ней
					var mouseUpLinkObject = oThis.getPageLinkByMouse();
					if (mouseUpLinkObject === oThis.mouseDownLinkObject)
					{
						oThis.navigateToLink(mouseUpLinkObject);
					}

					// если нет - то ничего не делаем
					if (mouseUpLinkObject)
						oThis.setCursorType("pointer");
					else
						oThis.setCursorType("grab");
				}
				else if (!oThis.isMouseMoveBetweenDownUp)
				{
					oThis.setCursorType("grab");

					// делаем клик в логическом документе, чтобы сбросить селект, если он был
					var pageObjectLogic = oThis.getPageByCoords2(AscCommon.global_mouseEvent.X - oThis.x, AscCommon.global_mouseEvent.Y - oThis.y);
					oThis.file.onMouseDown(pageObjectLogic.index, pageObjectLogic.x, pageObjectLogic.y);
					oThis.file.onMouseUp(pageObjectLogic.index, pageObjectLogic.x, pageObjectLogic.y);
				}
				else
				{
					oThis.setCursorType("grab");
				}

				oThis.isMouseMoveBetweenDownUp = false;
				oThis.MouseHandObject.Active = false;
				oThis.mouseDownLinkObject = null;
				return;
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

			AscCommon.check_MouseMoveEvent(e);
			if (e && e.preventDefault)
				e.preventDefault();

			// если мышка нажата и еще не вышли за eps - то проверяем, модет вышли сейчас?
			// и, если вышли - то эмулируем
			if (oThis.isMouseDown && !oThis.isMouseMoveBetweenDownUp)
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
				if (oThis.MouseHandObject.Active)
				{
					// двигаем рукой
					oThis.setCursorType("grabbing");

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
					if (oThis.isMouseDown)
					{
						if (oThis.mouseDownLinkObject)
						{
							// не меняем курсор с "ссылочного", если зажимали на ссылке
							oThis.setCursorType("pointer");
						}
						else
						{
							// даже если не двигали еще и ждем eps, все равно курсор меняем на зажатый
							oThis.setCursorType("grabbing");
						}
					}
					else
					{
						// просто водим мышкой - тогда смотрим, на ссылке или нет, чтобы выставить курсор
						var mouseMoveLinkObject = oThis.getPageLinkByMouse();
						if (mouseMoveLinkObject)
							oThis.setCursorType("pointer");
						else
							oThis.setCursorType("grab");
					}
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
					if (oThis.isMouseMoveBetweenDownUp)
					{
						// нажатая мышка - курсор всегда default (так как за eps вышли)
						oThis.setCursorType("default");

						var pageObjectLogic = oThis.getPageByCoords2(AscCommon.global_mouseEvent.X - oThis.x, AscCommon.global_mouseEvent.Y - oThis.y);
						oThis.file.onMouseMove(pageObjectLogic.index, pageObjectLogic.x, pageObjectLogic.y);
					}
					else
					{
						// пока на ссылке
						oThis.setCursorType("pointer");
					}
				}
				else
				{
					var mouseMoveLinkObject = oThis.getPageLinkByMouse();
					if (mouseMoveLinkObject)
						oThis.setCursorType("pointer");
					else
						oThis.setCursorType("default");
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
			if (global_mouseEvent.Y < positionMinY)
			{
				var delta = 30;
				if (20 > (positionMinY - global_mouseEvent.Y))
					delta = 10;

				scrollYVal = -delta;
			}
			else if (global_mouseEvent.Y > positionMaxY)
			{
				var delta = 30;
				if (20 > (global_mouseEvent.Y - positionMaxY))
					delta = 10;

				scrollYVal = delta;
			}

			var scrollXVal = 0;
			if (oThis.isVisibleHorScroll)
			{
				var positionMinX = oThis.x;
				var positionMaxX = oThis.x + oThis.width;

				if (global_mouseEvent.X < positionMinX)
				{
					var delta = 30;
					if (20 > (positionMinX - global_mouseEvent.X))
						delta = 10;

					scrollXVal = -delta;
				}
				else if (global_mouseEvent.X > positionMaxX)
				{
					var delta = 30;
					if (20 > (global_mouseEvent.X - positionMaxX))
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

			//if (!this.MouseHandObject)
			{
				ctx.fillStyle = "rgba(51,102,204,255)";
				ctx.beginPath();

				for (let i = this.startVisiblePage; i <= this.endVisiblePage; i++)
				{
					var pageCoords = this.pageDetector.pages[i - this.startVisiblePage];
					this.file.drawSelection(i, this.overlay, pageCoords.x, pageCoords.y, pageCoords.w, pageCoords.h);
				}

				ctx.fill();
				ctx.beginPath();
			}
			ctx.globalAlpha = 1.0;
		};

		this._paint = function()
		{
			if (!this.file.isValid())
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
					}
				}
			}

			this.pageDetector = new CCurrentPageDetector(this.canvas.width, this.canvas.height);

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

				if (!isStretchPaint)
				{
					if (!this.file.cacheManager)
					{
						if (this.isClearPages || (page.Image && ((page.Image.requestWidth != w) || (page.Image.requestHeight != h))))
							delete page.Image;
					}
					else
					{
						if (this.isClearPages || (page.Image && ((page.Image.requestWidth < w) || (page.Image.requestHeight < h))))
						{
							if (this.file.cacheManager)
								this.file.cacheManager.unlock(page.Image);

							delete page.Image;
						}
					}
				}

				if (!page.Image && !isStretchPaint)
				{
					page.Image = this.file.getPage(i, w, h, undefined, this.Api.isDarkMode ? 0x3A3A3A : 0xFFFFFF);
					if (this.Api.watermarkDraw)
						this.Api.watermarkDraw.Draw(page.Image.getContext("2d"), w, h);
				}

				let x = ((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - (w >> 1);
				let y = ((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

				if (page.Image)
				{
					ctx.drawImage(page.Image, 0, 0, page.Image.width, page.Image.height, x, y, w, h);
					this.pagesInfo.setPainted(i);
				}
				else
				{
					ctx.fillStyle = "#FFFFFF";
					ctx.fillRect(x, y, w, h);
				}
				ctx.strokeRect(x + lineW / 2, y + lineW / 2, w - lineW, h - lineW);

				this.pageDetector.addPage(i, x, y, w, h);
			}

			this.isClearPages = false;
			this.updateCurrentPage(this.pageDetector.getCurrentPage(this.currentPage));
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

		this.findText = function(text, isMachingCase, isNext, callback)
		{
			if (this.isFullTextMessage)
				return bRetValue;

			if (!this.isFullText)
			{
				this.fullTextMessageCallbackArgs = [text, isMachingCase, isNext, callback];
				this.fullTextMessageCallback = function() {
					this.file.findText(this.fullTextMessageCallbackArgs[0], this.fullTextMessageCallbackArgs[1], this.fullTextMessageCallbackArgs[2]);
					this.onUpdateOverlay();

					if (this.fullTextMessageCallbackArgs[3])
						this.fullTextMessageCallbackArgs[3](this.SearchResults.Count);
				};
				this.showTextMessage();
				return true; // async
			}

			this.file.findText(text, isMachingCase, isNext);
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

		this.OnKeyDown = function(e)
		{
			var bRetValue = false;

			if ( e.KeyCode == 33 ) // PgUp
			{
				this.m_oScrollVerApi.scrollByY(-this.height, false);
				this.timerSync();
			}
			else if ( e.KeyCode == 34 ) // PgDn
			{
				this.m_oScrollVerApi.scrollByY(this.height, false);
				this.timerSync();
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
				if (!this.isFocusOnThumbnails && this.isVisibleHorScroll)
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
				if (!this.isFocusOnThumbnails)
				{
					this.m_oScrollVerApi.scrollByY(-40);
				}
				else
				{
					var nextPage = -1;
					if (this.thumbnails)
						nextPage = this.currentPage - this.thumbnails.countPagesInBlock;
					if (nextPage < 0)
						nextPage = this.currentPage - 1;

					if (nextPage >= 0)
						this.navigateToPage(nextPage);
				}
				bRetValue = true;
			}
			else if ( e.KeyCode == 39 ) // Right Arrow
			{
				if (!this.isFocusOnThumbnails && this.isVisibleHorScroll)
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
				if (!this.isFocusOnThumbnails)
				{
					this.m_oScrollVerApi.scrollByY(40);
				}
				else
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
				bRetValue = true;
			}
			else if ( e.KeyCode == 65 && true === e.CtrlKey ) // Ctrl + A
			{
				bRetValue = true;
				if (this.isFullTextMessage)
					return bRetValue;

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
	}

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

})();
