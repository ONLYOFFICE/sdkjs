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

	var ModuleState = {
		None : 0,
		Loading : 1,
		Loaded : 2
	};

	var ZoomMode = {
		Custom : 0,
		Width : 1,
		Page : 2
	};

	function CPageInfo()
	{
		this.isReady = false;
		this.links = null;
		this.text = null;
	}
	function CDocumentPagesInfo()
	{
		this.pages = [];
		this.isCheck = false;
	}
	CDocumentPagesInfo.prototype.setCount = function(count)
	{
		this.pages = new Array(count);
		for (var i = 0; i < count; i++)
		{
			this.pages[i] = new CPageInfo();
		}
	};
	CDocumentPagesInfo.prototype.checkPage = function(index)
	{
		this.pages[index].isReady = true;
		if (null === this.pages[index].links)
			this.isCheck = true;
	};

	function CHtmlPage(id)
	{
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
		
		this.drawingPages = [];
		this.isRepaint = false;
		
		this.canvas = null;
		this.canvasOverlay = null;

		this.Selection = null;

		this.file = null;
		this.isStarted = false;

		this.scrollWidth = 14;
		this.isVisibleHorScroll = false;

		this.m_oScrollHorApi = null;
		this.m_oScrollVerApi = null;

		this.backgroundColor = "#E6E6E6";
		this.betweenPages = 20;

		this.moduleState = ModuleState.None;

		this.structure = null;
		this.currentPage = -1;

		this.startVisiblePage = -1;
		this.endVisiblePage = -1;
		this.pagesInfo = new CDocumentPagesInfo();

		this.handlers = {};

		this.SearchResults = {
			IsSearch    : false,
			Text        : "",
			MachingCase : false,
			Pages       : [],
			CurrentPage : -1,
			Current     : -1,
			Show        : false,
			Count       : 0
		};

		var oThis = this;

		this.createComponents = function()
		{
			var elements = "";
			elements += "<canvas id=\"id_viewer\" class=\"block_elem\" style=\"left:0px;top:0px;width:100;height:100;\"></canvas>";
			elements += "<canvas id=\"id_overlay\" class=\"block_elem\" style=\"left:0px;top:0px;width:100;height:100;\"></canvas>";
			elements += "<div id=\"id_vertical_scroll\" class=\"block_elem\" style=\"display:none;left:0px;top:0px;width:0px;height:0px;\"></div>";
			elements += "<div id=\"id_horizontal_scroll\" class=\"block_elem\" style=\"display:none;left:0px;top:0px;width:0px;height:0px;\"></div>";
		
			this.parent.style.backgroundColor = this.backgroundColor;
			this.parent.innerHTML = elements;

			this.canvas = document.getElementById("id_viewer");
			this.canvas.backgroundColor = this.backgroundColor;

			this.canvasOverlay = document.getElementById("id_overlay");
			this.canvasOverlay.style.pointerEvents = "none";
		};
		this.createComponents();

		this.setThumbnailsControl = function(thumbnails)
		{
			this.thumbnails = thumbnails;
			this.thumbnails.viewer = this;
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
			var isViewerTask = oThis.isRepaint;
			if (oThis.isRepaint)
			{
				oThis._paint();
				oThis.isRepaint = false;
			}
			else if (oThis.checkPageInfos())
			{
				isViewerTask = true;
			}

			if (oThis.thumbnails)
			{
				oThis.thumbnails.checkTasks(isViewerTask);
			}

			if (!oThis.UseRequestAnimationFrame)
			{
				setTimeout(oThis.timer, 40);
			}
		};

		this.CreateScrollSettings = function()
		{
			var settings = new AscCommon.ScrollSettings();
			settings.screenW = this.width;
			settings.screenH = this.height;
			settings.vscrollStep = 45;
			settings.hscrollStep = 45;
			return settings;
		};

		this.scrollHorizontal = function(pos, maxPos)
		{
			this.scrollX = pos;
			this.scrollMaxX = maxPos;
			this.paint();
		};
		this.scrollVertical = function(pos, maxPos)
		{
			this.scrollY = pos;
			this.scrollMaxY = maxPos;
			this.paint();
		};

		this.resize = function()
		{
			var rect = this.canvas.getBoundingClientRect();
			this.x = rect.x;
			this.y = rect.y;

			this.width = this.parent.offsetWidth - this.scrollWidth;
			this.height = this.parent.offsetHeight;

			if (this.zoomMode === ZoomMode.Width)
				this.zoom = this.calculateZoomToWidth();
			else if (this.zoomMode === ZoomMode.Page)
				this.zoom = this.calculateZoomToHeight();

			var lastPosition = this.getFirstPagePosition();

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

			if (lastPosition)
			{
				var drawingPage = this.drawingPages[lastPosition.page];
				var newScrollY = drawingPage.Y + lastPosition.scrollY - lastPosition.y;

				if (newScrollY < this.scrollMaxY)
					this.m_oScrollVerApi.scrollToY(newScrollY);
			}
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
			// TODO: проверить, есть ли страницы на экране
			this.paint();
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
				}

				if (this.file.isNeedPassword())
				{
					window["AscViewer"].setFilePassword(this.file, password);
				}
			}
			else
			{
				if (this.file)
					this.file.close();

				this.file = window["AscViewer"].createFile(data);
			}

			if (this.file.isNeedPassword())
			{
				this.sendEvent("onNeedPassword");
				return;
			}

			this.pagesInfo.setCount(this.file.pages.length);

			this.sendEvent("onFileOpened");

			this.file.onRepaintPages = this.onUpdatePages.bind(this);
			this.currentPage = -1;
			this.structure = this.file.getStructure();

			this.sendEvent("onPagesCount", this.file.pages.length);
			this.sendEvent("onCurrentPageChanged", 0);

			setTimeout(function(){
				oThis.sendEvent("onStructure", oThis.structure);
			}, 100);

			this.resize();

			if (!this.isStarted)
			{
				AscCommon.addMouseEvent(this.canvas, "down", this.onMouseDown);
				AscCommon.addMouseEvent(this.canvas, "move", this.onMouseMove);
				AscCommon.addMouseEvent(this.canvas, "up", this.onMouseUp);
		
				this.parent.onmousewheel = this.onMouseWhell;
				if (this.parent.addEventListener)
					this.parent.addEventListener("DOMMouseScroll", this.onMouseWhell, false);
				
				this.startTimer();				
			}

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
			if (0 === this.file.pages.length)
				return;

			var maxWidth = 0;
			for (let i = 0, len = this.file.pages.length; i < len; i++)
			{
				var pageW = (this.file.pages[i].W * 96 / this.file.pages[i].Dpi);
				if (pageW > maxWidth)
					maxWidth = pageW;
			}

			if (maxWidth < 1)
				return;

			return (this.width - 2 * this.betweenPages) / maxWidth;
		};
		this.calculateZoomToHeight = function()
		{
			if (0 === this.file.pages.length)
				return;

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
				return;

			var zoom1 = (this.width - 2 * this.betweenPages) / maxWidth;
			var zoom2 = (this.height - 2 * this.betweenPages) / maxHeight;

			return Math.min(zoom1, zoom2);
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

		this.navigate = function(id)
		{
			var item = this.structure[id];
			if (!item)
				return;

			var drawingPage = this.drawingPages[item["page"]];
			if (!drawingPage)
				return;

			var posY = drawingPage.Y;
			posY -= this.betweenPages;
			//posY += item["Y"];
			this.m_oScrollVerApi.scrollToY(posY);
		};

		this.navigateToPage = function(pageNum)
		{
			var drawingPage = this.drawingPages[pageNum];
			if (!drawingPage)
				return;

			var posY = drawingPage.Y;
			posY -= this.betweenPages;
			this.m_oScrollVerApi.scrollToY(posY);
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
			if (!this.file || !this.file.isValid)
				return;

			this.drawingPages = [];
			for (let i = 0, len = this.file.pages.length; i < len; i++)
			{
				this.drawingPages.push({ 
					X : 0,
					Y : 0,
					W : (this.file.pages[i].W * 96 * this.zoom / this.file.pages[i].Dpi) >> 0,
					H : (this.file.pages[i].H * 96 * this.zoom / this.file.pages[i].Dpi) >> 0,
				});
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
			this.canvas.style.cursor = cursor;
		};

		this.onMouseDown = function(e)
		{
			if (!oThis.file || !oThis.file.isValid())
				return;

			AscCommon.check_MouseDownEvent(e, true);
			AscCommon.global_mouseEvent.LockMouse();

			if (oThis.MouseHandObject)
			{
				if (AscCommon.global_keyboardEvent.CtrlKey)
				{
					var pageObject = oThis.getPageByCoords(AscCommon.global_mouseEvent.X - oThis.x, AscCommon.global_mouseEvent.Y - oThis.y);
					if (pageObject)
					{
						// links
						var pageLinks = oThis.pagesInfo.pages[pageObject.index];
						if (pageLinks.links)
						{
							for (var i = 0, len = pageLinks.links.length; i < len; i++)
							{
								if (pageObject.x >= pageLinks.links[i]["x"] && pageObject.x <= (pageLinks.links[i]["x"] + pageLinks.links[i]["w"]) &&
									pageObject.y >= pageLinks.links[i]["y"] && pageObject.y <= (pageLinks.links[i]["y"] + pageLinks.links[i]["h"]))
								{
									oThis.setCursorType("pointer");
									console.log(pageLinks.links[i]["link"]);
									return;
								}
							}
						}
					}
				}

				oThis.MouseHandObject.X = AscCommon.global_mouseEvent.X;
				oThis.MouseHandObject.Y = AscCommon.global_mouseEvent.Y;
				oThis.MouseHandObject.Active = true;
				oThis.MouseHandObject.ScrollX = oThis.scrollX;
				oThis.MouseHandObject.ScrollY = oThis.scrollY;
				oThis.setCursorType("grabbing");
				return;
			}

			// TODO: SELECT
		};
		
		this.onMouseUp = function(e)
		{
			if (!oThis.file || !oThis.file.isValid())
				return;

			AscCommon.check_MouseUpEvent(e);
			if (oThis.MouseHandObject && oThis.MouseHandObject.Active)
			{				
				oThis.MouseHandObject.Active = false;
				oThis.setCursorType("grab");
				return;
			}

			// TODO: SELECT
		};

		this.onMouseMove = function(e)
		{
			if (!oThis.file || !oThis.file.isValid())
				return;

			AscCommon.check_MouseMoveEvent(e);
			if (oThis.MouseHandObject)
			{
				if (oThis.MouseHandObject.Active)
				{
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
			}

			var pageObject = oThis.getPageByCoords(AscCommon.global_mouseEvent.X - oThis.x, AscCommon.global_mouseEvent.Y - oThis.y);
			if (pageObject)
			{
				// links
				var pageLinks = oThis.pagesInfo.pages[pageObject.index];
				if (pageLinks.links)
				{
					for (var i = 0, len = pageLinks.links.length; i < len; i++)
					{
						if (pageObject.x >= pageLinks.links[i]["x"] && pageObject.x <= (pageLinks.links[i]["x"] + pageLinks.links[i]["w"]) &&
							pageObject.y >= pageLinks.links[i]["y"] && pageObject.y <= (pageLinks.links[i]["y"] + pageLinks.links[i]["h"]))
						{
							oThis.setCursorType("pointer");
							return;
						}
					}
				}
			}

			if (oThis.MouseHandObject)
			{
				oThis.setCursorType("grab");
			}
			else
			{
				oThis.setCursorType("arrow");
			}
			
			// TODO: SELECT
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

			if (true === _ctrl)
			{
				if (e.preventDefault)
					e.preventDefault();
				else
					e.returnValue = false;

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

			if (e.preventDefault)
				e.preventDefault();
			else
				e.returnValue = false;

			return false;
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

		this._paint = function()
		{
			if (!this.file.isValid())
				return;

			this.canvas.width = this.canvas.width;
			let ctx = this.canvas.getContext("2d");
			ctx.strokeStyle = "#000000";
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

			for (let i = lStartPage; i <= lEndPage; i++)
			{
				// отрисовываем страницу
				let page = this.drawingPages[i];

				let w = (page.W * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
				let h = (page.H * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

				if (!this.file.cacheManager)
				{
					if (page.Image && ((page.Image.width != w) || (page.Image.height != h)))
						delete page.Image;
				}
				else
				{
					if (page.Image && ((page.Image.width < w) || (page.Image.height < h)))
					{
						if (this.file.cacheManager)
							this.file.cacheManager.unlock(page.Image);
						
						delete page.Image;	
					}
				}
				
				if (!page.Image)
					page.Image = this.file.getPage(i, w, h);

				let x = ((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - (w >> 1);
				let y = ((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

				if (page.Image)
				{
					ctx.drawImage(page.Image, 0, 0, w, h, x, y, w, h);
					this.pagesInfo.checkPage(i);
				}
				else
				{
					ctx.fillStyle = "#FFFFFF";
					ctx.fillRect(x, y, w, h);
				}
				ctx.strokeRect(x + lineW / 2, y + lineW / 2, w - lineW, h - lineW);

				this.pageDetector.addPage(i, x, y, w, h);
			}

			this.updateCurrentPage(this.pageDetector.getCurrentPage());
		};

		this.checkPageInfos = function()
		{
			if (!this.pagesInfo.isCheck)
				return false;

			this.pagesInfo.isCheck = false;
			if (this.startVisiblePage < 0 || this.endVisiblePage < 0)
				return false;

			for (var i = this.startVisiblePage; i <= this.endVisiblePage; i++)
			{
				var page = this.pagesInfo.pages[i];
				if (page.isReady && null === page.links)
				{
					page.links = this.file.getLinks(i);
				}
			}

			return true;
		};

		this.getPageByCoords = function(x, y)
		{
			if (this.startVisiblePage < 0 || this.endVisiblePage < 0)
				return null;

			for (var i = this.startVisiblePage; i <= this.endVisiblePage; i++)
			{
				var pageCoords = this.pageDetector.pages[i - this.startVisiblePage];
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
		};

		this.Copy = function(_text_format)
		{
			var ret = "<div>";
			if (_text_format && _text_format.Text)
				_text_format.Text = "";

			ret += "</div>";
			return ret;
		};

		this.findText = function(text, isMachingCase, isNext)
		{
			this.SearchResults.IsSearch = true;
			if (text === this.SearchResults.Text && isMachingCase === this.SearchResults.MachingCase)
			{
				if (this.SearchResults.Count === 0)
				{
					editor.WordControl.m_oDrawingDocument.CurrentSearchNavi = null;
					this.SearchResults.CurrentPage = -1;
					this.SearchResults.Current = -1;
					return;
				}

				// поиск совпал, просто делаем навигацию к нужному месту
				if (isNext)
				{
					if ((this.SearchResults.Current + 1) < this.SearchResults.Pages[this.SearchResults.CurrentPage].length)
					{
						// результат на этой же странице
						this.SearchResults.Current++;
					}
					else
					{
						var _pageFind = this.SearchResults.CurrentPage + 1;
						var _bIsFound = false;
						for (var i = _pageFind; i < this.PagesCount; i++)
						{
							if (0 < this.SearchResults.Pages[i].length)
							{
								this.SearchResults.Current = 0;
								this.SearchResults.CurrentPage = i;
								_bIsFound = true;
								break;
							}
						}
						if (!_bIsFound)
						{
							for (var i = 0; i < _pageFind; i++)
							{
								if (0 < this.SearchResults.Pages[i].length)
								{
									this.SearchResults.Current = 0;
									this.SearchResults.CurrentPage = i;
									_bIsFound = true;
									break;
								}
							}
						}
					}
				}
				else
				{
					if (this.SearchResults.Current > 0)
					{
						// результат на этой же странице
						this.SearchResults.Current--;
					}
					else
					{
						var _pageFind = this.SearchResults.CurrentPage - 1;
						var _bIsFound = false;
						for (var i = _pageFind; i >= 0; i--)
						{
							if (0 < this.SearchResults.Pages[i].length)
							{
								this.SearchResults.Current = this.SearchResults.Pages[i].length - 1;
								this.SearchResults.CurrentPage = i;
								_bIsFound = true;
								break;
							}
						}
						if (!_bIsFound)
						{
							for (var i = this.PagesCount - 1; i > _pageFind; i--)
							{
								if (0 < this.SearchResults.Pages[i].length)
								{
									this.SearchResults.Current = this.SearchResults.Pages[i].length - 1;
									this.SearchResults.CurrentPage = i;
									_bIsFound = true;
									break;
								}
							}
						}
					}
				}

				editor.WordControl.m_oDrawingDocument.CurrentSearchNavi =
					this.SearchResults.Pages[this.SearchResults.CurrentPage][this.SearchResults.Current];

				editor.WordControl.ToSearchResult();
				return;
			}
			// новый поиск
			for (var i = 0; i < this.PagesCount; i++)
			{
				this.SearchResults.Pages[i].splice(0,  this.SearchResults.Pages[i].length);
			}
			this.SearchResults.Count = 0;

			this.SearchResults.CurrentPage = -1;
			this.SearchResults.Current = -1;

			this.SearchResults.Text = text;
			this.SearchResults.MachingCase = isMachingCase;

			for (var i = 0; i < this.PagesCount; i++)
			{
				// TODO: Search
				// this.SearchPage2(i);
				this.SearchResults.Count += this.SearchResults.Pages[i].length;
			}

			if (this.SearchResults.Count == 0)
			{
				editor.WordControl.m_oDrawingDocument.CurrentSearchNavi = null;
				editor.WordControl.OnUpdateOverlay();
				return;
			}

			for (var i = 0; i < this.SearchResults.Pages.length; i++)
			{
				if (0 != this.SearchResults.Pages[i].length)
				{
					this.SearchResults.CurrentPage = i;
					this.SearchResults.Current = 0;

					break;
				}
			}

			editor.WordControl.m_oDrawingDocument.CurrentSearchNavi =
				this.SearchResults.Pages[this.SearchResults.CurrentPage][this.SearchResults.Current];

			editor.WordControl.ToSearchResult();
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
	CCurrentPageDetector.prototype.getCurrentPage = function()
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
