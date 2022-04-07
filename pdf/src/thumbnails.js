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

    // SKIN
    var PageStyle = {
        emptyColor : "#FFFFFF",

        numberColor : "#000000",
        numberFontSize : 10,
        numberFont : "Arial",
        numberFontOffset : 10,

        outlineColor : "#D9D9D9",
        outlineColorOffset : 0,
        outlineColorWidth : 1,

        hoverColor : "#BABABA",
        hoverColorOffset : 3,
        hoverColorWidth : 3,

        selectColor : "#888888",
        selectColorOffset : 3,
        selectColorWidth : 3,

        isDrawCurrentRect : true,
        drawCurrentColor : "#888888",
        drawCurrentWidth : 2
    };

    var ThumbnailsStyle = {
        backgroundColor : "#F1F1F1"
    };

    PageStyle.numberFontHeight = (function(){
        if (window["NATIVE_EDITOR_ENJINE"])
            return 7;
        var testCanvas = document.createElement("canvas");
        var w = 100;
        var h = 100;
        testCanvas.width = w;
        testCanvas.height = h;
        var ctx = testCanvas.getContext("2d");
        ctx.font = PageStyle.numberFont;
        ctx.fillStyle = "#FFFFFF";
        ctx.fillRect(0, 0, w, h);
        ctx.fillStyle = "#000000";
        ctx.font = PageStyle.numberFont;
        ctx.fillText("123456789", 0, h);
        var pixels = ctx.getImageData(0, 0, w, h).data;
        var index = 0;
        var indexLast = 4 * w * h;
        while (index < indexLast)
        {
            if (pixels[index] !== 255 || pixels[index + 1] !== 255 || pixels[index + 2] !== 255)
                break;
            index += 4;
        }
        return h - ((index / (4 * w)) >> 0);
    })();

    PageStyle.font = function()
    {
        var size = AscCommon.AscBrowser.convertToRetinaValue(this.numberFontSize, true);
        return "" + size + "px " + this.numberFont;
    };

    // LOGIC PAGE
    function CPage(w, h)
    {
        this.width = w;
        this.height = h;

        this.image = null;
    }

    CPage.prototype.draw = function(ctx, x, y, w, h)
    {
        if (null === this.image)
        {
            ctx.fillStyle = PageStyle.emptyColor;
            ctx.fillRect(x, y, w, h);
        }
        else
        {
            ctx.drawImage(this.image, x, y, w, h);
        }
        if (null !== PageStyle.outlineColor)
        {
            var lineW = Math.max(1, (PageStyle.outlineColorWidth * AscCommon.AscBrowser.retinaPixelRatio) >> 0);
            var offsetW = PageStyle.outlineColorOffset + 0.5 * lineW;
            
            ctx.lineWidth = lineW;
            ctx.strokeStyle = PageStyle.outlineColor;
            ctx.strokeRect(x - offsetW, y - offsetW, w + 2 * offsetW, h + 2 * offsetW);
        }        
    };

    // DRAWING PAGE
    function CDrawingPage(num, page)
    {
        this.page = page;
        this.pageRect = { x:0, y:0, w:0, h:0 };
        this.numRect = { x:0, y:0, w:0, h:0 };
        this.num = num;
    }

    CDrawingPage.prototype.draw = function(ctx, offsetV, doc)
    {
        this.page.draw(ctx, this.pageRect.x, this.pageRect.y - offsetV, this.pageRect.w, this.pageRect.h);

        var lineW, offsetW, color = undefined;
        if (this.num === doc.selectPage)
        {
            lineW = Math.max(1, (PageStyle.selectColorWidth * AscCommon.AscBrowser.retinaPixelRatio) >> 0);
            offsetW = PageStyle.selectColorOffset + 0.5 * lineW;
            color = PageStyle.selectColor;
        }
        else if (this.num === doc.hoverPage)
        {
            lineW = Math.max(1, (PageStyle.hoverColorWidth * AscCommon.AscBrowser.retinaPixelRatio) >> 0);
            offsetW = PageStyle.hoverColorOffset + 0.5 * lineW;
            color = PageStyle.hoverColor;
        }

        if (color)
        {
            ctx.lineWidth = lineW;
            ctx.strokeStyle = color;
            ctx.strokeRect(this.pageRect.x - offsetW, this.pageRect.y - offsetV - offsetW, this.pageRect.w + 2 * offsetW, this.pageRect.h + 2 * offsetW);
        }

        // currentRect
        var currentRect = null;
        if (PageStyle.isDrawCurrentRect && doc.selectPage === this.num)
            currentRect = doc.selectPageRect;
        if (currentRect)
        {
            var _x = currentRect.x;
            var _y = currentRect.y;
            var _r = currentRect.r;
            var _b = currentRect.b;

            if (_x < 0 || _x > 1 || _y < 0 || _y > 1 || _r < 0 || _r > 1 || _b < 0 || _b > 1)
                return;

            var pixX = (this.pageRect.x + _x * this.pageRect.w) >> 0;
            var pixY = (this.pageRect.y - offsetV + _y * this.pageRect.h) >> 0;
            var pixR = (this.pageRect.x + _r * this.pageRect.w) >> 0;
            var pixB = (this.pageRect.y - offsetV + _b * this.pageRect.h) >> 0;

            if (pixR <= pixX) return;
            if (pixB <= pixY) return;

            var lineW = Math.max(1, (PageStyle.drawCurrentWidth * AscCommon.AscBrowser.retinaPixelRatio) >> 0);
            var offsetW = 0.5 * lineW;

            ctx.lineWidth = lineW;
            ctx.strokeStyle = PageStyle.drawCurrentColor;
            ctx.strokeRect(pixX + offsetW, pixY + offsetW, pixR - pixX - 2 * offsetW, pixB - pixY - 2 * offsetW);
        }

        ctx.fillStyle = PageStyle.numberColor;
        ctx.fillText("" + (this.num + 1), this.numRect.x + this.numRect.w / 2, this.numRect.y + this.numRect.h - offsetV);
    };

    // BLOCK OF DRAWING PAGES
    function CBlock()
    {
        this.pages = [];
        this.top;
        this.bottom;
    }

    CBlock.prototype.getHeight = function(columnW, startOffset, betweenPages, zoom)
    {
        var maxPageHeight = 0;
        for (var i = 0, len = this.pages.length; i < len; i++)
        {
            if (this.pages[i].page.height > maxPageHeight)
                maxPageHeight = this.pages[i].page.height;
        }

        var blockHeight = (maxPageHeight * zoom) >> 0;
        var numberBlockH = AscCommon.AscBrowser.convertToRetinaValue(PageStyle.numberFontOffset + PageStyle.numberFontHeight, true);
        blockHeight += numberBlockH;

        var currentPosX = startOffset;
        for (var i = 0, len = this.pages.length; i < len; i++)
        {
            var drPage = this.pages[i];
            var pW = (drPage.page.width * zoom) >> 0;
            var pH = (drPage.page.height * zoom) >> 0;
            var curPageHeight = pH + PageStyle.numberFontOffset + PageStyle.numberFontHeight;

            drPage.pageRect.y = this.top + ((blockHeight - curPageHeight) >> 1);
            drPage.pageRect.h = pH;
            drPage.pageRect.x = currentPosX + ((columnW - pW) >> 1);
            drPage.pageRect.w = pW;

            drPage.numRect.y = (drPage.pageRect.y + drPage.pageRect.h);
            drPage.numRect.h = numberBlockH;
            drPage.numRect.x = drPage.pageRect.x;
            drPage.numRect.w = drPage.pageRect.w;

            currentPosX += (columnW + betweenPages);
        }

        return blockHeight;
    };

    CBlock.prototype.draw = function(ctx, offsetV, doc)
    {
        for (var i = 0, len = this.pages.length; i < len; i++)
        {
            this.pages[i].draw(ctx, offsetV, doc);
        }
    };

    // ГЛАВНЫЙ КЛАСС
    function CDocument(id)
    {
        this.id = id;
        this.viewer = null;
        this.isEnabled = true;

        this.coordsOffset = {x: 0, y: 0};

        this.pages = [];
        this.countPagesInBlock = 1;

        this.blocks = [];

        this.settings = {
            marginW : 20,
            marginH : 10,
            betweenW : 30,
            betweenH : 20
        };

        this.marginW = 20;
        this.marginH = 10;

        this.betweenW = 30;
        this.betweenH = 20;

        this.startBlock = -1;
        this.endBlock = -1;

        this.documentWidth = 0;
        this.documentHeight = 0;

        this.panelWidth = 0;
        this.panelHeight = 0;

        this.scrollY = 0;
        this.scrollMaxY = 0;

        this.minSizePage = 20;
        this.defaultPageW = 150;
        this.zoomMin = 1;
        this.zoomMax = 1; 
        this.zoom = 1;

        this.canvas = null;
        this.canvasOverlay = null;
        this.isRepaint = false;

        this.scrollWidth = 10;
        this.m_oScrollVerApi = null;

        this.selectPage = -1;
        this.selectPageRect = null;
        this.hoverPage = -1;

        this.handlers = {};

        this.createComponents();
    }

    // INTERFACE
    CDocument.prototype.repaint = function()
    {
        this.isRepaint = true;
    };
    CDocument.prototype.setZoom = function(zoom)
    {
        this.zoom = this.zoomMin + zoom * (this.zoomMax - this.zoomMin);
        this.resize(false);
    };
    CDocument.prototype.resize = function(isZoomUpdated)
    {
        this._resize(isZoomUpdated);
    };
    CDocument.prototype.setIsDrawCurrentRect = function(isDrawCurrentRect)
    {
        PageStyle.isDrawCurrentRect = isDrawCurrentRect;
        this.repaint();
    };
    CDocument.prototype.setEnabled = function(isEnabled)
    {
        this.isEnabled = isEnabled;
        if (this.isEnabled)
            this.repaint();
    };
    CDocument.prototype.registerEvent = function(name, handler)
    {
        if (this.handlers[name] === undefined)
            this.handlers[name] = [];
        this.handlers[name].push(handler);
    };

    // HTML/INTERFACE
    CDocument.prototype.createComponents = function()
    {
        this.updateSkin();

        var parent = document.getElementById(this.id);
        var elements = "";
        elements += "<canvas id=\"id_viewer_th\" class=\"block_elem\" style=\"left:0px;top:0px;width:100;height:100;\"></canvas>";
        elements += "<canvas id=\"id_overlay_th\" class=\"block_elem\" style=\"left:0px;top:0px;width:100;height:100;\"></canvas>";
        elements += "<div id=\"id_vertical_scroll_th\" class=\"block_elem\" style=\"display:none;left:0px;top:0px;width:0px;height:0px;\"></div>";
    
        parent.style.backgroundColor = ThumbnailsStyle.backgroundColor;
        parent.innerHTML = elements;

        this.canvas = document.getElementById("id_viewer_th");
        this.canvas.backgroundColor = ThumbnailsStyle.backgroundColor;

        this.canvasOverlay = document.getElementById("id_overlay_th");
        this.canvasOverlay.style.pointerEvents = "none";

        parent.onmousewheel = this.onMouseWhell.bind(this);
        if (parent.addEventListener)
			parent.addEventListener("DOMMouseScroll", this.onMouseWhell.bind(this), false);

        AscCommon.addMouseEvent(this.canvas, "down", this.onMouseDown.bind(this));
        AscCommon.addMouseEvent(this.canvas, "move", this.onMouseMove.bind(this));
        AscCommon.addMouseEvent(this.canvas, "up", this.onMouseUp.bind(this));
    };

    CDocument.prototype.sendEvent = function()
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

    // SCROLL    
    CDocument.prototype.CreateScrollSettings = function()
    {
        var settings = new AscCommon.ScrollSettings();
        settings.screenW = this.panelWidth;
        settings.screenH = this.panelHeight;
        settings.vscrollStep = 45;
        settings.hscrollStep = 45;

        //settings.isNeedInvertOnActive = GlobalSkin.isNeedInvertOnActive;
        settings.showArrows = false;
        settings.cornerRadius = 1;
        settings.slimScroll = true;

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

    CDocument.prototype.scrollVertical = function(pos, maxPos)
    {
        this.scrollY = pos;
        this.scrollMaxY = maxPos;
        this.calculateVisibleBlocks();
        this.repaint();
    };

    CDocument.prototype.updateScroll = function(scrollV)
    {
        scrollV.style.display = (this.documentHeight > this.panelHeight) ? "block" : "none";

        var settings = this.CreateScrollSettings();
        settings.isHorizontalScroll = false;
        settings.isVerticalScroll = true;
        settings.contentH = this.documentHeight;
        if (this.m_oScrollVerApi)
            this.m_oScrollVerApi.Repos(settings, undefined, true);
        else
        {
            this.m_oScrollVerApi = new AscCommon.ScrollObject("id_vertical_scroll_th", settings);

            this.m_oScrollVerApi.onLockMouse  = function(evt) {
                AscCommon.check_MouseDownEvent(evt, true);
                AscCommon.global_mouseEvent.LockMouse();
            };
            this.m_oScrollVerApi.offLockMouse = function(evt) {
                AscCommon.check_MouseUpEvent(evt);
            };
            var _t = this;
            this.m_oScrollVerApi.bind("scrollvertical", function(evt) {
                _t.scrollVertical(evt.scrollD, evt.maxScrollY);
            });
        }

        this.scrollMaxY = this.m_oScrollVerApi.getMaxScrolledY();
        if (this.scrollY >= this.scrollMaxY)
            this.scrollY = this.scrollMaxY;
    };

    // очередь задач - нужно ли перерисоваться и/или перерисовать страницу
    CDocument.prototype.checkTasks = function(isViewerTask)
    {
        var isNeedTasks = false;
        if (!this.isEnabled)
            return isNeedTasks;

        if (!isViewerTask && -1 != this.startBlock)
        {
            // смотрим, какие страницы нужно перерисовать. 
            // делаем это по одной, так как задачи вьюера важнее
            var needPage = null;
            var drPage, block;
            for (var blockNum = this.startBlock; (blockNum <= this.endBlock) && !needPage; blockNum++)
            {
                block = this.blocks[blockNum];

                for (var pageNum = 0, pagesCount = block.pages.length; pageNum < pagesCount; pageNum++)
                {
                    drPage = block.pages[pageNum];
                    if (drPage.page.image === null ||
                        (drPage.page.image.requestWidth != drPage.pageRect.w || drPage.page.image.requestHeight != drPage.pageRect.h))
                    {
                        needPage = drPage;
                        break;
                    }
                }            
            }

            if (needPage)
            {
                isNeedTasks = true;
                needPage.page.image = this.viewer.file.getPage(needPage.num, needPage.pageRect.w, needPage.pageRect.h, undefined, this.viewer.Api.isDarkMode ? 0x3A3A3A : 0xFFFFFF);
                this.isRepaint = true;
            }
        }

        // проверяем, нужна ли перерисовка
        if (this.isRepaint)
        {
            this._paint();
            this.isRepaint = false;
        }

        return isNeedTasks;
    };

    CDocument.prototype.updateCurrentPage = function(pageObject)
    {
        this.selectPageRect = pageObject;
        if (this.selectPage != pageObject.num)
        {
            this.selectPage = pageObject.num;
            var drPage = this.getDrawingPage(this.selectPage);
            if (!drPage)
                return;

            // или подскролливаем, или просто перерисовываем            
            if (drPage.pageRect.y < this.scrollY)
                this.m_oScrollVerApi.scrollToY(drPage.pageRect.y - this.betweenH);
            else
            {
                var b = drPage.pageRect.y + drPage.pageRect.h + drPage.numRect.h;
                if (b > (this.scrollY + this.panelHeight))
                    this.m_oScrollVerApi.scrollToY(b - this.panelHeight + this.betweenH);
                else
                    this.repaint();
            }
        }
        else if (PageStyle.isDrawCurrentRect)
            this.repaint();
    };

    // сама отрисовка
    CDocument.prototype._paint = function()
    {
        this.canvas.width = this.canvas.width;
        var ctx = this.canvas.getContext("2d");
        ctx.fillStyle = ThumbnailsStyle.backgroundColor;
        ctx.fillRect(0, 0, this.panelWidth, this.panelHeight);

        if (-1 == this.startBlock)
            return;
        
        ctx.font = PageStyle.font();
        ctx.textAlign = "center";
        for (var block = this.startBlock; block <= this.endBlock; block++)
        {
            this.blocks[block].draw(ctx, this.scrollY >> 0, this);
        }
    };

    CDocument.prototype.init = function()
    {
        this.pages = [];
        if (this.viewer.file && this.viewer.file.isValid())
        {
            var pages = this.viewer.file.pages;
            let koef = 1;
            for (let i = 0, len = pages.length; i < len; i++)
            {
                koef = 1;
                if (pages[i].Dpi > 1)
                    koef = 100 / pages[i].Dpi;
                this.pages.push(new CPage(koef * pages[i].W, koef * pages[i].H));
            }
        }
        
        this.resize();
    };

    CDocument.prototype._resize = function(isZoomUpdated)
    {
        var element = document.getElementById(this.id);

        if (0 === element.offsetWidth)
            return;

        // размер панели
        this.panelWidth = element.offsetWidth;
        this.panelHeight = element.offsetHeight;

        this.canvas.style.width = this.panelWidth + "px";
        this.canvas.style.height = this.panelHeight + "px";
        this.canvasOverlay.style.width = this.panelWidth + "px";
        this.canvasOverlay.style.height = this.panelHeight + "px";
        
        AscCommon.calculateCanvasSize(this.canvas);
        AscCommon.calculateCanvasSize(this.canvasOverlay);

        var canvasBounds = this.canvas.getBoundingClientRect();
        this.coordsOffset.x = canvasBounds ? canvasBounds.left : 0;
        this.coordsOffset.y = canvasBounds ? canvasBounds.top : 0;

        var scrollV = document.getElementById("id_vertical_scroll_th");
        scrollV.style.display = "none";
        scrollV.style.left = this.panelWidth - this.scrollWidth + "px";
        scrollV.style.top = "0px";
        scrollV.style.width = this.scrollWidth + "px";
        scrollV.style.height = this.panelHeight + "px";

        this.panelWidth = AscCommon.AscBrowser.convertToRetinaValue(this.panelWidth, true);
        this.panelHeight = AscCommon.AscBrowser.convertToRetinaValue(this.panelHeight, true);

        this.marginW = AscCommon.AscBrowser.convertToRetinaValue(this.settings.marginW, true);
        this.marginH = AscCommon.AscBrowser.convertToRetinaValue(this.settings.marginH, true);
        this.betweenW = AscCommon.AscBrowser.convertToRetinaValue(this.settings.betweenW, true);
        this.betweenH = AscCommon.AscBrowser.convertToRetinaValue(this.settings.betweenH, true);

        if (this.pages.length == 0)
            return;

        // делим страницы на колонки одинаковой ширины (по максимальной странице)

        var pageWidthMax = this.getMaxPageWidth();
        var sizeMax = Math.max(pageWidthMax, this.getMaxPageHeight());

        // максимальные/минимальные зумы
        this.zoomMin = this.minSizePage / sizeMax;
        this.zoomMax = (this.panelWidth - (2 * this.marginW)) / pageWidthMax;
        
        if (this.defaultPageW != 0)
        {
            // зум "по умолчанию"
            this.zoom = AscCommon.AscBrowser.convertToRetinaValue(this.defaultPageW, true) / pageWidthMax;

            if (0 != this.panelWidth)
                this.defaultPageW = 0;
        }

        // корректировка зумов
        if (this.zoomMax < this.zoomMin)
            this.zoomMax = this.zoomMin;
        if (this.zoom < this.zoomMin)
            this.zoom = this.zoomMin;
        if (this.zoom > this.zoomMax)
            this.zoom = this.zoomMax;

        if (isZoomUpdated !== false)
        {
            var interfaceZoom = (this.zoomMax - this.zoomMin) < 0.001 ? 0 : (this.zoom - this.zoomMin) / (this.zoomMax - this.zoomMin);
            this.sendEvent("onZoomChanged", interfaceZoom);
        }

        // смотрим, сколько столбцов влезает
        // уравнение:
        // (pageWidthMax * this.zoom) * x + this.betweenW * (x - 1) = this.panelWidth - 2 * this.marginW;
        var blockW = (pageWidthMax * this.zoom) >> 0;
        this.countPagesInBlock = (this.panelWidth - 2 * this.marginW + this.betweenW) / (blockW + this.betweenW);
        this.countPagesInBlock >>= 0;
        if (this.countPagesInBlock < 1)
            this.countPagesInBlock = 1;

        if (this.countPagesInBlock > this.pages.length)
            this.countPagesInBlock = this.pages.length;

        this.documentWidth = this.countPagesInBlock * blockW + 2 * this.marginW + this.betweenW * (this.countPagesInBlock - 1);

        // теперь набиваем блоки
        this.blocks = [];
        var blocksCount = 0;
        var countInCurrentBlock = this.countPagesInBlock;
        for (let i = 0, len = this.pages.length; i < len; i++)
        {            
            if (countInCurrentBlock == this.countPagesInBlock)
            {
                this.blocks[blocksCount++] = new CBlock();
                countInCurrentBlock = 0;
            }
            this.blocks[blocksCount - 1].pages.push(new CDrawingPage(i, this.pages[i]));
            ++countInCurrentBlock;
        }

        // теперь считаем позиции страниц в блоке (координаты сквозные)
        var blockTop = this.betweenH;
        var startOffsetX = this.marginW + ((this.panelWidth - this.documentWidth) >> 1);
        for (let i = 0, len = this.blocks.length; i < len; i++)
        {
            var block = this.blocks[i];
            block.top = blockTop;
            block.bottom = block.top + block.getHeight(blockW, startOffsetX, this.betweenW, this.zoom);
            blockTop = block.bottom + this.betweenH;
        }

        this.documentHeight = blockTop;

        this.updateScroll(scrollV);
        this.calculateVisibleBlocks();
        this.repaint();
    };

    CDocument.prototype.calculateVisibleBlocks = function()
    {
        this.startBlock = -1;
        this.endBlock = -1;
        var blocksCount = this.blocks.length;
        var block;
        for (var i = 0; i < blocksCount; i++)
        {
            block = this.blocks[i];
            if (block.bottom > this.scrollY)
            {
                // первый видимый блок!
                this.startBlock = i;
                break;
            }
            else
            {
                // выкидываем страницу из кэша
                for (var pageNum = 0, pagesCount = block.pages.length; pageNum < pagesCount; pageNum++)
                {
                    this.pages[block.pages[pageNum].num].image = null;
                }
            }
        }

        if (this.startBlock != -1)
        {
            for (var i = this.startBlock; i < blocksCount; i++)
            {
                block = this.blocks[i];
                if (block.top > (this.scrollY + this.panelHeight))
                {
                    // уже невидимый блок!
                    this.endBlock = i - 1;
                    break;
                }
            }
        }

        // проверяем - могли дойти до конца
        if (this.startBlock >= 0 && this.endBlock == -1)
            this.endBlock = blocksCount - 1;

        for (var i = this.endBlock + 1; i < blocksCount; i++)
        {
            block = this.blocks[i];
            // выкидываем страницу из кэша
            for (var pageNum = 0, pagesCount = block.pages.length; pageNum < pagesCount; pageNum++)
            {
                this.pages[block.pages[pageNum].num].image = null;
            }
        }        
    };

    CDocument.prototype.getMaxPageWidth = function()
    {
        var size = 0, page = null;
        for (var i = 0, count = this.pages.length; i < count; i++)
        {
            page = this.pages[i];
            if (size < page.width)
                size = page.width;
        }
        return size;
    };
    CDocument.prototype.getMaxPageHeight = function()
    {
        var size = 0, page = null;
        for (var i = 0, count = this.pages.length; i < count; i++)
        {
            page = this.pages[i];
            if (size < page.height)
                size = page.height;
        }
        return size;
    };

    CDocument.prototype.getPageByCoords = function(x, y)
    {
        // тут ТОЛЬКО для попадания. поэтому смотрим только видимые блоки
        if (-1 === this.startBlock || -1 === this.endBlock)
            return null;

        x -= this.coordsOffset.x;
        y -= this.coordsOffset.y;

        x = AscCommon.AscBrowser.convertToRetinaValue(x, true);
        y = AscCommon.AscBrowser.convertToRetinaValue(y, true);

        y += this.scrollY;

        var pages = null;
        for (var block = this.startBlock; block <= this.endBlock; block++)
        {
            if (y >= this.blocks[block].top && y <= this.blocks[block].bottom)
            {
                pages = this.blocks[block].pages;
                break;
            }
        }

        if (!pages)
            return null;

        var drPage;
        for (var pageNum = 0, count = pages.length; pageNum < count; pageNum++)
        {
            drPage = pages[pageNum];
            if (x >= drPage.pageRect.x && x <= (drPage.pageRect.x + drPage.pageRect.w) &&
                y >= drPage.pageRect.y && y <= (drPage.pageRect.y + drPage.pageRect.h))
            {
                return drPage;
            }
        }

        return null;
    };

    CDocument.prototype.getDrawingPage = function(pageNum)
    {
        if (pageNum < 0 || pageNum >= this.pages.length)
            return null;

        var block = (pageNum / this.countPagesInBlock) >> 0;
        if (!this.blocks[block])
            return null;

        var pageInBlock = pageNum - block * this.countPagesInBlock;
        return this.blocks[block].pages[pageInBlock];
    };

    // UI-EVENTS
    CDocument.prototype.onMouseDown = function(e)
    {
        AscCommon.check_MouseDownEvent(e, true);
        AscCommon.global_mouseEvent.LockMouse();
        this.viewer.isFocusOnThumbnails = true;
        
        var drPage = this.getPageByCoords(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
        if (drPage && drPage.num !== this.selectPage)
        {
            this.viewer.navigateToPage(drPage.num);
        }

        AscCommon.stopEvent(e);
        return false;
    };
    
    CDocument.prototype.onMouseUp = function(e)
    {
        AscCommon.check_MouseUpEvent(e);
        if (e && e.preventDefault)
            e.preventDefault();
        return false;
    };

    CDocument.prototype.onMouseMove = function(e)
    {
        AscCommon.check_MouseMoveEvent(e);

        if (AscCommon.global_mouseEvent.IsLocked &&
            this.canvas != AscCommon.global_mouseEvent.Sender)
        {
            return;
        }

        if (!AscCommon.global_mouseEvent.IsLocked)
        {
            var drPage = this.getPageByCoords(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
            var hoverNum = drPage ? drPage.num : -1;
            if (hoverNum !== this.hoverPage)
            {
                this.hoverPage = hoverNum;
                this._paint();
            }
        }

        if (e && e.preventDefault)
            e.preventDefault();
        return false;
    };

    CDocument.prototype.onMouseWhell = function(e)
    {
        AscCommon.stopEvent(e);
        if (this.scrollMaxY == 0)
            return false;

        var _ctrl = false;
        if (e.metaKey !== undefined)
            _ctrl = e.ctrlKey || e.metaKey;
        else
            _ctrl = e.ctrlKey;

        if (true === _ctrl)
            return false;

        var delta  = 0;
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

        delta = delta >> 0;
        if (0 != delta)
            this.m_oScrollVerApi.scrollBy(0, delta, false);

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

        this.onMouseMove(_e);
        // ------------------------------------------------------
        return false;
    };

    CDocument.prototype.updateSkin = function()
    {
        ThumbnailsStyle.backgroundColor = AscCommon.GlobalSkin.BackgroundColorThumbnails;
        PageStyle.hoverColor = AscCommon.GlobalSkin.ThumbnailsPageOutlineHover;
        PageStyle.selectColor = AscCommon.GlobalSkin.ThumbnailsPageOutlineActive;
        PageStyle.numberColor = AscCommon.GlobalSkin.ThumbnailsPageNumberText;

        if (this.canvas)
            this.canvas.style.backgroundColor = ThumbnailsStyle.backgroundColor;

        this.resize();
    };

    CDocument.prototype.checkPageEmptyStyle = function()
    {
        PageStyle.emptyColor = "#FFFFFF";
        if (this.viewer)
        {
            var backColor = this.viewer.Api.getPageBackgroundColor();
            PageStyle.emptyColor = "#" + backColor[0].toString(16) + backColor[1].toString(16) + backColor[2].toString(16);
        }
    }

    CDocument.prototype.clearCachePages = function()
    {
        this.checkPageEmptyStyle();

        for (var blockNum = 0, blocksCount = this.blocks.length; blockNum < blocksCount; blockNum++)
        {
            block = this.blocks[blockNum];

            for (var pageNum = 0, pagesCount = block.pages.length; pageNum < pagesCount; pageNum++)
            {
                drPage = block.pages[pageNum];
                if (drPage.page.image)
                {
                    drPage.page.image = null;
                }
            }
        }
    };

    // export
    AscCommon.ThumbnailsControl = CDocument;
    AscCommon["ThumbnailsControl"] = AscCommon.ThumbnailsControl;
    var prot = AscCommon["ThumbnailsControl"].prototype;
    prot["repaint"] = prot.repaint;
    prot["setZoom"] = prot.setZoom;
    prot["resize"] = prot.resize;
    prot["setEnabled"] = prot.setEnabled;
})();
