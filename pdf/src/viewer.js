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
		this.Id = null;
		if ((AscCommon.g_oIdCounter.m_bLoad || AscCommon.History.CanAddChanges())) {
			this.Id = AscCommon.g_oIdCounter.Get_NewId();
			AscCommon.g_oTableId.Add(this, this.Id);
		}
		
		this.annotsContentChanges = new AscCommon.CContentChanges(); // список изменений(добавление/удаление элементов)
        this.fieldsContentChanges = new AscCommon.CContentChanges(); // список изменений(добавление/удаление элементов)
        this.drawingsContentChanges = new AscCommon.CContentChanges(); // список изменений(добавление/удаление элементов)

		this.isPainted				= false;
		this.links					= null;
		this.fields					= [];
		this.annots					= [];
		this.drawings				= [];
		this.needRedrawForms		= true;
		this.needRedrawDrawings		= true;
		this.needRedrawAnnots		= true;

		this.deleteLock = null;
		this.rotateLock = null;
		this.editPageLock = null;

		this.SetLocks(new AscPDF.PropLocker(this.GetId()), new AscPDF.PropLocker(this.GetId()), new AscPDF.PropLocker(this.GetId()));
	}
	AscFormat.InitClass(CPageInfo, AscFormat.CBaseNoIdObject, AscDFH.historyitem_type_Pdf_Page);
	CPageInfo.prototype.constructor = CPageInfo;

	Object.defineProperties(CPageInfo.prototype, {
		PageNum: {
			get: function () {
				return this.GetIndex();
			}
		}
	});

	CPageInfo.prototype.RedrawDrawings = function() {
		let oViewer = Asc.editor.getDocumentRenderer();
		let _t = this;

		let nIdx = _t.GetIndex();
		function setRedrawPageOnRepaint() {
            _t.needRedrawDrawings = true;
			nIdx != -1 && oViewer.thumbnails && oViewer.thumbnails._repaintPage(nIdx);
        }

        oViewer.paint(setRedrawPageOnRepaint);
	};
	CPageInfo.prototype.RedrawForms = function() {
		let oViewer = Asc.editor.getDocumentRenderer();
		let _t = this;

		let nIdx = _t.GetIndex();
		function setRedrawPageOnRepaint() {
            _t.needRedrawForms = true;
			nIdx != -1 && oViewer.thumbnails && oViewer.thumbnails._repaintPage(nIdx);
        }

        oViewer.paint(setRedrawPageOnRepaint);
	};
	CPageInfo.prototype.RedrawAnnots = function(isTextMarkup) {
		let oViewer = Asc.editor.getDocumentRenderer();
		let _t = this;

		let nIdx = _t.GetIndex();
		function setRedrawPageOnRepaint() {
			if (isTextMarkup) {
				_t.needRedrawMarkups = true;
			}
			else {
				_t.needRedrawAnnots = true;
			}

			nIdx != -1 && oViewer.thumbnails && oViewer.thumbnails._repaintPage(nIdx);
        }

        oViewer.paint(setRedrawPageOnRepaint);
	};

	CPageInfo.prototype.GetDocument = function() {
		return Asc.editor.getPDFDoc();
	};
	CPageInfo.prototype.Clear_ContentChanges = function() {
        this.annotsContentChanges.Clear();
        this.fieldsContentChanges.Clear();
        this.drawingsContentChanges.Clear();
    };
    CPageInfo.prototype.Add_ContentChanges = function(Changes) {
        let oChange = Changes.m_pData.Data;
        
        switch (oChange.Type) {
            case AscDFH.historyitem_PDF_Document_AnnotsContent:
                this.annotsContentChanges.Add(Changes);
                break;
            case AscDFH.historyitem_PDF_Document_FieldsContent:
                this.fieldsContentChanges.Add(Changes);
                break;
            case AscDFH.historyitem_PDF_Document_DrawingsContent:
                this.drawingsContentChanges.Add(Changes);
                break;
        }
    };
    CPageInfo.prototype.Refresh_ContentChanges = function() {
        this.annotsContentChanges.Refresh();
        this.fieldsContentChanges.Refresh();
        this.drawingsContentChanges.Refresh();
    };
	CPageInfo.prototype.AddDrawing = function(oDrawing, nPos) {
		if (nPos == undefined) {
            nPos = this.drawings.length;
        }

		this.drawings.splice(nPos, 0, oDrawing);
        oDrawing.SetParentPage(this);

		AscCommon.History.Add(new CChangesPDFDocumentDrawingsContent(this, nPos, [oDrawing], true));

		this.RedrawDrawings();
	};
	CPageInfo.prototype.RemoveDrawing = function(sId) {
        let oDrawing = this.drawings.find(function(drawing) {
            return drawing.GetId() === sId;
        });

        if (!oDrawing)
            return;

        let nPos = this.drawings.indexOf(oDrawing);
        this.drawings.splice(nPos, 1);
        
        AscCommon.History.Add(new CChangesPDFDocumentDrawingsContent(this, nPos, [oDrawing], false));
		this.RedrawDrawings();
	};
	CPageInfo.prototype.AddAnnot = function(oAnnot, nPos) {
		if (nPos == undefined) {
            nPos = this.annots.length;
        }

        this.annots.splice(nPos, 0, oAnnot);
        oAnnot.SetParentPage(this);

        AscCommon.History.Add(new CChangesPDFDocumentAnnotsContent(this, nPos, [oAnnot], true));
		this.RedrawAnnots();
	};
	CPageInfo.prototype.RemoveAnnot = function(sId) {
		let oAnnot = this.annots.find(function(annot) {
            return annot.GetId() === sId;
        });

        if (!oAnnot)
            return;

        let nPos = this.annots.indexOf(oAnnot);
        this.annots.splice(nPos, 1);
        
        AscCommon.History.Add(new CChangesPDFDocumentAnnotsContent(this, nPos, [oAnnot], false));
		this.RedrawAnnots(oAnnot.IsTextMarkup());
	};
	CPageInfo.prototype.AddField = function(oField, nPos) {
		if (nPos == undefined) {
            nPos = this.fields.length;
        }

		this.fields.splice(nPos, 0, oField);
		oField.SetParentPage(this);

        AscCommon.History.Add(new CChangesPDFDocumentFieldsContent(this, nPos, [oField], true));
		this.RedrawForms();
	};
	CPageInfo.prototype.RemoveField = function(sId) {
		let oDoc = this.GetDocument();
		let oField = this.fields.find(function(field) {
            return field.GetId() === sId;
        });

        if (!oField)
            return;

        let nPos = this.fields.indexOf(oField);
        this.fields.splice(nPos, 1);
        
        AscCommon.History.Add(new CChangesPDFDocumentFieldsContent(this, nPos, [oField], false));

		// удаляем из родителя
        let oParent = oField.GetParent();
        if (oParent) {
            oParent.RemoveKid(oField);
            oDoc.CheckParentForm(oParent); // проверяем родителя
        }

		this.RedrawForms();
	};
	CPageInfo.prototype.SetLocks = function(deleteLock, rotateLock, editPageLock) {
        this.deleteLock = deleteLock;
        this.rotateLock = rotateLock;
        this.editPageLock = editPageLock;
       	AscCommon.History.Add(new CChangesPDFDocumentPageLocks(this, deleteLock, rotateLock, editPageLock));
    };
	CPageInfo.prototype.IsLocked = function() {
		return this.IsDeleteLock() || this.IsRotateLock() || this.IsEditPageLock();
	}
	CPageInfo.prototype.IsDeleteLock = function() {
		return false == [AscCommon.c_oAscLockTypes.kLockTypeNone, AscCommon.c_oAscLockTypes.kLockTypeMine].includes(this.deleteLock.Lock.Get_Type());
	};
	CPageInfo.prototype.IsRotateLock = function() {
		return false == [AscCommon.c_oAscLockTypes.kLockTypeNone, AscCommon.c_oAscLockTypes.kLockTypeMine].includes(this.rotateLock.Lock.Get_Type());
	};
	CPageInfo.prototype.IsEditPageLock = function() {
		return false == [AscCommon.c_oAscLockTypes.kLockTypeNone, AscCommon.c_oAscLockTypes.kLockTypeMine].includes(this.editPageLock.Lock.Get_Type());
	};
	CPageInfo.prototype.GetIndex = function() {
		let oViewer = Asc.editor.getDocumentRenderer();
		let oDocPages = oViewer.pagesInfo;

		return oDocPages.pages.indexOf(this);
    };
	CPageInfo.prototype.SetRotate = function(nAngle) {
		let oDoc		= this.GetDocument();
		let oViewer     = oDoc.Viewer;
		let oFile       = oViewer.file;
		let nIndex		= this.GetIndex();

		if (oFile.pages[nIndex].Rotate == nAngle) {
			return;
		}

        oDoc.History.Add(new CChangesPDFDocumentRotatePage(this, oFile.pages[nIndex].Rotate, nAngle));
		oFile.pages[nIndex].Rotate = nAngle;
	};
	CPageInfo.prototype.GetRotate = function() {
		let oDoc		= this.GetDocument();
		let oFile       = oDoc.Viewer.file;
		let nIndex		= this.GetIndex();

		return oFile.pages[nIndex].Rotate;
	};
	CPageInfo.prototype.SetRecognized = function(isRecognized) {
		let oDoc		= this.GetDocument();
		let oViewer     = oDoc.Viewer;
		let oFile       = oViewer.file;
		let nIndex		= this.GetIndex();

		if (oFile.pages[nIndex].isRecognized == isRecognized) {
			return;
		}

        oDoc.History.Add(new CChangesPDFDocumentRecognizePage(this, oFile.pages[nIndex].isRecognized, isRecognized));
		oFile.pages[nIndex].isRecognized = isRecognized;
		delete oViewer.drawingPages[nIndex].Image;
	};
	CPageInfo.prototype.IsRecognized = function() {
		let oDoc		= this.GetDocument();
		let oFile       = oDoc.Viewer.file;
		let nIndex		= this.GetIndex();

		return oFile.pages[nIndex].Rotate;
	};
	CPageInfo.prototype.SetPosition = function(nNewPos) {
		let nCurPos = this.GetIndex();
		if (nCurPos === nNewPos) return false;
	
		let oDoc 		= this.GetDocument();
        let aFilePages  = oDoc.Viewer.file.pages;
        let aPagesInfo  = oDoc.Viewer.pagesInfo.pages;

        if (nNewPos < 0 || nNewPos >= aFilePages.length) {
            return false;
        }
    
        oDoc.History.Add(new CChangesPDFDocumentMovePage(this, nCurPos, nNewPos));

        let oMovedFilePage = aFilePages.splice(nCurPos, 1)[0];
        let oMovedPageInfo = aPagesInfo.splice(nCurPos, 1)[0];
        
        aFilePages.splice(nNewPos, 0, oMovedFilePage);
        aPagesInfo.splice(nNewPos, 0, oMovedPageInfo);

        return true;
	};
	CPageInfo.prototype.Is_Inline = function(){};

	function PropLocker(objectId) {
		this.objectId = null;
		this.Lock = new AscCommon.CLock();
		this.Id = AscCommon.g_oIdCounter.Get_NewId();
		AscCommon.g_oTableId.Add(this, this.Id);

		if (typeof objectId === "string") {
			this.setObjectId(objectId);
		}
	}

	PropLocker.prototype = {

		getObjectType: function() {
			return AscDFH.historyitem_type_Pdf_PropLocker;
		},

		setObjectId: function(id) {
			AscCommon.History.Add(new CChangesPDFPropLockerObjectId(this, this.objectId, id));
			this.objectId = id;
		},

		Get_Id: function() {
			return this.Id;
		},
		GetId: function() {
			return this.Get_Id();
		},

		Write_ToBinary2: function(w) {
			w.WriteLong(AscDFH.historyitem_type_Pdf_PropLocker);
			w.WriteString2(this.Id);
		},

		Read_FromBinary2: function(r) {
			this.Id = r.GetString2();
		},

		Refresh_RecalcData: function()
		{}
	};

	
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

		this.isFocusOnThumbnails = false;
		this.isDocumentContentReady = false;

		this.doc = new AscPDF.CPDFDoc(this);
		AscCommon.History.Document = this.doc;

		this.drawingDocument		= Asc.editor.WordControl.m_oDrawingDocument;
		this.DrawingObjects			= new AscPDF.CGraphicObjects(this.doc, this.drawingDocument, this.Api);
		this.doc.DrawingObjects		= this.DrawingObjects;
		this.doc.DrawingDocument	= this.drawingDocument;
		Asc.editor.WordControl.m_oLogicDocument = this.doc;
		Asc.editor.WordControl.m_oDrawingDocument.m_oLogicDocument = this.doc;
		this.touchManager = null;

		this.isXP = ((AscCommon.AscBrowser.userAgent.indexOf("windowsxp") > -1) || (AscCommon.AscBrowser.userAgent.indexOf("chrome/49") > -1)) ? true : false;
		if (!this.isXP && AscCommon.AscBrowser.isIE && !AscCommon.AscBrowser.isIeEdge)
			this.isXP = true;

		if (this.isXP)
		{
			AscCommon.g_oHtmlCursor.register(AscCommon.Cursors.Grab, "7 8", "pointer");
			AscCommon.g_oHtmlCursor.register(AscCommon.Cursors.Grabbing, "6 6", "pointer");
		}

		var oThis = this;

		this.onRepaintFormsCallbacks = [];
		this.onRepaintAnnotsCallbacks = [];

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

			if (this.disabledPaintOnScroll != true) {
				this._paint();
				this.onUpdateOverlay();
			}
		};
		this.UpdateDrDocDrawingPages = function() {
			let oThis = this;
			let xCenter = this.width >> 1;
			let yPos = this.scrollY >> 0;
			if (this.documentWidth > this.width)
				xCenter = (this.documentWidth >> 1) - (this.scrollX) >> 0;

			this.getPDFDoc().GetDrawingDocument().m_arrPages = this.drawingPages.map(function(page, index) {
				
				let w = page.W;
				let h = page.H;

				let x = xCenter - (w >> 1) >> 0;
				let y = page.Y - yPos >> 0;

				if (oThis.isLandscapePage(index)) {
					w = page.H;
					h = page.W;

					x = xCenter - (w >> 1) >> 0;
					y = page.Y - yPos >> 0;
				}

				return {
					width_mm: w / oThis.zoom * g_dKoef_pix_to_mm,
					height_mm: h / oThis.zoom * g_dKoef_pix_to_mm,
					drawingPage: {
						left: x,
						top: y,
						right: x + w,
						bottom: y + h
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

			if (this.disabledPaintOnScroll != true) {
				this._paint();
				this.onUpdateOverlay();
			}
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

			let oThumbnails = this.thumbnails;
			for (var i = 0, len = pages.length; i < len; i++)
			{
				oThumbnails && oThumbnails._repaintPage(pages[i]);
				if (pages[i] >= this.startVisiblePage && pages[i] <= this.endVisiblePage)
				{
					delete this.drawingPages[pages[i]].Image;
				}
			}

			this.scheduleRepaint();
		};
		this.scheduleRepaint = function(formsCallBack, annotsCallback) {
			let oThis = this;
			if (this.scheduledRepaintTimer == null) {
				this.scheduledRepaintTimer = setTimeout(function() {
					oThis.scheduledRepaintTimer = null;
					oThis.isRepaint = false;

					oThis.onRepaintFormsCallbacks.forEach(function(callback) {
						callback();
					});
					oThis.onRepaintAnnotsCallbacks.forEach(function(callback) {
						callback();
					});
					oThis.onRepaintFormsCallbacks = [];
					oThis.onRepaintAnnotsCallbacks = [];

					if (oThis.Api && oThis.Api.printPreview)
						oThis.Api.printPreview.update();

					oThis.isRepaint = true;
				});
			}
			
			if (formsCallBack)
				this.onRepaintFormsCallbacks.push(formsCallBack);
			if (annotsCallback)
				this.onRepaintAnnotsCallbacks.push(annotsCallback);
		};

		this.onRepaintForms = function(pages) {
			if (this.startVisiblePage < 0 || this.endVisiblePage < 0)
				return false;

			for (var i = 0, len = pages.length; i < len; i++)
			{
				if (pages[i] >= this.startVisiblePage && pages[i] <= this.endVisiblePage)
				{
					this.pagesInfo.pages[pages[i]].needRedrawForms = true;
					this.doc.ClearCacheForms(pages[i]);
				}
			}

			this.scheduleRepaint();
		};
		this.onRepaintAnnots = function(pages) {
			if (this.startVisiblePage < 0 || this.endVisiblePage < 0)
				return false;

			for (var i = 0, len = pages.length; i < len; i++)
			{
				if (pages[i] >= this.startVisiblePage && pages[i] <= this.endVisiblePage)
				{
					this.pagesInfo.pages[pages[i]].needRedrawAnnots = true;
					this.pagesInfo.pages[pages[i]].needRedrawMarkups = true;
					this.doc.ClearCacheAnnots(pages[i]);
				}
			}

			this.scheduleRepaint();
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

		this.checkReady = function() {
			let _t = this;
			let oDoc = this.getPDFDoc();

			// в интерфейсе есть проблема - нужно посылать onDocumentContentReady после setAdvancedOptions
			setTimeout(function(){

				if (!_t.isStarted)
				{
					AscCommon.addMouseEvent(_t.canvasForms, "down", _t.onMouseDown);
					AscCommon.addMouseEvent(_t.canvasForms, "move", _t.onMouseMove);
					AscCommon.addMouseEvent(_t.canvasForms, "up", _t.onMouseUp);

					let targetElem = document.getElementById('id_target_cursor');
					if (targetElem)
						targetElem.style.pointerEvents = "none";

					global_mouseEvent.Sender = _t.canvasForms;

					_t.parent.onmousewheel = _t.onMouseWhell;
					if (_t.parent.addEventListener)
						_t.parent.addEventListener("DOMMouseScroll", _t.onMouseWhell, false);

					_t.startTimer();
				}

				if (_t.isStarted && _t.pageDetector && !oDoc.fontLoader.isWorking() && _t.IsOpenFormsInProgress == false) {
					_t.sendEvent("onFileOpened");

					_t.sendEvent("onPagesCount", _t.file.pages.length);
					_t.sendEvent("onCurrentPageChanged", 0);

					_t.sendEvent("onStructure", _t.structure);
					_t.id_main.style.display = "";
				}
				else {
					_t.id_main.style.display = "none";
					_t.checkReady();
				}
				
			}, 0);
		};
		this.onDocumentReady = function()
		{
			this.checkReady();
			this.file.onRepaintPages = this.onUpdatePages.bind(this);
			this.file.onRepaintForms = this.onRepaintForms.bind(this);
			this.file.onRepaintAnnots = this.onRepaintAnnots.bind(this);
			this.file.onUpdateStatistics = this.onUpdateStatistics.bind(this);
			this.currentPage = -1;
			this.structure = this.file.getStructure();

			this.resize(true);
			
			this.openForms();
			this.openAnnots();
			
			if (this.thumbnails)
				this.thumbnails.init(this);

			this.setMouseLockMode(true);

			if (this.drawingPages[0]) {
				this.navigateToPage(0, 0, this.scrollMaxX / 2);
			}
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
					this.Api.asc_setCurrentPassword(password);
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
			{
				this.savedPassword = password || "";
				this.Api.sendEvent("asc_onDocumentPassword", "" !== this.savedPassword);
			}

			this.afterOpen();
		};

		this.afterOpen = function()
		{
			this.pagesInfo.setCount(this.file.pages.length);
			let oDoc = this.getPDFDoc();
			oDoc.GetDrawingDocument().m_lPagesCount = this.file.pages.length;

			for (let i = 0; i < this.file.pages.length; i++) {
				this.DrawingObjects.mergeDrawings(i);
			}

			let standardFonts = this.file.nativeFile["getInteractiveFormsStandardFonts"]();
			let embeddedFonts = this.file.nativeFile["getInteractiveFormsEmbeddedFonts"]();
			AscFonts.initEmbeddedFonts(standardFonts.concat(embeddedFonts));

			var _nativeFile = this.file.nativeFile;
			AscFonts.loadEmbeddedFont = function(id)
			{
				let prefix = AscFonts.getEmbeddedFontPrefix();
				if (id.startsWith(prefix))
					id = id.substr(prefix.length);

				return _nativeFile["getFontByID"](id);
			}

			g_fontApplication.GetFontInfo = g_fontApplication.GetFontInfoWithEmbed;
			g_fontApplication.LoadFont = g_fontApplication.LoadFontWithEmbed;

			this.checkLoadCMap();

			AscCommon.g_oIdCounter.Set_Load(false); // to do возможно не тут стоит выключать флаг

			if (this.file && !this.file.isNeedPassword() && !this.file.isValid())
				this.Api.sendEvent("asc_onError", Asc.c_oAscError.ID.ConvertationOpenError, Asc.c_oAscError.Level.Critical);

			this.Api.WordControl.m_oOverlayApi = this.overlay;
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
			this.onUpdateOverlay();
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
			
			let aFontsToLoad = [];
			for (let i = 0; i < aFormsInfo["Fields"].length; i++)
			{
				oFormInfo	= aFormsInfo["Fields"][i];
				oRect		= oFormInfo["rect"];

				oForm = this.doc.AddFieldByParams(oFormInfo["name"], oFormInfo["type"], oFormInfo["page"], [oRect["x1"], oRect["y1"], oRect["x2"], oRect["y2"]]);
				
				if (!oForm) {
					// console.log("Error while reading form, index " + oFormInfo["AP"]["i"]);
					continue;
				}

				if (oFormInfo["AP"] != null) {
					oForm.SetApIdx(oFormInfo["AP"]["i"]);
					oForm.SetHasOriginView(Boolean(oFormInfo["AP"]["have"]));
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
					oForm.SetHeaderPosition(oFormInfo["position"]);
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
				if (null != oFormInfo["TI"])
				{
					oForm.SetTopIndex(oFormInfo["TI"]);
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
				if (oFormInfo["alignment"] != null && [AscPDF.FIELD_TYPES.combobox, AscPDF.FIELD_TYPES.text, AscPDF.FIELD_TYPES.listbox].includes(oFormInfo["type"]))
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
				if (oFormInfo["noexport"])
				{
					oForm.SetNoExport(Boolean(oFormInfo["noexport"]));
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
				
				if (oFormInfo["curIdxs"])
				{
					oForm.SetCurIdxs(oFormInfo["curIdxs"]);
					if (oFormInfo["value"] != null)
					{
						oForm._value = oFormInfo["value"];
					}
				}
				else if (oFormInfo["value"] != null && oForm.GetType() != AscPDF.FIELD_TYPES.button)
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

				if (oFormInfo["font"]) {
					if (oFormInfo["font"]["color"] != null)
						oForm.SetTextColor(oFormInfo["font"]["color"]);
					if (oFormInfo["font"]["name"])
						oForm.SetTextFont(oFormInfo["font"]["name"]);
					if (oForm.GetType() == AscPDF.FIELD_TYPES.button && oFormInfo["font"]["AP"])
						oForm.SetTextFontActual(oFormInfo["font"]["AP"]);
					else if (oFormInfo["font"]["actual"]) {
						oForm.SetTextFontActual(oFormInfo["font"]["actual"]);
					}
					else if (oFormInfo["font"]["name"]) {
						oForm.SetTextFontActual(AscFonts.getEmbeddedFontPrefix() + oFormInfo["font"]["name"]);
					}
					else {
						oForm.SetTextFontActual(AscPDF.DEFAULT_FIELD_FONT);
					}

					// внутренний ключ для пересылки обратно (зачем? - попросили)
					oForm.SetFontKey(oFormInfo["font"]["key"]);

					if (oFormInfo["font"]["size"] != null)
						oForm.SetTextSize(oFormInfo["font"]["size"]);

					if (oFormInfo["font"]["style"] != null) {
						oForm.SetFontStyle({
							bold: Boolean((oFormInfo["font"]["style"] >> 0) & 1),
							italic: Boolean((oFormInfo["font"]["style"] >> 1) & 1),
						});
					}
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

			this.doc.FillButtonsIconsOnOpen();
		};
		this.openAnnots = function() {
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
					let creationDate	= oAnnotInfo["CreationDate"] ? AscPDF.ParsePDFDate(oAnnotInfo["CreationDate"]) : null;
					let creationStamp	= creationDate ? String(creationDate.getTime()) : undefined;
					let modDate			= oAnnotInfo["LastModified"] ? AscPDF.ParsePDFDate(oAnnotInfo["LastModified"]) : null;
					let modStamp		= modDate ? String(modDate.getTime()) : undefined;

					oAnnot = oDoc.AddAnnotByProps({
						page:			oAnnotInfo["page"],
						name:			oAnnotInfo["UniqueName"], 
						creationDate:	creationStamp,
						modDate:		modStamp,
						contents:		oAnnotInfo["Contents"],
						author:			oAnnotInfo["User"],
						rect:			aRect,
						type:			oAnnotInfo["Type"],
						apIdx:			oAnnotInfo["AP"]["i"],
						uid:			oAnnotInfo["OUserID"]
					});
	
					oAnnot.SetDrawFromStream(Boolean(oAnnotInfo["AP"]["have"]));
					oAnnot.SetOriginPage(oAnnotInfo["page"]);

					if (oAnnotInfo["RefTo"] == null)
						oAnnotsMap[oAnnotInfo["AP"]["i"]] = oAnnot;
					
					if (oAnnotInfo["IT"] != null)
						oAnnot.SetIntent(oAnnotInfo["IT"]);

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

					if (oAnnotInfo["Icon"] != null)
                        oAnnot.SetIconType(oAnnotInfo["Icon"]);
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
						oAnnot.SetRectangleDiff(oAnnotInfo["RD"]);
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
					if (oAnnotInfo["Icon"] != null)
						oAnnot.SetIconType(oAnnotInfo["Icon"]);
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

					if (oAnnotInfo["Rotate"] != null)
						oAnnot.SetRotate(oAnnotInfo["Rotate"]);
					if (oAnnotInfo["InRect"] != null)
						oAnnot.SetInRect(oAnnotInfo["InRect"]);
					
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
					else {
						oAnnot.SetWidth(1);
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
				oDoc.CheckComment(oAnnotsMap[apIdx]);
			}
			
			this.IsOpenAnnotsInProgress = false;
			oDoc.UpdateMaxApIdx(nMaxIdx);
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
			this.zoomCoordinate = this.getPageByCoords2(x + this.x, y + this.y);
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
				yOffset = yOffset >> 0;
				posY += yOffset;
			}

			if (posY > this.scrollMaxY)
				posY = this.scrollMaxY;

			var posX = 0;

			if (xOffset)
			{
				xOffset = xOffset >> 0;
				posX += xOffset;
			}

			if (posX > this.scrollMaxX)
				posX = this.scrollMaxX;

			let oDoc		= this.getPDFDoc();
			let oActiveObj	= oDoc.GetActiveObject();
			let nPage		= oActiveObj ? oActiveObj.GetPage() : undefined;

			this.checkVisiblePages();
			// выход из активного объекта если сместились на другую страницу
			if (oActiveObj && !(nPage >= this.startVisiblePage && nPage <= this.endVisiblePage)) {
				oDoc.BlurActiveObject();
			}

			this.m_oScrollVerApi.scrollToY(posY);
			this.m_oScrollHorApi.scrollToX(posX);
		};
		this.scrollToXY = function(posY, posX) {
			let oDoc		= this.getPDFDoc();
			let oActiveObj	= oDoc.GetActiveObject();
			let nPage		= oActiveObj ? oActiveObj.GetPage() : undefined;

			this.m_oScrollVerApi.scrollToY(posY);
			this.m_oScrollVerApi.scrollToX(posX);

			this.checkVisiblePages();
			// выход из активного объекта если сместились на другую страницу
			if (oActiveObj && !(nPage >= this.startVisiblePage && nPage <= this.endVisiblePage)) {
				oDoc.BlurActiveObject();
			}
		};

		this.navigateToLink = function(link)
		{
			if ("" === link["link"])
				return;

			if ("#" === link["link"].charAt(0))
			{
				let nPage	= parseInt(link["link"].substring(1));
				let oTr		= this.getPDFDoc().pagesTransform[nPage].invert;
				let oPos	= oTr.TransformPoint(0, link["dest"]);

				this.scrollToXY(this.scrollY + oPos.y, this.scrollX + oPos.x);
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

			function isLandscape(angle) {
				// Углы поворота, указывающие на ландшафтную ориентацию
				const landscapeAngles = [90, 270];
				return landscapeAngles.includes(angle);
			}

			this.documentWidth = 0;
			for (let i = 0, len = this.drawingPages.length; i < len; i++)
			{
				let rotateAngle = this.getPageRotate(i);
				let pageW = isLandscape(rotateAngle) ? this.drawingPages[i].H : this.drawingPages[i].W;

				if (pageW > this.documentWidth)
					this.documentWidth = pageW;
			}
			// прибавим немного
			this.documentWidth += (4 * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

			var curTop = this.betweenPages + this.offsetTop;
			for (let i = 0, len = this.drawingPages.length; i < len; i++)
			{
				let rotateAngle = this.getPageRotate(i);
				let pageW = isLandscape(rotateAngle) ? this.drawingPages[i].H : this.drawingPages[i].W;
				let pageH = isLandscape(rotateAngle) ? this.drawingPages[i].W : this.drawingPages[i].H;

				this.drawingPages[i].X = (this.documentWidth - pageW) >> 1;
				this.drawingPages[i].Y = curTop;

				curTop += pageH;
				curTop += this.betweenPages;
			}

			this.documentHeight = curTop;
		};

		this.setCursorType = function(cursor)
		{
			let oDoc = this.getPDFDoc();
			let oDrDoc = oDoc.GetDrawingDocument();

			if (oDrDoc.m_sLockedCursorType) {
				return;
			}

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
			var pageObject = this.getPageByCoords(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
			if (!pageObject)
				return null;

			// после конвертации не даем кликать линки на странице
			if (this.file.pages[pageObject.index].isRecognized) {
				return null;
			}

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
		this.getPageFieldByCoords = function(x, y, pageIndex, bGetHidden)
		{
			var pageFields = this.pagesInfo.pages[pageIndex];
			if (pageFields.fields)
			{
				for (var i = 0, len = pageFields.fields.length; i < len; i++)
				{
					if (x >= pageFields.fields[i]._origRect[0] && x <= pageFields.fields[i]._origRect[2] &&
						y >= pageFields.fields[i]._origRect[1] && y <= pageFields.fields[i]._origRect[3]) {
						if (bGetHidden) {
							return pageFields.fields[i];
						}
						else if (pageFields.fields[i].IsHidden() == false) {
							return pageFields.fields[i];
						}
					}
				}
			}
			return null;
		};
		this.getPageFieldByMouse = function(bGetHidden)
		{
			var pageObject = this.getPageByCoords(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
			if (!pageObject)
				return null;

			return this.getPageFieldByCoords(pageObject.x, pageObject.y, pageObject.index, bGetHidden);
		};
		this.getPageAnnotByMouse = function(bGetHidden)
		{
			let oDoc = this.getPDFDoc();
			let pageObject = this.getPageByCoords(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
			if (!pageObject)
				return null;

			const nPage	= pageObject.index;

			// координаты клика на странице в MM
			let pageObjectMM = this.getPageByCoords2(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);

			let page = this.pagesInfo.pages[nPage];
			
			// если есть заселекченная shape base аннотация под мышкой, залезающая на другую страницу
			if (oDoc.mouseDownAnnot) {
				let oController = oDoc.GetController();
				let oActiveAnnot = oDoc.mouseDownAnnot;

				if (oActiveAnnot.IsShapeBased()) {
					let nAnnotPage = oActiveAnnot.GetPage();

					let oPt = pageObjectMM.index != nAnnotPage ? AscPDF.ConvertCoordsToAnotherPage(pageObjectMM.x, pageObjectMM.y, pageObjectMM.index, nAnnotPage) : pageObjectMM;
					let oGeomEditSelection = oController.selection.geometrySelection ? oController.selection.geometrySelection : null;

					if (oActiveAnnot == oDoc.GetShapeBasedAnnotById(this.DrawingObjects.getGraphicInfoUnderCursor(nAnnotPage, oPt.x, oPt.y).objectId)) {
						if (oActiveAnnot.hitToHandles(oPt.x, oPt.y) != -1 || (oGeomEditSelection && oGeomEditSelection.hitToGeometryEdit(oPt.x, oPt.y))) {
							return oActiveAnnot;
						}
					}
				}
			}

			if (page.annots)
			{
				// сначала ищем text annot (sticky note)
				for (let i = page.annots.length -1; i >= 0; i--)
				{
					let oAnnot = page.annots[i];
					let nAnnotWidth		= 20 / (this.zoom);
					let nAnnotHeight	= 20 / (this.zoom);
					
					if (true !== bGetHidden && oAnnot.IsHidden() == true || false == oAnnot.IsComment())
						continue;
					
					if (pageObject.x >= oAnnot._origRect[0] && pageObject.x <= oAnnot._origRect[0] + nAnnotWidth &&
					pageObject.y >= oAnnot._origRect[1] && pageObject.y <= oAnnot._origRect[1] + nAnnotHeight) {
						if (bGetHidden) {
							return oAnnot;
						}
						else if (oAnnot.IsHidden() == false) {
							return oAnnot;
						}
					}
				}

				for (let i = page.annots.length -1; i >= 0; i--)
				{
					let oAnnot = page.annots[i];
					let nAnnotWidth		= (oAnnot._origRect[2] - oAnnot._origRect[0]);
					let nAnnotHeight	= (oAnnot._origRect[3] - oAnnot._origRect[1]);
					
					if (true !== bGetHidden && oAnnot.IsHidden() == true || oAnnot.IsComment())
						continue;

					// у draw аннотаций ищем по path
					if (oAnnot.IsShapeBased())
					{
						if (oAnnot.hitInBoundingRect(pageObjectMM.x, pageObjectMM.y) || oAnnot.hitToHandles(pageObjectMM.x, pageObjectMM.y) != -1 || oAnnot.hitInPath(pageObjectMM.x, pageObjectMM.y) || oAnnot.hitInInnerArea(pageObjectMM.x, pageObjectMM.y))
							return oAnnot;
					}

					if (pageObject.x >= oAnnot._origRect[0] && pageObject.x <= oAnnot._origRect[0] + nAnnotWidth &&
						pageObject.y >= oAnnot._origRect[1] && pageObject.y <= oAnnot._origRect[1] + nAnnotHeight)
					{
						// у маркап аннотаций ищем по quads (т.к. rect too wide)
						if (oAnnot.IsTextMarkup())
						{
							if (oAnnot.IsInQuads(pageObject.x, pageObject.y))
								return oAnnot;
						}
					}
				}
			}
			return null;
		};
		this.canInteract = function() {
			// не даем взаимодействовать с документом пока не произошла отрисовка
			return this.scheduledRepaintTimer == null && this.isRepaint != true && this.initPaintDone == true;
		};
		this.getPageDrawingByMouse = function()
		{
			var pageObject = this.getPageByCoords2(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
			if (!pageObject)
				return null;

			let oCurState = this.DrawingObjects.curState;
			this.DrawingObjects.curState = this.DrawingObjects.nullState;

			let oDoc = this.getPDFDoc();
			let oDrawingUnderMouse;

			// handle selected drawing in another page
			let oActiveDrawing = oDoc.activeDrawing;
			if (oActiveDrawing) {
				let nDrawingPage = oActiveDrawing.GetPage();

				if (pageObject.index != nDrawingPage) {
					let oPt = AscPDF.ConvertCoordsToAnotherPage(pageObject.x, pageObject.y, pageObject.index, nDrawingPage);

					if (oActiveDrawing == oDoc.GetDrawingById(this.DrawingObjects.getGraphicInfoUnderCursor(nDrawingPage, oPt.x, oPt.y).objectId)) {
						if (oActiveDrawing.hitToHandles(oPt.x, oPt.y) != -1) {
							oDrawingUnderMouse = oActiveDrawing;
						}
					}
				}
			}
			
			if (null == oDrawingUnderMouse) {
				oDrawingUnderMouse = oDoc.GetDrawingById(this.DrawingObjects.getGraphicInfoUnderCursor(pageObject.index, pageObject.x, pageObject.y).objectId);
				if (oDrawingUnderMouse && oDrawingUnderMouse.GetPage() != pageObject.index) {
					oDrawingUnderMouse = null;
				}
			}
			
			this.DrawingObjects.curState = oCurState;

			return oDrawingUnderMouse;
		};

		this.onMouseDown = function(e)
		{
			Asc.editor.checkLastWork();

			if (oThis.touchManager && oThis.touchManager.checkTouchEvent(e))
			{
				oThis.touchManager.startTouchingInProcess();
				let res = oThis.touchManager.mainOnTouchStart(e);
				oThis.touchManager.stopTouchingInProcess();
				return res;
			}

			if (oThis.touchManager)
				oThis.touchManager.checkMouseFocus(e);

			oThis.isFocusOnThumbnails = false;
			AscCommon.stopEvent(e);

			let oDoc = oThis.getPDFDoc();
			oDoc.HideComments();

			var mouseButton = AscCommon.getMouseButton(e || {});
			AscCommon.check_MouseDownEvent(e, true);

			if (mouseButton !== 0)
			{
				if (2 === mouseButton)
				{
					var x = AscCommon.global_mouseEvent.X - oThis.x;
					var y = AscCommon.global_mouseEvent.Y - oThis.y;

					var isInSelection = false;
					
					if (oThis.overlay.m_oContext)
					{
						let pageCoords = oThis.getPageByCoords(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
						let isSelectionUse = oThis.file.isSelectionUse();
						let selection = oThis.file.getSelection();
						let pageSelQuads = pageCoords ? selection.quads.find(function(pageQuads) {
							return pageQuads.page == pageCoords.index;
						}) : null;

						if (oThis.canSelectPageText() && (pageCoords && isSelectionUse && pageSelQuads) && AscPDF.IsInQuads(pageSelQuads.quads, pageCoords.x, pageCoords.y))
						{
							isInSelection = true;
						}
					}

					if (isInSelection)
					{
						oThis.Api.sync_BeginCatchSelectedElements();
						oThis.Api.sync_ChangeLastSelectedElement(Asc.c_oAscTypeSelectElement.Text, undefined);
						oThis.Api.sync_EndCatchSelectedElements();

						oThis.Api.sync_ContextMenuCallback({
							Type: Asc.c_oAscPdfContextMenuTypes.Common,
							X_abs: x,
							Y_abs: y
						});
					}
					else
					{
						oThis.removeSelection();
						oDoc.OnMouseDown(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y, AscCommon.global_mouseEvent);
						oThis.Api.sync_ContextMenuCallback({
							Type: Asc.c_oAscPdfContextMenuTypes.Common,
							X_abs: x,
							Y_abs: y
						});
					}
				}
				return;
			}

			oThis.isMouseDown = true;

			if (!oThis.file || !oThis.file.isValid())
				return;

			global_mouseEvent.LockMouse();

			oThis.mouseDownCoords.X = AscCommon.global_mouseEvent.X;
			oThis.mouseDownCoords.Y = AscCommon.global_mouseEvent.Y;

			oThis.isMouseMoveBetweenDownUp = false;
			oDoc.OnMouseDown(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y, AscCommon.global_mouseEvent);
		};

		this.onMouseDownEpsilon = function(e)
		{
			if (oThis.MouseHandObject)
			{
				if (oThis.getPDFDoc().mouseDownLinkObject)
				{
					// если нажали на ссылке - то не зажимаем лапу
					oThis.setCursorType("pointer");
					return;
				}
				// режим лапы. просто начинаем режим Active - зажимаем лапу
				oThis.MouseHandObject.X = oThis.mouseDownCoords.X;
				oThis.MouseHandObject.Y = oThis.mouseDownCoords.Y;
				oThis.MouseHandObject.Active = true;
				oThis.MouseHandObject.ScrollX = oThis.scrollX;
				oThis.MouseHandObject.ScrollY = oThis.scrollY;
				return;
			}

			var pageObjectLogic = this.getPageByCoords2(oThis.mouseDownCoords.X, oThis.mouseDownCoords.Y);
			if (e.ShiftKey) {
				this.file.Selection.IsSelection = true;
				this.file.onMouseMove(pageObjectLogic.index, pageObjectLogic.x, pageObjectLogic.y);
			}
			else {
				this.file.onMouseDown(pageObjectLogic.index, pageObjectLogic.x, pageObjectLogic.y);
			}

			if (-1 === this.timerScrollSelect && AscCommon.global_mouseEvent.IsLocked)
			{
				this.timerScrollSelect = setInterval(this.selectWheel, 20);
			}
		};

		this.onMouseUp = function(e)
		{
			Asc.editor.checkLastWork();

			if (oThis.touchManager && oThis.touchManager.checkTouchEvent(e))
			{
				oThis.touchManager.startTouchingInProcess();
				let res = oThis.touchManager.mainOnTouchEnd(e);
				oThis.touchManager.stopTouchingInProcess();
				return res;
			}

			oThis.isFocusOnThumbnails = false;
			//if (e && e.preventDefault)
			//	e.preventDefault();

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
			oDoc.OnMouseUp(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y, AscCommon.global_mouseEvent);

			if (oThis.canSelectPageText() && !oThis.MouseHandObject && !oDoc.mouseDownAnnot && !oDoc.mouseDownField)
			{
				var pageObjectLogic = oThis.getPageByCoords2(oThis.mouseDownCoords.X, oThis.mouseDownCoords.Y);
				if (global_mouseEvent.ClickCount == 2)
					oThis.file.selectWholeWord(pageObjectLogic.index, pageObjectLogic.x, pageObjectLogic.y);
				else if (global_mouseEvent.ClickCount == 3)
					oThis.file.selectWholeRow(pageObjectLogic.index, pageObjectLogic.x, pageObjectLogic.y);
				else if (global_mouseEvent.ClickCount == 4)
					oThis.file.selectWholePage(pageObjectLogic.index);
			}

			// если было нажатие - то отжимаем
			if (oThis.isMouseMoveBetweenDownUp)
				oThis.file.onMouseUp();

			if (oThis.MouseHandObject)
				oThis.MouseHandObject.Active = false;
			oThis.isMouseMoveBetweenDownUp = false;

			if (-1 !== oThis.timerScrollSelect)
			{
				clearInterval(oThis.timerScrollSelect);
				oThis.timerScrollSelect = -1;
			}
		};

		this.onMouseMove = function(e)
		{
			Asc.editor.checkLastWork();

			if (oThis.touchManager && oThis.touchManager.checkTouchEvent(e))
			{
				oThis.touchManager.startTouchingInProcess();
				let res = oThis.touchManager.mainOnTouchMove(e);
				oThis.touchManager.stopTouchingInProcess();
				return res;
			}

			if (!oThis.file || !oThis.file.isValid())
				return;

			let oDoc = oThis.getPDFDoc();
			AscCommon.check_MouseMoveEvent(e);

			if (e && e.preventDefault)
				e.preventDefault();

			if (oThis.MouseHandObject)
			{
				if (oThis.canMovePageByHand())
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
						oDoc.OnMouseMove(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y, AscCommon.global_mouseEvent);
				}
				return;
			}
			else
			{
				if (oThis.getPDFDoc().mouseDownLinkObject)
				{
					// селект начат на ссылке. смотрим, нужно ли начать реально селект
					if (oThis.isMouseMoveBetweenDownUp)
					{
						// вышли за eps
						oThis.getPDFDoc().mouseDownLinkObject = null;
						oThis.setCursorType("default");
					}
					else
					{
						oThis.setCursorType("pointer");
					}
				}

				if (oThis.isMouseDown)
				{
					if (oThis.canSelectPageText())
					{
						// нажатая мышка - курсор всегда default (так как за eps вышли)
						oThis.setCursorType("default");

						var pageObjectLogic = oThis.getPageByCoords2(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
						oThis.file.onMouseMove(pageObjectLogic.index, pageObjectLogic.x, pageObjectLogic.y);
					}
					else
					{
						if (false == editor.isEmbedVersion)
							oDoc.OnMouseMove(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y, AscCommon.global_mouseEvent);
					}
				}
				else
				{
					oThis.getPDFDoc().OnMouseMove(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y, AscCommon.global_mouseEvent);
				}
			}
			return false;
		};
		this.canMovePageByHand = function() {
			return this.isMouseDown && this.MouseHandObject.Active && !this.doc.activeDrawing && !this.doc.mouseDownAnnot && !this.Api.isInkDrawerOn() && !this.Api.isStartAddShape;
		};
		this.canSelectPageText = function() {
			let oDoc = this.getPDFDoc();
			return !oDoc.activeDrawing && !oDoc.activeForm && (!oDoc.mouseDownAnnot || (oDoc.mouseDownAnnot && oDoc.mouseDownAnnot.IsTextMarkup() == true)) && !this.Api.isInkDrawerOn() && !this.Api.isStartAddShape;
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

			let values = AscCommon.checkMouseWhell(e, {
				isSupportBidirectional : false,
				isAllowHorizontal : oThis.isVisibleHorScroll,
				isUseMaximumDelta : true
			});
			
			if (0 !== values.x)
				oThis.m_oScrollHorApi.scrollBy(values.x, 0, false);
			if (0 !== values.y)
				oThis.m_oScrollVerApi.scrollBy(0, values.y, false);

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

		this.paint = function(formsCallBack, annotsCallback)
		{
			this.scheduleRepaint(formsCallBack, annotsCallback);
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

					this.overlay.CheckPoint(_x,_y);
					this.overlay.CheckPoint(_x + _w, _y + _h);

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
			let width = AscCommon.AscBrowser.convertToRetinaValue(this.drawingPages[pageIndex].W, true) >> 0;
        	let height = AscCommon.AscBrowser.convertToRetinaValue(this.drawingPages[pageIndex].H, true) >> 0;

			var dKoefX = scale * width / this.file.pages[pageIndex].W;
			var dKoefY = scale * height / this.file.pages[pageIndex].H;

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
			let width = AscCommon.AscBrowser.convertToRetinaValue(this.drawingPages[pageIndex].W, true) >> 0;
        	let height = AscCommon.AscBrowser.convertToRetinaValue(this.drawingPages[pageIndex].H, true) >> 0;

			var dKoefX = scale * width / this.file.pages[pageIndex].W;
			var dKoefY = scale * height / this.file.pages[pageIndex].H;

			for (var i = 0; i < searchingObj.length; i++)
			{
				this.drawSearchPlaces(dKoefX, dKoefY, pageCoords.x, pageCoords.y, searchingObj[i]);
			}
		};

		this.onUpdateOverlay = function()
		{
			Asc.editor.checkLastWork();
			
			if (!this.overlay || this.scheduledRepaintTimer != null)
				return;

			let oDoc = this.getPDFDoc();
			let oDrDoc = oDoc.GetDrawingDocument();
			if (oDoc.fontLoader.isWorking() || this.IsOpenFormsInProgress)
				return;

			this.overlay.Clear();

			oDrDoc.AutoShapesTrack.PageIndex = -1;

			if (!this.file)
				return;

			if (this.startVisiblePage < 0 || this.endVisiblePage < 0)
				return;

			// seletion
			var ctx = this.overlay.m_oContext;
			ctx.globalAlpha = 0.2;

			if (this.IsSearch)
			{
				if (oDoc.SearchEngine.Show)
				{
					ctx.globalAlpha = 0.5;
					ctx.fillStyle = "rgba(255,200,0,1)";
					ctx.beginPath();

					for (let i = this.startVisiblePage; i <= this.endVisiblePage; i++)
					{
						var pageMatches = oDoc.SearchEngine.GetPdfPageMatches(i);
						if (0 != pageMatches.length) {
							oDrDoc.AutoShapesTrack.SetCurrentPage(i, true);
							this.drawSearch(i, pageMatches);
						}
							
					}

					ctx.fill();
					ctx.globalAlpha = 0.2;
				}
				ctx.beginPath();

				if (this.CurrentSearchNavi && oDoc.SearchEngine.Show)
				{
					if (!(this.CurrentSearchNavi instanceof AscWord.Paragraph)) {
						var pageNum  = this.CurrentSearchNavi[0].PageNum;
						ctx.fillStyle = "rgba(51,102,204,255)";
						if (pageNum >= this.startVisiblePage && pageNum <= this.endVisiblePage)
						{
							oDrDoc.AutoShapesTrack.SetCurrentPage(pageNum, true);
							ctx.globalAlpha = 0.2;
							this.drawSearchCur(pageNum, this.CurrentSearchNavi);
						}
					}
				}
			}

			oDrDoc.private_StartDrawSelection(this.overlay);
			
			//if (!this.MouseHandObject)
			{
				ctx.fillStyle = "rgba(51,102,204,255)";
				ctx.beginPath();

				if (this.file.isSelectionUse())
				{
					for (let i = this.startVisiblePage; i <= this.endVisiblePage; i++)
					{
						var pageCoords = this.pageDetector.pages[i - this.startVisiblePage];
						ctx.beginPath();
						oDrDoc.AutoShapesTrack.SetCurrentPage(i, true);
						ctx.globalAlpha = 0.2;
						this.file.drawSelection(i, this.overlay, pageCoords.x, pageCoords.y);
						ctx.fill();
					}
				}

				this.DrawingObjects.updateSelectionState();

				if (this.DrawingObjects.needUpdateOverlay())
				{
					oDrDoc.AutoShapesTrack.PageIndex = -1;
					this.DrawingObjects.drawOnOverlay(oDrDoc.AutoShapesTrack);
					oDrDoc.AutoShapesTrack.CorrectOverlayBounds();
				}
				else if (oDoc.mouseDownAnnot)
				{
					if (oDoc.mouseDownAnnot.IsFreeText() && oDoc.mouseDownAnnot.IsInTextBox() && oDoc.mouseDownAnnot.GetDocContent().IsSelectionUse()) {
						ctx.beginPath();
						oDrDoc.SetTextSelectionOutline(true);
						oDoc.mouseDownAnnot.GetDocContent().DrawSelectionOnPage(0);
						oDrDoc.private_EndDrawSelection();
					}
					else {
						let nPage = oDoc.mouseDownAnnot.GetPage();
						oDrDoc.AutoShapesTrack.SetCurrentPage(nPage, true);
						this.DrawingObjects.drawSelect(nPage);
					}
				}
				else if (oDoc.activeDrawing) {
					let nPage = oDoc.activeDrawing.GetPage();
					oDrDoc.SetTextSelectionOutline(true);
					oDrDoc.private_EndDrawSelection();
					oDrDoc.AutoShapesTrack.PageIndex = nPage;
					this.DrawingObjects.drawSelect(nPage);
				}
			}
			if (oDoc.activeForm && oDoc.activeForm.content && oDoc.activeForm.content.IsSelectionUse() && oDoc.activeForm.content.IsSelectionEmpty() == false)
			{
				ctx.beginPath();
				oDoc.activeForm.content.DrawSelectionOnPage(0);
				oDrDoc.private_EndDrawSelection();
			}

			if (oDrDoc.MathTrack.IsActive())
			{
				var dGlobalAplpha = ctx.globalAlpha;
				ctx.globalAlpha = 1.0;
				oDrDoc.DrawMathTrack(this.overlay);
				ctx.globalAlpha = dGlobalAplpha;
			}
			
			ctx.globalAlpha = 1.0;

			for (let i = this.startVisiblePage; i <= this.endVisiblePage; i++)
			{
				oDrDoc.AutoShapesTrack.SetCurrentPage(i, true);
				ctx.globalAlpha = 1.0;
				oDoc.Draw_ForeingSelection(i);
				oDoc.CollaborativeEditing.Update_ForeignSelectedObjectsLabelsPositions(i);
			}
		};

		this.checkVisiblePages = function()
		{
			let oDoc = this.getPDFDoc();

			let yPos = this.scrollY >> 0;
			let yMax = yPos + this.height;
			let xCenter = this.width >> 1;
			if (this.documentWidth > this.width)
				xCenter = (this.documentWidth >> 1) - (this.scrollX) >> 0;

			let lStartPage = -1;
			let lEndPage = -1; 

			let lPagesCount = this.drawingPages.length;
			for (let i = 0; i < lPagesCount; i++)
			{
				let isLandscapePage = this.isLandscapePage(i);

				let page = this.drawingPages[i];
				let pageT = page.Y;
				let pageB = page.Y + (isLandscapePage ? page.W : page.H);
				
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
						delete page.ImageTmp;
						delete page.ImageForms;
						delete page.ImageAnnots;
						delete page.ImageDrawings;
						oDoc.ClearCache(i);
					}
				}
			}

			let oDrDoc = oDoc.GetDrawingDocument();
			oDrDoc.m_lDrawingFirst = lStartPage;
			oDrDoc.m_lDrawingEnd = lEndPage;
			this.startVisiblePage = lStartPage;
			this.endVisiblePage = lEndPage;

			this.updatePageDetector();
		};
		this._paint = function()
		{
			let oDoc = this.getPDFDoc();
			Asc.editor.checkLastWork();
			
			if (oDoc.fontLoader.isFontLoadInProgress() || this.IsOpenFormsInProgress || AscCommon.CollaborativeEditing.waitingImagesForLoad) {
				this.paint();
				return;
			}
			
			if (!this.file || !this.file.isValid() || !this.canvas) {
				this.paint();
				return;
			}

			oDoc.UpdatePagesTransform();
			
			let ctx = this.canvas.getContext("2d");
			let lineW = AscCommon.AscBrowser.retinaPixelRatio >> 0;

			let yPos = this.scrollY >> 0;
			let xCenter = this.width >> 1;
			if (this.documentWidth > this.width)
			{
				xCenter = (this.documentWidth >> 1) - (this.scrollX) >> 0;
			}

			this.checkVisiblePages();
			
			if (this._checkFontsOnPages(this.startVisiblePage, this.endVisiblePage) == false) {
				this.paint();
				return;
			}

			this.canvas.width = this.canvas.width;

			ctx.strokeStyle = AscCommon.GlobalSkin.PageOutline;
			ctx.lineWidth = lineW;

			let isStretchPaint = this.isStretchPaint();
			if (this.isClearPages)
				isStretchPaint = false;

			for (let i = this.startVisiblePage; i <= this.endVisiblePage; i++)
			{
				// отрисовываем страницу
				let page = this.drawingPages[i];
				if (!page)
					break;

				let w = AscCommon.AscBrowser.convertToRetinaValue(page.W, true);
				let h = AscCommon.AscBrowser.convertToRetinaValue(page.H, true);
				let x = ((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - (w >> 1);
				let y = ((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

				let needNewPage = this.isClearPages || !page.Image || (page.Image && ((page.Image.requestWidth !== w) || (page.Image.requestHeight !== h)));

				if (!isStretchPaint)
				{
					let isClearAttack = this.isClearPages;
					if (page.Image && page.createdInStretchMode === true)
					{
						isClearAttack = true;
						delete page.createdInStretchMode;
					}

					if (!this.file.cacheManager)
					{
						if (isClearAttack || (page.Image && ((page.Image.requestWidth !== w) || (page.Image.requestHeight !== h))))
							delete page.Image;
					}
					else
					{
						if (isClearAttack || (page.Image && ((page.Image.requestWidth < w) || (page.Image.requestHeight < h))))
						{
							if (this.file.cacheManager)
								this.file.cacheManager.unlock(page.Image);

							delete page.Image;
						}
					}
				}

				let pageColor = this.Api.getPageBackgroundColor();
				let oImageToDraw = null;
				let bRedrawAnnotsOnMainLayer = false;
				if (!this.file.pages[i].isRecognized) {
					if (!page.Image && !isStretchPaint)
					{
						page.Image = this.file.getPage(i, w, h, undefined, (pageColor.R << 16) | (pageColor.G << 8) | pageColor.B);
						if (this.bCachedMarkupAnnnots) {
							bRedrawAnnotsOnMainLayer = true;
						}

						// нельзя кэшировать с вотермарком - так как есть поворот
						//if (this.Api.watermarkDraw)
						//	this.Api.watermarkDraw.Draw(page.Image.getContext("2d"), w, h);
					}
				}
				
				if (null == page.Image) {
					page.Image = document.createElement('canvas');
					page.Image.width = w;
					page.Image.height = h;
					page.Image.requestWidth = w;
					page.Image.requestHeight = h;
					let tmpPageCtx = page.Image.getContext('2d');
					tmpPageCtx.fillStyle = "rgba(" + pageColor.R + "," + pageColor.G + "," + pageColor.B + ",1)";
					tmpPageCtx.fillRect(0, 0, w, h);

					if (isStretchPaint)
						page.createdInStretchMode = true;
				}

				if (bRedrawAnnotsOnMainLayer) {
					let ctx = page.Image.getContext("2d");
					this._drawDrawingsOnCtx(i, ctx);
					this._drawMarkupAnnotsOnCtx(i, ctx);
					oImageToDraw = page.Image;
				}
				
				if (!this.bCachedMarkupAnnnots) {
					let pageInfo = this.pagesInfo.pages[i];
					let aMarkups = pageInfo.annots.filter(function(annot) {
						return annot.IsTextMarkup();
					});

					let hasMarkups = aMarkups.length != 0;
					let hasDrawings = pageInfo.drawings.length != 0;

					if (false == hasDrawings && false == hasMarkups) {
						oImageToDraw = page.Image;
					}
					else {
						let tmpPageImage = page.TmpImage ? page.TmpImage : document.createElement('canvas');
						let tmpPageCtx = tmpPageImage.getContext('2d');
						
						if (!page.TmpImage) {
							page.TmpImage = tmpPageImage;
						}

						if (pageInfo.needRedrawDrawings || pageInfo.needRedrawMarkups || (needNewPage && !isStretchPaint)) {
							if (page.Image) {
								tmpPageImage.width = page.Image.width;
								tmpPageImage.height = page.Image.height;
							}
							else {
								tmpPageImage.width = w;
								tmpPageImage.height = h;
							}
		
							tmpPageCtx.drawImage(page.Image, 0, 0);
		
							this._drawDrawingsOnCtx(i, tmpPageCtx);
							this._drawMarkupAnnotsOnCtx(i, tmpPageCtx);

							pageInfo.needRedrawDrawings = false;
							pageInfo.needRedrawMarkups = false;
						}
						
						oImageToDraw = tmpPageImage;
					}
				}

				this.blitPageToCtx(ctx, oImageToDraw, i);
				this.pagesInfo.setPainted(i);
				
				if (this.Api.watermarkDraw)
					this.Api.watermarkDraw.Draw(ctx, x, y, w, h);
			}
			
			this.isClearPages = false;
			this.updateCurrentPage(this.pageDetector.getCurrentPage(this.currentPage));
			let oActiveObj = oDoc.GetActiveObject();

			// выход из активного объекта если сместились на другую страницу
			if (oActiveObj && this.pageDetector.pages.map(function(item) {
				return item.num;
			}).includes(oActiveObj.GetPage()) == false) {
				oDoc.BlurActiveObject();
			}

			this._paintAnnots();
			this._paintForms();
			this._paintFormsHighlight();
			this._paintFormsMarkers();
			oDoc.UpdateInterface(true);
			oDoc.UpdateInterfaceTracks();
			
			// Обязательно делаем в конце, т.к. во время отрисовки происходит пересчет
			this._checkTargetUpdate();

			this.initPaintDone = true;
		};
		this.updatePageDetector = function() {
			this.pageDetector = new CCurrentPageDetector(this.canvas.width, this.canvas.height);

			let yPos = this.scrollY >> 0;
			let xCenter = this.width >> 1;
			if (this.documentWidth > this.width)
				xCenter = (this.documentWidth >> 1) - (this.scrollX) >> 0;
			
			for (let i = this.startVisiblePage; i <= this.endVisiblePage; i++)
			{
				let page = this.drawingPages[i];
				if (!page)
					break;

				let w = AscCommon.AscBrowser.convertToRetinaValue(page.W, true);
				let h = AscCommon.AscBrowser.convertToRetinaValue(page.H, true);
				let x = ((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - (w >> 1);
				let y = ((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

				if (this.isLandscapePage(i)) {
					let x = ((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - (h >> 1);
					this.pageDetector.addPage(i, x, y, h, w);
				}
				else {
					this.pageDetector.addPage(i, x, y, w, h);
				}
			}
		};
		this.isStretchPaint = function() {
			return this.Api.WordControl.NoneRepaintPages;
		};
		this._checkTargetUpdate = function() {
			let pdfDocument = this.getPDFDoc();
			let docApi = pdfDocument.GetApi();
			let drawingDocument = pdfDocument.GetDrawingDocument();
			if (docApi.isLockTargetUpdate)
				return;
			
			drawingDocument.UpdateTargetFromPaint = true;
			pdfDocument.CheckTargetUpdate(drawingDocument.UpdateTargetCheck);
			drawingDocument.UpdateTargetCheck = false;
			drawingDocument.UpdateTargetFromPaint = false;
			pdfDocument.CollaborativeEditing.Update_ForeignCursorsPositions();
			drawingDocument.Collaborative_TargetsUpdate(true);
			drawingDocument.CheckTargetShow();
			drawingDocument.CheckTrackTable();
		};
		this.blitPageToCtx = function(ctx, imagePage, pageIndex) {
			if (!imagePage) {
				return;
			}

			let angle = this.getPageRotate(pageIndex);
			let angleRad = angle * Math.PI / 180;

			let imgHeight = imagePage.height;
			let imgWidth = imagePage.width;

			let page = this.drawingPages[pageIndex];

			let xCenter = this.width >> 1;
			let yPos = this.scrollY >> 0;
			if (this.documentWidth > this.width) {
				xCenter = (this.documentWidth >> 1) - (this.scrollX) >> 0;
			}

			let cx = ((xCenter * AscCommon.AscBrowser.retinaPixelRatio) >> 0);
			let yInd = ((page.Y - yPos) * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
			
			let w = AscCommon.AscBrowser.convertToRetinaValue(page.W, true);
			let h = AscCommon.AscBrowser.convertToRetinaValue(page.H, true);

			ctx.save();
			switch (angle) {
				case 90:
					ctx.translate(cx, yInd);
					ctx.rotate(angleRad);
					ctx.drawImage(imagePage, 0, 0, imgWidth, imgHeight, 0, -h >> 1, w, h);
					break;
				case 180:
					ctx.translate(cx, yInd + (h >> 1));
					ctx.rotate(angleRad);
					ctx.drawImage(imagePage, 0, 0, imgWidth, imgHeight, -(w >> 1), -(h >> 1), w, h);
					break;
				case 270:
					ctx.translate(cx, yInd + w >> 0);
					ctx.rotate(angleRad);
					ctx.drawImage(imagePage, 0, 0, imgWidth, imgHeight, 0, -h >> 1, w, h);
					break;
				default: // 0 градусов, по умолчанию
					ctx.drawImage(imagePage,  0, 0, imgWidth, imgHeight, cx - (w >> 1), yInd, w, h);
					break;
			}

			ctx.restore();
		};
		this.isLandscapePage = function(nPage) {
			const angle = this.getPageRotate(nPage);
			// Углы поворота, указывающие на ландшафтную ориентацию
			const landscapeAngles = [90, -90, 270, -270];
			return landscapeAngles.includes(angle);
		};
		this.Get_PageLimits = function(nPage) {
			let oPage = this.file.pages[nPage] || this.file.pages[0];

			return {
				X: 0,
				Y: 0,
				XLimit: oPage.W * g_dKoef_pt_to_mm,
				YLimit: oPage.H * g_dKoef_pt_to_mm
			}
		};
		this.SelectNextForm = function()
		{
			this.doc.SelectNextForm();
		};
		this.SelectPrevForm = function()
		{
			this.doc.SelectPrevForm();
		};
		
		this.checkPagesLinks = function()
		{
			if (this.startVisiblePage < 0 || this.endVisiblePage < 0)
				return false;

			for (var i = this.startVisiblePage; i <= this.endVisiblePage; i++)
			{
				var page = this.pagesInfo.pages[i];
				if (page.isPainted && null === page.links && this.file.pages[i].originIndex != undefined)
				{
					page.links = this.file.getLinks(this.file.pages[i].originIndex);
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
				if (null === this.file.pages[i].text)
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
					if (null !== this.file.pages[this.pagesInfo.countTextPages].text)
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
			if (this.startVisiblePage < 0 || this.endVisiblePage < 0 || !this.pageDetector)
				return null;

			let x = xInp - this.x;
			let y = yInp - this.y;

			var pageCoords = null;
			var pageIndex = 0;
			for (pageIndex = this.startVisiblePage; pageIndex <= this.endVisiblePage; pageIndex++)
			{
				pageCoords = this.pageDetector.pages[pageIndex - this.startVisiblePage];
				if (y >= pageCoords.y / AscCommon.AscBrowser.retinaPixelRatio && y <= (pageCoords.y + pageCoords.h) / AscCommon.AscBrowser.retinaPixelRatio)
					break;
			}
			if (pageIndex > this.endVisiblePage)
				pageIndex = this.endVisiblePage;

			let oDoc = this.getPDFDoc();

			let _x = oDoc.pagesTransform[pageIndex].normal.TransformPointX(x, y);
			let _y = oDoc.pagesTransform[pageIndex].normal.TransformPointY(x, y);

			return {
				index : pageIndex,
				x : _x,
				y : _y
			};
		};

		this.getPageByCoords2 = function(xInp, yInp)
		{
			if (this.startVisiblePage < 0 || this.endVisiblePage < 0 || !this.pageDetector)
				return null;

			let x = xInp - this.x;
			let y = yInp - this.y;

			var pageCoords = null;
			var pageIndex = 0;
			for (pageIndex = this.startVisiblePage; pageIndex <= this.endVisiblePage; pageIndex++)
			{
				pageCoords = this.pageDetector.pages[pageIndex - this.startVisiblePage];
				if (y >= pageCoords.y / AscCommon.AscBrowser.retinaPixelRatio && y <= (pageCoords.y + pageCoords.h) / AscCommon.AscBrowser.retinaPixelRatio)
					break;
			}
			if (pageIndex > this.endVisiblePage)
				pageIndex = this.endVisiblePage;

			let oDoc = this.getPDFDoc();

			let _x = oDoc.pagesTransform[pageIndex].normal.TransformPointX(x, y) * g_dKoef_pt_to_mm;
			let _y = oDoc.pagesTransform[pageIndex].normal.TransformPointY(x, y) * g_dKoef_pt_to_mm;

			return {
				index : pageIndex,
				x : _x,
				y : _y
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
			let oDoc	= Asc.editor.getPDFDoc();
			let oFile	= Asc.editor.getDocumentRenderer().file;

			let oTr		= oDoc.GetPageTransform(pageIndex, true).normal;
			let inchC	= (25.4 / oFile.pages[pageIndex].Dpi);
			AscCommon.global_MatrixTransformer.ScaleAppend(oTr, inchC, inchC);
			oTr.Invert();

			let oPt = oTr.TransformPoint(x, y);

			return ( {x : oPt.x, y : oPt.y, w : oDoc.GetPageWidthMM(pageIndex) / inchC , h: oDoc.GetPageHeightMM(pageIndex) / inchC} );
		};

		this.getPageLikeDetector = function(pageIndex)
		{
			let curPosition = this.getViewportPosition();

			let page = this.drawingPages[pageIndex];
			let w = AscCommon.AscBrowser.convertToRetinaValue(page.W, true);
			let h = AscCommon.AscBrowser.convertToRetinaValue(page.H, true);

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
			var pageObjectLogic = this.getPageByCoords2(AscCommon.global_mouseEvent.X, AscCommon.global_mouseEvent.Y);
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

		this.findText = function(props, isNext, callback)
		{
			let oDoc = this.getPDFDoc();

			if (!this.isFullText)
			{
				this.fullTextMessageCallbackArgs = [props, isNext, callback];
				this.fullTextMessageCallback = function() {
					let oSearchEnginge = oDoc.Search(this.fullTextMessageCallbackArgs[0], this.fullTextMessageCallbackArgs[1], this.fullTextMessageCallbackArgs[2]);
					let nCurrentMatch = oDoc.GetSearchElementId(true);
					oDoc.SelectSearchElement(nCurrentMatch);
					
					this.onUpdateOverlay();

					if (this.fullTextMessageCallbackArgs[4])
						this.fullTextMessageCallbackArgs[4].call(this.Api, nCurrentMatch, oSearchEnginge.Count);
				};
				this.showTextMessage();
				return true; // async
			}

			oDoc.Search(props, isNext);
			let nCurrentMatch = oDoc.GetSearchElementId(isNext);
			oDoc.SelectSearchElement(nCurrentMatch);

			this.onUpdateOverlay();
			return false;
		};

		this.ToSearchResult = function()
		{
			var naviG = this.CurrentSearchNavi;
			if (!naviG) {
				return null;
			}
			
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
			var bRetValue	= false;
			let oDoc		= this.getPDFDoc();
			let oDrDoc		= oDoc.GetDrawingDocument();

			let shortcutType = this.Api.getShortcut(e);
			if (oDoc.executeShortcut(shortcutType))
			{
				bRetValue = keydownresult_PreventAll;
			}
			else if (e.KeyCode === 8) // BackSpace
			{
				oDoc.DoAction(function() {
					oDoc.Remove(-1, e.CtrlKey == true);
				}, AscDFH.historydescription_Document_BackSpaceButton);

				bRetValue = true;
			}
			else if (e.KeyCode === 9) // Tab
			{
				window.event.preventDefault();

				let oActiveObj = oDoc.GetActiveObject();
				if (oActiveObj && oActiveObj.IsDrawing()) {
					if (oActiveObj.IsGraphicFrame()) {
						oActiveObj.MoveCursorToCell(e.ShiftKey ? false : true);
						if (false == AscCommon.History.Is_LastPointEmpty()) {
							oActiveObj.SetNeedRecalc(true);
						}

						oDrDoc.TargetStart(true);
						oDoc.SetNeedUpdateTarget(true);
						this._checkTargetUpdate();
					}
					else {
						oDoc.AddToParagraph(new AscWord.CRunTab());
					}
				}
				else {
					if (true == e.ShiftKey)
						this.SelectPrevForm();
					else
						this.SelectNextForm();
				}

				bRetValue = true;
			}
			else if (e.KeyCode === 13) // Enter
			{
				window.event.stopPropagation();
				oDoc.EnterDown(e.ShiftKey === true);
				bRetValue = true;
			}
			else if (e.KeyCode === 27) // Esc
			{
				if (this.Api.isInkDrawerOn())
				{
					this.Api.stopInkDrawer();
				}
				else if (this.Api.isMarkerFormat)
				{
					this.Api.sendEvent("asc_onMarkerFormatChanged", this.Api.curMarkerType, false);
					this.Api.SetMarkerFormat(this.Api.curMarkerType, false);
				}
				else if (this.Api.isStartAddShape)
				{
					this.Api.sync_StartAddShapeCallback(false);
					this.Api.sync_EndAddShape();
					this.DrawingObjects.endTrackNewShape();
				}
				else {
					const oController = oDoc.GetController();
					oController.resetSelection();
					this.onUpdateOverlay();
					oDoc.EscapeForm();
				
					editor.sync_HideComment();
				}

				bRetValue = true;
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

				bRetValue = true;
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

				bRetValue = true;
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
				if (oDoc.activeForm || oDoc.mouseDownAnnot || oDoc.activeDrawing) {
					oDoc.MoveCursorLeft(true === e.ShiftKey, true === e.CtrlKey);
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
				if (oDoc.activeForm || oDoc.mouseDownAnnot || oDoc.activeDrawing) {
					oDoc.MoveCursorUp(true === e.ShiftKey, true === e.CtrlKey);
				}
				else if (!this.isFocusOnThumbnails) {
					this.m_oScrollVerApi.scrollByY(-40);
				}
				bRetValue = true;
			}
			else if ( e.KeyCode == 39 ) // Right Arrow
			{	
				if (oDoc.activeForm || oDoc.mouseDownAnnot || oDoc.activeDrawing) {
					oDoc.MoveCursorRight(true === e.ShiftKey, true === e.CtrlKey);
				}
				else if (!this.isFocusOnThumbnails && this.isVisibleHorScroll) {
					this.m_oScrollHorApi.scrollByX(40);
				}
				else if (this.isFocusOnThumbnails) {
					if (this.currentPage < (this.getPagesCount() - 1))
						this.navigateToPage(this.currentPage + 1);
				}
				bRetValue = true;
			}
			else if ( e.KeyCode == 40 ) // Bottom Arrow
			{
				if (oDoc.activeForm || oDoc.mouseDownAnnot || oDoc.activeDrawing) {
					oDoc.MoveCursorDown(true === e.ShiftKey, true === e.CtrlKey);
				}
				else if (!this.isFocusOnThumbnails)
				{
					this.m_oScrollVerApi.scrollByY(40);
				}
				bRetValue = true;
			}
			else if (e.KeyCode === 46) // Delete
			{
				oDoc.DoAction(function() {
					oDoc.Remove(1, e.CtrlKey == true);
				}, AscDFH.historydescription_Document_DeleteButton);

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

			if (page.Rotate && ismultiply)
				page.Rotate += angle;
			else
				page.Rotate = angle;

			page.Rotate += 4;
			page.Rotate &= 0x03;

			if (0 === page.Rotate)
				delete page.Rotate;

			this.resize();
			this.thumbnails && this.thumbnails.resize();
		};

		this.getPageRotate = function(pageNum)
		{
			if (!this.file || !this.file.isValid())
				return 0;

			if (!this.file.pages[pageNum])
				return 0;

			let value = this.file.pages[pageNum].Rotate;
			return (undefined === value) ? 0 : value;
		};
		this.getDrawingPageScale = function(pageNum)
		{
			return 96 / this.file.pages[pageNum].Dpi;
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
		
		for (let i = this.startVisiblePage; i <= this.endVisiblePage; i++)
		{
			let page = this.drawingPages[i];
			if (!page)
				break;

			let aForms = this.pagesInfo.pages[i].fields != null ? this.pagesInfo.pages[i].fields : null;
			if (aForms.length == 0)
				continue;

			let w = AscCommon.AscBrowser.convertToRetinaValue(page.W, true);
			let h = AscCommon.AscBrowser.convertToRetinaValue(page.H, true);

			let isStretchPaint	= this.isStretchPaint();
			let isZoom			= page.ImageForms && (page.ImageForms.width != w || page.ImageForms.height != h);
			let isNeedRedraw	= this.pagesInfo.pages[i].needRedrawForms;

			let bDrawEmpty = false;
			if ((null == page.ImageForms || isNeedRedraw || isZoom)) {
				if (isStretchPaint) {
					if (isNeedRedraw) {
						bDrawEmpty = true;
					}
				}
				else {
					// рисуем на отдельном канвасе, кешируем
					let tmpCanvas		= page.ImageForms ? page.ImageForms : document.createElement('canvas');
					let tmpCanvasCtx	= tmpCanvas.getContext('2d');
					
					tmpCanvas.width		= w;
					tmpCanvas.height	= h;

					if (page.ImageForms)
						tmpCanvasCtx.clearRect(0, 0, w, h);

					this._drawFieldsOnCtx(i, tmpCanvasCtx);
					
					page.ImageForms = tmpCanvas;
					this.pagesInfo.pages[i].needRedrawForms = false;
				}
			}
			
			this.blitPageToCtx(ctx, bDrawEmpty ? null : page.ImageForms, i);
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
	CHtmlPage.prototype._checkFontsOnPages = function(nStart, nEnd) {
		let oDoc = this.getPDFDoc();
		let fontMap = {};

		for (let i = nStart; i <= nEnd; i++)
		{
			let page = this.drawingPages[i];
			if (!page)
				break;

			let aForms		= this.pagesInfo.pages[i].fields != null ? this.pagesInfo.pages[i].fields : [];
			let aDrawings	= this.pagesInfo.pages[i].drawings != null ? this.pagesInfo.pages[i].drawings : [];
			let aFreeText	= this.pagesInfo.pages[i].annots != null ? this.pagesInfo.pages[i].annots.filter(function(annot) {
				return annot.IsFreeText();
			}): [];

			if (this.pagesInfo.pages[i].needRedrawForms)
			{
				aForms.forEach(function(field) {
					if (field.IsNeedDrawFromStream() == false) {
						let sFont = field.GetTextFontActual();
						if (sFont)
							fontMap[sFont] = true;
					}
				});
			}
			
			if (this.pagesInfo.pages[i].needRedrawDrawings)
			{
				aDrawings.forEach(function(drawing) {
					drawing.GetAllFonts(fontMap);
				});
			}
			
			if (this.pagesInfo.pages[i].needRedrawAnnots)
			{
				aFreeText.forEach(function(annot) {
					annot.GetAllFonts(fontMap);
				});
			}
		}

		let oThis = this;
		return oDoc.checkFonts(Object.keys(fontMap), function() {
			oThis.scheduleRepaint();
		});
	};
	CHtmlPage.prototype._paintAnnots = function()
	{
		const ctx = this.canvasForms.getContext('2d');
		ctx.clearRect(0, 0, this.canvasForms.width, this.canvasForms.height);
		ctx.globalAlpha = 1;
		
		for (let i = this.startVisiblePage; i <= this.endVisiblePage; i++)
		{
			let page = this.drawingPages[i];
			if (!page)
				break;

			let aAnnots = this.pagesInfo.pages[i].annots != null ? this.pagesInfo.pages[i].annots : null;
			if (aAnnots.length == 0)
				continue;

			let w = AscCommon.AscBrowser.convertToRetinaValue(page.W, true);
			let h = AscCommon.AscBrowser.convertToRetinaValue(page.H, true);

			let isStretchPaint	= this.isStretchPaint();
			let isZoom			= page.ImageAnnots && (page.ImageAnnots.width != w || page.ImageAnnots.height != h);
			let isNeedRedraw	= this.pagesInfo.pages[i].needRedrawAnnots;

			let bDrawEmpty = false;
			if ((null == page.ImageAnnots || isNeedRedraw || isZoom)) {
				if (isStretchPaint) {
					if (isNeedRedraw) {
						bDrawEmpty = true;
					}
				}
				else {
					// рисуем на отдельном канвасе, кешируем
					let tmpCanvas		= page.ImageAnnots ? page.ImageAnnots : document.createElement('canvas');
					let tmpCanvasCtx	= tmpCanvas.getContext('2d');
					
					tmpCanvas.width		= w;
					tmpCanvas.height	= h;

					if (page.ImageAnnots)
						tmpCanvasCtx.clearRect(0, 0, w, h);

					this._drawAnnotsOnCtx(i, tmpCanvasCtx);
					
					page.ImageAnnots = tmpCanvas;
					this.pagesInfo.pages[i].needRedrawAnnots = false;
				}
			}
			
			this.blitPageToCtx(ctx, bDrawEmpty ? null : page.ImageAnnots, i);
		}
		
		if (this.doc.mouseDownAnnot && this.doc.mouseDownAnnot.IsFreeText()) {
			let oContent = this.doc.mouseDownAnnot.GetDocContent();
			if (oContent.IsSelectionUse() === false)
				oContent.RecalculateCurPos();
		}
	};
	CHtmlPage.prototype._drawMarkupAnnotsOnCtx = function(nPage, ctx)
	{
        let aAnnots = this.pagesInfo.pages[nPage].annots.filter(function(annot) {
			return annot.IsTextMarkup();
		});

		if (aAnnots.length == 0)
			return;
		
		let page = this.drawingPages[nPage];
		if (!page)
			return;

		let widthPx		= ctx.canvas.width;
		let heightPx	= ctx.canvas.height;
		
		let oGraphicsPDF = new AscPDF.CPDFGraphics();
		oGraphicsPDF.Init(ctx, widthPx, heightPx, this.file.getPageWidth(nPage) , this.file.getPageHeight(nPage));
		oGraphicsPDF.SetCurPage(nPage);

		aAnnots.forEach(function(annot) {
			if (false == annot.IsNeedDrawFromStream()) {
				annot.Draw(oGraphicsPDF);
			}
			else {
				annot.DrawFromStream(oGraphicsPDF);
			}
		});
	};
	CHtmlPage.prototype._paintDrawings = function() {
		const ctx = this.canvas.getContext('2d');
		ctx.globalAlpha = 1;
		
		for (let i = this.startVisiblePage; i <= this.endVisiblePage; i++)
		{
			let page = this.drawingPages[i];
			if (!page)
				break;

			let aDrawings = this.pagesInfo.pages[i].drawings != null ? this.pagesInfo.pages[i].drawings : null;
			if (aDrawings.length == 0)
				continue;

			let w = AscCommon.AscBrowser.convertToRetinaValue(page.W, true);
			let h = AscCommon.AscBrowser.convertToRetinaValue(page.H, true);

			let isStretchPaint	= this.isStretchPaint();
			let isZoom			= page.ImageDrawings && (page.ImageDrawings.width != w || page.ImageDrawings.height != h);
			let isNeedRedraw	= this.pagesInfo.pages[i].needRedrawDrawings;

			let bDrawEmpty = false;
			if ((null == page.ImageDrawings || isNeedRedraw || isZoom)) {
				if (isStretchPaint) {
					if (isNeedRedraw) {
						bDrawEmpty = true;
					}
				}
				else {
					// рисуем на отдельном канвасе, кешируем
					let tmpCanvas		= page.ImageDrawings ? page.ImageDrawings : document.createElement('canvas');
					let tmpCanvasCtx	= tmpCanvas.getContext('2d');
					
					tmpCanvas.width		= w;
					tmpCanvas.height	= h;

					if (page.ImageDrawings)
						tmpCanvasCtx.clearRect(0, 0, w, h);

					this._drawDrawingsOnCtx(i, tmpCanvasCtx);
					
					page.ImageDrawings = tmpCanvas;
					this.pagesInfo.pages[i].needRedrawDrawings = false;
				}
			}
			
			this.blitPageToCtx(ctx, bDrawEmpty ? null : page.ImageDrawings, i);
		}
		
		if (this.doc.activeDrawing) {
			let oContent = this.doc.activeDrawing.GetDocContent();
			if (oContent && oContent.IsSelectionUse() === false) {
				oContent.RecalculateCurPos();
			}
		}
	};
	CHtmlPage.prototype._paintFormsHighlight = function()
	{
		let oDrDoc = this.getPDFDoc().GetDrawingDocument();
		let oCtx = this.canvasForms.getContext("2d");
		oCtx.save();

		for (let i = this.startVisiblePage; i <= this.endVisiblePage; i++)
		{
			let page = this.drawingPages[i];
			if (!page)
				break;

			let aForms = this.pagesInfo.pages[i].fields;
			
			if (aForms.length == 0)
				continue;
			
			oDrDoc.AutoShapesTrack.SetCurrentPage(i, true);
			
			let oTr = this.overlay.CanvasTransform;
			oCtx.setTransform(oTr.sx,oTr.shy,oTr.shx,oTr.sy,oTr.tx,oTr.ty);

			if (this.pagesInfo.pages[i].fields != null) {
				this.pagesInfo.pages[i].fields.forEach(function(field) {
					if (field.IsNeedDrawHighlight())
						field.DrawHighlight(oCtx);

					// маркеры
					if (field.GetType() == AscPDF.FIELD_TYPES.combobox)
						field.DrawMarker(oCtx);
					else if (field.GetType() == AscPDF.FIELD_TYPES.text && field.IsDateFormat()) {
						field.IsNeedDrawHighlight() == false && field.DrawDateMarker(oCtx);
					}
				});
			}
		}

		oCtx.restore();
	};
	CHtmlPage.prototype._paintFormsMarkers = function()
	{
		return;
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
					else if (field.GetType() == AscPDF.FIELD_TYPES.text && field.IsDateFormat()) {
						field.IsNeedDrawHighlight() == false && field.DrawDateMarker(oCtx);
					}
				});
			}
		}
	};
	// возвращает видимый рект страницы (процентах от полной), не учитывая поворот
	CHtmlPage.prototype.getViewingRect = function(nPage) {
		let oPageDetector = this.pageDetector;

		let page = oPageDetector.pages.find(function(page) {
			return page.num == nPage;
		});

		if (!page)
		{
			return {
				num : nPage,
				x : 0,
				y : 0,
				r : 0,
				b : 0
			};
		}

		let x = 0;
		if (page.x < 0) 
			x = -page.x / page.w;

		let y = 0;
		if (page.y < 0)
			y = -page.y / page.h;

		let r = 1;
		if ((page.x + page.w) > oPageDetector.width)
			r -= (page.x + page.w - oPageDetector.width) / page.w;

		let b = 1;
		if ((page.y + page.h) > oPageDetector.height)
			b -= (page.y + page.h - oPageDetector.height) / page.h;
		
		return {
			num : nPage,
			x : x,
			y : y,
			r : r,
			b : b
		};
	};
	// возвращает видимый рект страницы (процентах от полной) учитывая поворот
	CHtmlPage.prototype.getViewingRect2 = function(nPage) {
		let oViewRect = this.getViewingRect(nPage);
		let nRotAngle = this.getPageRotate(nPage);

		let oRotViewRect = {};
		switch (nRotAngle) {
			case 90: {
				oRotViewRect.x = oViewRect.y;
				oRotViewRect.y = 1 - oViewRect.r;
				oRotViewRect.r = oViewRect.b;
				oRotViewRect.b = 1 - oViewRect.x;
				break;
			}
			case 180: {
				oRotViewRect.x = 1 - oViewRect.r;
				oRotViewRect.y = 1 - oViewRect.b;
				oRotViewRect.r = 1 - oViewRect.x;
				oRotViewRect.b = 1 - oViewRect.y;
				break;
			}
			case 270: {
				oRotViewRect.x = 1 - oViewRect.b;
				oRotViewRect.y = oViewRect.x;
				oRotViewRect.r = 1 - oViewRect.y;
				oRotViewRect.b = oViewRect.r;
				break;
			}
			default: {
				oRotViewRect = oViewRect;
				break;
			}
		}

		return oRotViewRect;
	};
	CHtmlPage.prototype.GetPageForThumbnails = function(nPage, nWidthPx, nHeightPx) {
		let oFile = this.file;
		let image = !oFile.pages[nPage].isRecognized ? this.file.getPage(nPage, nWidthPx, nHeightPx, undefined, this.Api.isDarkMode ? 0x3A3A3A : 0xFFFFFF) : null;

		if (!image) {
			let pageColor = this.Api.getPageBackgroundColor();

			image = document.createElement('canvas');

			let ctx = image.getContext('2d');

			image.width = nWidthPx;
			image.height = nHeightPx;

			ctx.fillStyle = "rgba(" + pageColor.R + "," + pageColor.G + "," + pageColor.B + ",1)";
			ctx.fillRect(0, 0, nWidthPx, nHeightPx);
		}

		image.requestWidth = nWidthPx;
		image.requestHeight = nHeightPx;

		let ctx = image.getContext('2d');

		this._drawDrawingsOnCtx(nPage, ctx, true);
		this._drawMarkupAnnotsOnCtx(nPage, ctx);
		this._drawAnnotsOnCtx(nPage, ctx, true);
		this._drawFieldsOnCtx(nPage, ctx, true);

		return ctx.canvas;
	};
	CHtmlPage.prototype.GetPrintPage = function(nPage, nWidthPx, nHeightPx) {
		let oFile = this.file;
		let image = !oFile.pages[nPage].isRecognized ? this.file.getPage(nPage, nWidthPx, nHeightPx, undefined, 0xFFFFFF) : null;

		if (!image) {
			let pageColor = this.Api.getPageBackgroundColor();

			image = document.createElement('canvas');

			let ctx = image.getContext('2d');

			image.width = nWidthPx;
			image.height = nHeightPx;

			ctx.fillStyle = "rgba(" + pageColor.R + "," + pageColor.G + "," + pageColor.B + ",1)";
			ctx.fillRect(0, 0, nWidthPx, nHeightPx);
		}

		image.requestWidth = nWidthPx;
		image.requestHeight = nHeightPx;

		let ctx = image.getContext('2d');

		this._drawDrawingsOnCtx(nPage, ctx);
		this._drawMarkupAnnotsOnCtx(nPage, ctx);
		this._drawAnnotsOnCtx(nPage, ctx);
		this._drawFieldsOnCtx(nPage, ctx);

		return ctx.canvas;
	};
    CHtmlPage.prototype._drawAnnotsOnCtx = function(nPage, ctx, isThumbnails) {
		let oDoc		= this.getPDFDoc();
        let widthPx		= ctx.canvas.width;
        let heightPx	= ctx.canvas.height;
        
		let oGraphicsPDF = new AscPDF.CPDFGraphics();
		oGraphicsPDF.isThumbnails = isThumbnails;
        oGraphicsPDF.Init(ctx, widthPx, heightPx, this.file.getPageWidth(nPage) , this.file.getPageHeight(nPage));
        oGraphicsPDF.SetCurPage(nPage);

        let oGraphicsWord = new AscCommon.CGraphics();
        oGraphicsWord.init(ctx, widthPx, heightPx, oDoc.GetPageWidthMM(nPage), oDoc.GetPageHeightMM(nPage));
        oGraphicsWord.m_oFontManager = AscCommon.g_fontManager;
        oGraphicsWord.setEndGlobalAlphaColor(255, 255, 255);
        oGraphicsWord.transform(1, 0, 0, 1, 0, 0);
        
        if (this.pagesInfo.pages[nPage].annots != null) {
            this.pagesInfo.pages[nPage].annots.forEach(function(annot) {
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
    };
    CHtmlPage.prototype._drawFieldsOnCtx = function(nPage, ctx, isThumbnails) {
		let oDoc		= this.getPDFDoc();
        let widthPx		= ctx.canvas.width;
        let heightPx    = ctx.canvas.height;
        
        let oGraphicsPDF = new AscPDF.CPDFGraphics();
		oGraphicsPDF.isThumbnails = isThumbnails;
        oGraphicsPDF.Init(ctx, widthPx, heightPx, this.file.getPageWidth(nPage) , this.file.getPageHeight(nPage));
        oGraphicsPDF.SetCurPage(nPage);

        let oGraphicsWord = new AscCommon.CGraphics();
        oGraphicsWord.init(ctx, widthPx, heightPx, oDoc.GetPageWidthMM(nPage) , oDoc.GetPageHeightMM(nPage));
        oGraphicsWord.m_oFontManager = AscCommon.g_fontManager;
        oGraphicsWord.setEndGlobalAlphaColor(255, 255, 255);
        oGraphicsWord.transform(1, 0, 0, 1, 0, 0);
        
        if (this.pagesInfo.pages[nPage].fields != null) {
            this.pagesInfo.pages[nPage].fields.forEach(function(field) {
                field.DrawOnPage(oGraphicsPDF, oGraphicsWord, nPage);
            });
        }
    };
	CHtmlPage.prototype._drawDrawingsOnCtx = function(nPage, ctx, isThumbnails) {
		let aDrawings = this.pagesInfo.pages[nPage].drawings;
		if (aDrawings.length == 0) {
			return;
		}

		let oDoc		= this.getPDFDoc();
		let widthPx		= ctx.canvas.width;
		let heightPx    = ctx.canvas.height;

		let oGraphicsWord = new AscCommon.CGraphics();
		oGraphicsWord.isThumbnails = isThumbnails;
		oGraphicsWord.init(ctx, widthPx, heightPx, oDoc.GetPageWidthMM(nPage) , oDoc.GetPageHeightMM(nPage));
		oGraphicsWord.m_oFontManager = AscCommon.g_fontManager;
		oGraphicsWord.setEndGlobalAlphaColor(255, 255, 255);
		oGraphicsWord.transform(1, 0, 0, 1, 0, 0);

		aDrawings.forEach(function(drawing) {
			drawing.Draw(oGraphicsWord);
		});
	};
	CHtmlPage.prototype.createComponents = function()
	{
		if (AscCommon.g_inputContext)
			AscCommon.g_inputContext = null;

		var elements = "<div id=\"id_main\" class=\"block_elem\" style=\"touch-action:none;-ms-touch-action:none;-moz-user-select:none;user-select:none;-webkit-user-select:none;background-color:" + AscCommon.GlobalSkin.BackgroundColor + ";overflow:hidden;\" UNSELECTABLE=\"on\">";
		elements += "<canvas id=\"id_viewer\" class=\"block_elem\" style=\"touch-action:none;-ms-touch-action:none;-moz-user-select:none;user-select:none;-webkit-user-select:none;left:0px;top:0px;width:100;height:100;\"></canvas>";
		elements += "<canvas id=\"id_forms\" class=\"block_elem\" style=\"touch-action:none;-ms-touch-action:none;-moz-user-select:none;user-select:none;-webkit-user-select:none;left:0px;top:0px;width:100;height:100;\"></canvas>";
		elements += "<div id=\"id_target_cursor\" class=\"block_elem\" width=\"1\" height=\"1\" style=\"touch-action:none;-ms-touch-action: none;-webkit-user-select: none;width:2px;height:13px;z-index:4;\"></div>"
		elements += "<canvas id=\"id_overlay\" class=\"block_elem\" style=\"touch-action:none;-ms-touch-action:none;-moz-user-select:none;user-select:none;-webkit-user-select:none;left:0px;top:0px;width:100;height:100;\"></canvas>";
		elements += "</div>";

		elements += "<div id=\"id_vertical_scroll\" class=\"block_elem\" style=\"display:none;left:0px;top:0px;width:0px;height:0px;\"></div>";
		elements += "<div id=\"id_horizontal_scroll\" class=\"block_elem\" style=\"display:none;left:0px;top:0px;width:0px;height:0px;\"></div>";
		//this.parent.style.backgroundColor = this.backgroundColor; <= this color from theme
		this.parent.innerHTML = elements;
		
		let oControl = editor.WordControl.m_oBody.Controls.find(function(control) {
			return control.HtmlElement.id == "id_main";
		});

		if (!oControl)
			oControl = {};

		oControl.HtmlElement = document.getElementById("id_main");
		
		this.id_main = oControl.HtmlElement;

		this.canvas = document.getElementById("id_viewer");
		this.canvas.style.backgroundColor = this.backgroundColor;
		
		this.canvasOverlay = document.getElementById("id_overlay");
		this.canvasOverlay.style.pointerEvents = "none";
		
		this.canvasForms = document.getElementById("id_forms");
		
		this.Api.WordControl.m_oDrawingDocument.TargetHtmlElement = document.getElementById('id_target_cursor');
		this.Api.WordControl.m_oDrawingDocument.m_oWordControl.m_oMainView = oControl;
		
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
		if (window["NATIVE_EDITOR_ENJINE"])
			return;

		AscCommon.AscBrowser.checkZoom();
		if (isDisablePaint) {
			this.disabledPaintOnScroll = true;
		}

		let oThis		= this;
		let oEditorPage	= this.Api.WordControl;

		this.isFocusOnThumbnails = false;

		oEditorPage.checkBodySize();
		oEditorPage.m_oBody.Resize(oEditorPage.Width * g_dKoef_pix_to_mm, oEditorPage.Height * g_dKoef_pix_to_mm, this);

		var rect = AscCommon.UI.getBoundingClientRect(this.canvas);
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
		
		if (this.zoomCoordinate && this.isDocumentContentReady && Asc.editor.isDocumentLoadComplete)
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

		this.UpdateDrDocDrawingPages();
		if (true !== isDisablePaint) {
			this._paint();
			this.onUpdateOverlay();	
		}
		else {
			this.getPDFDoc().UpdatePagesTransform();
		}

		if (isDisablePaint) {
			this.disabledPaintOnScroll = false;
		}
	};
	CHtmlPage.prototype.repaintFormsOnPage = function(pageIndex)
	{
		if (this.pagesInfo.pages.length > pageIndex)
		{
			this.pagesInfo.pages[pageIndex].needRedrawForms = true;
			this.isRepaint = true;
		}
	};
	CHtmlPage.prototype.InitDocRenderer = function(oMemory, nPage) {
		let oDoc        = this.getPDFDoc();
        let oRenderer   = new AscCommon.CDocumentRenderer();
		oRenderer.InitPicker(AscCommon.g_oTextMeasurer.m_oManager);

		oRenderer.Memory		= oMemory;
		oMemory.docRenderer	= oRenderer;

        oMemory.context = new AscCommon.XmlWriterContext(AscCommon.c_oEditorId.Presentation);
        oMemory.context.docType	= AscFormat.XMLWRITER_DOC_TYPE_PPTX;

        oRenderer.m_arrayPages[oRenderer.m_arrayPages.length]						= new AscCommon.CMetafile(oDoc.GetPageWidthMM(nPage), oDoc.GetPageHeightMM(nPage));
        oRenderer.m_lPagesCount														= oRenderer.m_arrayPages.length;
        oRenderer.m_arrayPages[oRenderer.m_lPagesCount - 1].Memory					= oRenderer.Memory;
        oRenderer.m_arrayPages[oRenderer.m_lPagesCount - 1].StartOffset				= oRenderer.Memory.pos;
        oRenderer.m_arrayPages[oRenderer.m_lPagesCount - 1].VectorMemoryForPrint	= oRenderer.VectorMemoryForPrint;
        oRenderer.m_arrayPages[oRenderer.m_lPagesCount - 1].FontPicker				= oRenderer.FontPicker;

        oRenderer.m_arrayPages[oRenderer.m_lPagesCount - 1].FontPicker				= oRenderer.FontPicker;

        if (oRenderer.FontPicker)
            oRenderer.m_arrayPages[oRenderer.m_lPagesCount - 1].FontPicker.Metafile  = oRenderer.m_arrayPages[oRenderer.m_lPagesCount - 1];

        let _page = oRenderer.m_arrayPages[oRenderer.m_lPagesCount - 1];
        oRenderer.m_oPen       = _page.m_oPen;
        oRenderer.m_oBrush     = _page.m_oBrush;
        oRenderer.m_oTransform = _page.m_oTransform;

		return oRenderer;
	};

	CHtmlPage.prototype.Save = function()
	{
		let memoryInitSize	= 1024 * 500; // 500Kb
		let oMemory			= null;

		let aPagesInfo	= this.pagesInfo.pages;
		let oFile		= this.file;

		// по информации аннотаций определим какие были удалены
		let oDoc		= this.getPDFDoc();
		oDoc.BlurActiveObject();
		
		let aAnnotsInfo	= oFile.nativeFile["getAnnotationsInfo"]();
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

		// edit		- 0
		// add		- 1
		// delete	- 2

		function writePageInfo(operation, originIndex) {
			if (!oMemory)
			{
				oMemory = new AscCommon.CMemory(true);
				oMemory.Init(memoryInitSize);
				oMemory.images = [];

				// compiled changes signature
				oMemory.WriteByte("%".charCodeAt(0));
				oMemory.WriteByte("P".charCodeAt(0));
				oMemory.WriteByte("D".charCodeAt(0));
				oMemory.WriteByte("F".charCodeAt(0));
			}

			let nCommandType = operation[0];
			let curIndex = operation[1];

			let nStartPos = oMemory.GetCurPosition();
			oMemory.Skip(4);
			oMemory.WriteByte(nCommandType);
			oMemory.WriteLong(originIndex != undefined ? originIndex : curIndex);
			
			if ([AscPDF.CommandType.editPage, AscPDF.CommandType.addPage].includes(nCommandType)) {
				let nRotAngle = this.getPageRotate(curIndex);
				let bClearPage = !!oFile.pages[curIndex].isRecognized;
				
				oMemory.WriteByte(AscCommon.CommandType.ctPageRotate);
				oMemory.WriteLong(8);
				oMemory.WriteLong(nRotAngle);

				// edit page
				if (nCommandType == AscPDF.CommandType.editPage) {
					if (bClearPage) {
						oMemory.WriteByte(AscCommon.CommandType.ctPageClear);
						oMemory.WriteLong(4);
					}
				}
				
				// add page
				if (nCommandType == AscPDF.CommandType.addPage) {
					oMemory.WriteByte(AscCommon.CommandType.ctPageWidth);
					oMemory.WriteDouble(oFile.pages[curIndex].W);
					
					oMemory.WriteByte(AscCommon.CommandType.ctPageHeight);
					oMemory.WriteDouble(oFile.pages[curIndex].H);
				}
			}
			else if (nCommandType == AscPDF.CommandType.movePage) {
				let nNewPos = operation[2];
				oMemory.WriteLong(nNewPos);
				let nEndPos = oMemory.GetCurPosition();
				// длина комманд на стринице
				oMemory.Seek(nStartPos);
				oMemory.WriteLong(nEndPos - nStartPos);
				oMemory.Seek(nEndPos);

				return;
			}
			// remove page
			else if (nCommandType == AscPDF.CommandType.removePage) {
				let nEndPos = oMemory.GetCurPosition();
				// длина комманд на стринице
				oMemory.Seek(nStartPos);
				oMemory.WriteLong(nEndPos - nStartPos);
				oMemory.Seek(nEndPos);
				return;
			}
			
			let oPageInfo = aPagesInfo[curIndex];
			let oRenderer = this.InitDocRenderer(oMemory, curIndex);

			// annots
			if (oPageInfo.annots) {
				for (let nAnnot = 0; nAnnot < oPageInfo.annots.length; nAnnot++) {
					oPageInfo.annots[nAnnot].IsChanged() && oPageInfo.annots[nAnnot].WriteToBinary(oMemory);
					oPageInfo.annots[nAnnot].GetReplies().forEach(function(reply) {
						reply.IsChanged() && reply.WriteToBinary(oMemory); 
					});
				}
			}

			if (aDeleted[originIndex]) {
				for (let j = 0; j < aDeleted[originIndex].length; j++) {
					oMemory.WriteByte(AscCommon.CommandType.ctAnnotFieldDelete);
					oMemory.WriteLong(8);
					oMemory.WriteLong(aDeleted[originIndex][j]);
				}
			}

			// forms
			if (oPageInfo.fields) {
				for (let nForm = 0; nForm < oPageInfo.fields.length; nForm++) {
					if (oPageInfo.fields[nForm].IsChanged())
						oPageInfo.fields[nForm].WriteToBinary(oMemory);
				}
			}

			// drawings
			if (oPageInfo.drawings && oPageInfo.drawings.length != 0) {
				for (let nDr = 0; nDr < oPageInfo.drawings.length; nDr++) {
					let oDrawing = oPageInfo.drawings[nDr];

					if (oDrawing.IsGraphicFrame()) {
						let sTableStyleId = oDrawing.graphicObject.GetTableStyle();
						if (sTableStyleId) {
							let oStyle = oDoc.globalTableStyles.Get(sTableStyleId);
							if (oStyle) {
								let sStyleGUID = oStyle.GetStyleId();
								if (sTableStyleId != undefined) {
									oMemory.context.tableStylesIdToGuid[sTableStyleId] = sStyleGUID;
								}
							}
						}
					}
				}
				
				for (let nShape = 0; nShape < oPageInfo.drawings.length; nShape++) {
					let oTextShape = oPageInfo.drawings[nShape];

					oMemory.WriteByte(167); // shape start

					// тут будет длина комманд
					let nStartPos = oMemory.GetCurPosition();
					oMemory.Skip(4);

					// тут будет длина xml строки
					let nStrLengthPos = oMemory.GetCurPosition();
					oMemory.Skip(4);

					oTextShape.WriteToBinary(oMemory);

					// запись длины xml строки
					let nEndPos = oMemory.GetCurPosition();
					oMemory.Seek(nStrLengthPos);
					oMemory.WriteLong(nEndPos - nStrLengthPos - 4); // вычитаем 4 так как должна быть длина строки, без учета команды длины
					oMemory.Seek(nEndPos);

					// запись длины комманд
					nEndPos = oMemory.GetCurPosition();
					oMemory.Seek(nStartPos);
					oMemory.WriteLong(nEndPos - nStartPos);
					oMemory.Seek(nEndPos);

					// запись графики
					oTextShape.draw(oRenderer); 

					oMemory.WriteByte(168); // shape end
					oMemory.WriteLong(4); // для обратной совместимости
				}
			}

			let nEndPos = oMemory.GetCurPosition();
			// длина комманд на стринице
			oMemory.Seek(nStartPos);
			oMemory.WriteLong(nEndPos - nStartPos);
			oMemory.Seek(nEndPos);
		}

		function generateOperations(originalPageCount, finalPages) {
			var operations = [];
		
			// 1. Collect a list of old page indexes that are needed in the final document
			var finalOriginIndexes = [];
			for (var i = 0; i < finalPages.length; i++) {
				if (finalPages[i].originIndex !== undefined) {
					finalOriginIndexes.push(finalPages[i].originIndex);
				}
			}
		
			// 2. Remove pages that are not present in the finalPages list
			var deletedCount = 0;
			for (var i = 0; i < originalPageCount; i++) {
				if (finalOriginIndexes.indexOf(i) === -1) {
					operations.push([AscPDF.CommandType.removePage, i - deletedCount]);
					deletedCount++;
				}
			}
		
			// 3. Create an array of "remaining" old pages (after virtual deletion)
			var remainingOriginPages = [];
			for (var i = 0; i < originalPageCount; i++) {
				if (finalOriginIndexes.indexOf(i) !== -1) {
					remainingOriginPages.push(i);
				}
			}
		
			// 4. Determine the required order of old pages in the final document
			var finalOldOrder = [];
			for (var i = 0; i < finalPages.length; i++) {
				if (finalPages[i].originIndex !== undefined) {
					finalOldOrder.push(finalPages[i].originIndex);
				}
			}
		
			// 5. Rearrange old pages to match finalOldOrder
			var pagesCopy = remainingOriginPages.slice();
			for (var targetIndex = 0; targetIndex < finalOldOrder.length; targetIndex++) {
				var pageNeeded = finalOldOrder[targetIndex];
				var currentIndex = pagesCopy.indexOf(pageNeeded);
		
				if (currentIndex !== targetIndex) {
					operations.push([AscPDF.CommandType.movePage, currentIndex, targetIndex]);
					pagesCopy.splice(currentIndex, 1);
					pagesCopy.splice(targetIndex, 0, pageNeeded);
				}
			}
		
			// 6. Insert new pages (where originIndex === undefined)
			for (var i = 0; i < finalPages.length; i++) {
				if (finalPages[i].originIndex === undefined) {
					operations.push([AscPDF.CommandType.addPage, i]);
				}
			}
		
			return operations;
		}
		

		function checkNeedEditOrigPage(nPage) {
			let aDrawings		= aPagesInfo[nPage].drawings;
			let aAnnots			= aPagesInfo[nPage].annots;
			let aForms			= aPagesInfo[nPage].fields;
			let nOriginIndex	= oFile.pages[nPage].originIndex;
			let aDeletedObj		= aDeleted[nOriginIndex] || [];
			let nOrigRotAngle	= oFile.pages[nPage].originRotate;
			let nRotAngle		= oFile.pages[nPage].Rotate;

			let bNeedEdit = false;
			bNeedEdit ||= nRotAngle != nOrigRotAngle;
			bNeedEdit ||= aDrawings.length != 0;
			bNeedEdit ||= aAnnots.find(function(annot) {
				let aReplies = annot.GetReplies();

				return annot.IsChanged() || aReplies.find(function(reply) { return reply.IsChanged()});
			});
			bNeedEdit ||= aForms.find(function(form) {

				return form.IsChanged();
			});
			bNeedEdit ||= aDeletedObj.length != 0;
			
			bNeedEdit &&= nOriginIndex != undefined;

			return bNeedEdit;
		}

		// сначала edit исходных страниц
		for (let i = 0; i < aPagesInfo.length; i++) {
			if (checkNeedEditOrigPage(i)) {
				writePageInfo.call(this, [AscPDF.CommandType.editPage, i], oFile.pages[i].originIndex);
			}
		}

		// составляем порядок операций удаления/добавления страниц
		let aOrder = generateOperations(oFile.originalPagesCount, oFile.pages);

		// пишем по порядку
		for (let i = 0; i < aOrder.length; i++) {
			writePageInfo.call(this, aOrder[i], undefined);
		}

		if (oMemory) {
			let nStartPos = oMemory.GetCurPosition();
			oMemory.Skip(4);

			// parents and CO (calculaction order)
			oMemory.WriteByte(AscPDF.CommandType.widgetInfo);
			oMemory.WriteByte(AscCommon.CommandType.ctWidgetsInfo);
			let nPosForLenght = oMemory.GetCurPosition();
			oMemory.Skip(4);

			let aCO = oDoc.GetCalculateInfo().ids;
			oMemory.WriteLong(aCO.length);
			for (let i = 0; i < aCO.length; i++) {
				oMemory.WriteLong(aCO[i]);
			}

			let nPosForParentLenght = oMemory.GetCurPosition();
			oMemory.Skip(4);
			let nParents = 0;
			oDoc.widgetsParents.forEach(function(field) {
				if (field.IsChanged()) {
					nParents++;
					field.WriteToBinaryAsParent(oMemory);
				}
			});

			let nEndPos = oMemory.GetCurPosition();
			
			// количество изменённых родителей
			oMemory.Seek(nPosForParentLenght);
			oMemory.WriteLong(nParents);
			oMemory.Seek(nEndPos);

			// пишем изображения
			oMemory.WriteLong(oMemory.images.length);
			for (let i = 0; i < oMemory.images.length; i++) {
				oMemory.WriteStringA(oMemory.images[i]);
			}

			nEndPos = oMemory.GetCurPosition();

			// длина комманд с информацией о родителях, CO и картинках
			oMemory.Seek(nPosForLenght);
			oMemory.WriteLong(nEndPos - nPosForLenght);
			oMemory.Seek(nEndPos);

			// Общая длина комманд
			oMemory.Seek(nStartPos);
			oMemory.WriteLong(nEndPos - nStartPos);
			oMemory.Seek(nEndPos);

			return new Uint8Array(oMemory.data.buffer, 0, oMemory.GetCurPosition());
		}
			
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
	window["AscPDF"].PropLocker = PropLocker;

})();
