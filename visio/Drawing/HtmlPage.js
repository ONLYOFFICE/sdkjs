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

// Import
var MOVE_DELTA = AscFormat.MOVE_DELTA;

var g_anchor_left          = AscCommon.g_anchor_left;
var g_anchor_top           = AscCommon.g_anchor_top;
var g_anchor_right         = AscCommon.g_anchor_right;
var g_anchor_bottom        = AscCommon.g_anchor_bottom;
var CreateControlContainer = AscCommon.CreateControlContainer;
var CreateControl          = AscCommon.CreateControl;
var global_keyboardEvent   = AscCommon.global_keyboardEvent;
var global_mouseEvent      = AscCommon.global_mouseEvent;
var History                = AscCommon.History;
var g_dKoef_pix_to_mm      = AscCommon.g_dKoef_pix_to_mm;
var g_dKoef_mm_to_pix      = AscCommon.g_dKoef_mm_to_pix;

var Page_Width  = 297;
var Page_Height = 210;

var X_Left_Margin   = 30;  // 3   cm
var X_Right_Margin  = 15;  // 1.5 cm
var Y_Bottom_Margin = 20;  // 2   cm
var Y_Top_Margin    = 20;  // 2   cm

var X_Right_Field  = Page_Width - X_Right_Margin;
var Y_Bottom_Field = Page_Height - Y_Bottom_Margin;

var HIDDEN_PANE_HEIGHT = 1;

var HEADER_HEIGHT = 40 * g_dKoef_pix_to_mm;
var TIMELINE_HEIGHT = 40 * g_dKoef_pix_to_mm;
var TIMELINE_LIST_RIGHT_MARGIN = 23 * g_dKoef_pix_to_mm;
var TIMELINE_HEADER_RIGHT_MARGIN = 18 * g_dKoef_pix_to_mm;

AscCommon.HIDDEN_PANE_HEIGHT = HIDDEN_PANE_HEIGHT;
AscCommon.TIMELINE_HEIGHT = TIMELINE_HEIGHT;
AscCommon.TIMELINE_LIST_RIGHT_MARGIN = TIMELINE_LIST_RIGHT_MARGIN;
AscCommon.TIMELINE_HEADER_RIGHT_MARGIN = TIMELINE_HEADER_RIGHT_MARGIN;

function CEditorPage(api)
{
	// ------------------------------------------------------------------
	this.Name           = "";
	this.EditorType = "draw";

	this.X      = 0;
	this.Y      = 0;
	this.Width  = 10;
	this.Height = 10;

	// body
	this.m_oBody = null;

	// thumbnails
	this.m_oThumbnailsContainer = null;
	this.m_oThumbnailsBack      = null;
	this.m_oThumbnailsSplit		= null;
	this.m_oThumbnails          = null;
	this.m_oThumbnails_scroll   = null;

	//todo remove
	this.m_oNotesApi	   = {};

	// main
	this.m_oMainParent = null;
	this.m_oMainContent = null;

	// right panel (vertical scroll & buttons (page & rulersEnabled))
	this.m_oPanelRight                = null;
	this.m_oPanelRight_buttonRulers   = null;
	this.m_oPanelRight_vertScroll     = null;
	this.m_oPanelRight_buttonPrevPage = null;
	this.m_oPanelRight_buttonNextPage = null;

	// vertical ruler (left panel)
	this.m_oLeftRuler             = null;
	this.m_oLeftRuler_buttonsTabs = null;
	this.m_oLeftRuler_vertRuler   = null;

	// horizontal ruler (top panel)
	this.m_oTopRuler          = null;
	this.m_oTopRuler_horRuler = null;

	this.ScrollWidthPx = 14;

	// main view
	this.m_oMainView                            = null;
	this.m_oEditor                              = null;
	this.m_oOverlay                             = null;
	this.m_oOverlayApi                          = new AscCommon.COverlay();
	this.m_oOverlayApi.m_bIsAlwaysUpdateOverlay = true;

	// scrolls api
	this.m_oScrollHor_   = null;
	this.m_oScrollVer_   = null;
	this.m_oScrollThumb_ = null;
	this.m_nVerticalSlideChangeOnScrollAllow = false;
	this.m_nVerticalSlideChangeOnScrollInterval = 300; // как часто можно менять слайды при вертикальном скролле
	this.m_nVerticalSlideChangeOnScrollLast = -1;
    this.m_nVerticalSlideChangeOnScrollEnabled = false;

	this.m_oScrollHorApi   = null;
	this.m_oScrollVerApi   = null;
	this.m_oScrollThumbApi = null;

	this.StartVerticalScroll     = false;
	this.VerticalScrollOnMouseUp = {SlideNum : 0, ScrollY : 0, ScrollY_max : 0};

	// properties
	this.m_bDocumentPlaceChangedEnabled = false;

	this.m_nZoomValue = 100;
	this.zoom_values  = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 320, 340, 360, 380, 400, 425, 450, 475, 500];
	this.m_nZoomType  = 2; // 0 - custom, 1 - fitToWodth, 2 - fitToPage

	this.m_oBoundsController = new AscFormat.CBoundsController();
	this.m_nTabsType         = tab_Left;

	// position
	this.m_dScrollY     = 0;
	this.m_dScrollX     = 0;
	this.m_dScrollY_max = 1;
	this.m_dScrollX_max = 1;

	this.m_dScrollX_Central   = 0;
	this.m_dScrollY_Central   = 0;
	this.m_bIsRePaintOnScroll = true;

	this.m_dDocumentWidth      = 0;
	this.m_dDocumentHeight     = 0;
	this.m_dDocumentPageWidth  = 0;
	this.m_dDocumentPageHeight = 0;

	this.m_bIsHorScrollVisible = false;
	this.m_bIsScroll = false;

	// rulers
	this.m_bIsRuler = false;

	this.m_oHorRuler                     = new CHorRuler();
	this.m_oHorRuler.IsCanMoveMargins    = false;
	this.m_oHorRuler.IsCanMoveAnyMarkers = false;
	this.m_oHorRuler.IsDrawAnyMarkers    = false;

	this.m_oVerRuler                  = new CVerRuler();
	this.m_oVerRuler.IsCanMoveMargins = false;

	this.m_oHorRuler.m_oWordControl = this;
	this.m_oVerRuler.m_oWordControl = this;

	this.m_bIsUpdateHorRuler       = false;
	this.m_bIsUpdateVerRuler       = false;

	this.IsEnabledRulerMarkers = false;

	// drawing document
	this.m_oDrawingDocument = new AscCommon.CDrawingDocument();
	this.m_oLogicDocument   = null;

	this.m_oDrawingDocument.m_oWordControl           = this;
	this.m_oDrawingDocument.TransitionSlide.HtmlPage = this;
	this.m_oDrawingDocument.m_oLogicDocument         = this.m_oLogicDocument;

	// flags
	this.m_bIsUpdateTargetNoAttack = false;
	this.arrayEventHandlers = [];

	this.m_oTimerScrollSelect = -1;
	this.IsFocus              = true;
	this.m_bIsMouseLock       = false;
	this.bIsUseKeyPress = true;
	this.bIsEventPaste  = false;
	this.ZoomFreePageNum = -1;
	this.MainScrollsEnabledFlag = 0;
	this.m_bIsIE = AscCommon.AscBrowser.isIE;

	// thumbnails
	this.Thumbnails                 = new CThumbnailsManager();//todo override CThumbnailsManager
	this.Thumbnails.showContextMenu = function(bPosBySelect) {}
	this.Thumbnails.onKeyDown = function(e) {return true;}

	// сплиттеры (для табнейлов и для заметок)
	this.Splitter1Pos    = 0;
    this.Splitter1PosSetUp = 0;
	this.Splitter1PosMin = 0;
	this.Splitter1PosMax = 0;

	this.Splitter2Pos    = 0;
	this.Splitter2PosMin = 0;
	this.Splitter2PosMax = 0;

	this.Splitter3Pos = 0;
	this.Splitter3PosMin = 0;
	this.Splitter3PosMax = 0;

	this.SplitterDiv  = null;
	this.SplitterType = 0;

	this.OldSplitter1Pos = 0;
	this.OldSplitter2Pos = 0;
	this.OldSplitter3Pos = 0;

	this.IsUseNullThumbnailsSplitter = false;

	// axis Y
	this.SlideScrollMIN = 0;
	this.SlideScrollMAX = 0;

	// поддерживает ли браузер нецелые пикселы
	this.bIsDoublePx = AscCommon.isSupportDoublePx();

	this.m_nCurrentTimeClearCache = 0;
	this.IsGoToPageMAXPosition = false;

	this.retinaScaling = AscCommon.AscBrowser.retinaPixelRatio;

	// overlay flags
	this.IsUpdateOverlayOnlyEnd    = false;
	this.IsUpdateOverlayOnEndCheck = false;

	// mobile
	this.MobileTouchManager 			= null;
	this.MobileTouchManagerThumbnails 	= null;

	// draw
	this.SlideDrawer                = new CSlideDrawer();
	this.SlideBoundsOnCalculateSize = new AscFormat.CBoundsController();
	this.DrawingFreeze = false;
	this.NoneRepaintPages = false;

	this.paintMessageLoop = new AscCommon.PaintMessageLoop(40, api);

	this.m_oApi = api;
	var oThis   = this;

	// methods ---

	// controls
	this.checkBodyOffset = function()
	{
		let element = this.m_oBody.HtmlElement;
		if (!element)
			element = document.getElementById(this.Name);
		if (!element)
			return;

		var pos = AscCommon.UI.getBoundingClientRect(element);
		if (pos)
		{
			if (undefined !== pos.x)
				this.X = pos.x;
			else if (undefined !== pos.left)
				this.X = pos.left;

			if (undefined !== pos.y)
				this.Y = pos.y;
			else if (undefined !== pos.top)
				this.Y = pos.top;
		}
	};

	this.checkBodySize = function()
	{
		this.checkBodyOffset();

		var el = document.getElementById(this.Name);

		if (this.Width != el.offsetWidth || this.Height != el.offsetHeight)
		{
			this.Width  = el.offsetWidth;
			this.Height = el.offsetHeight;
			return true;
		}
		return false;
	};

	this.Init = function()
	{
		this.m_oBody = CreateControlContainer(this.Name);
		this.m_oBody.HtmlElement.style.touchAction = "none";

		this.Splitter1PosMin = 20;
		this.Splitter1PosMax = 80;
		this.Splitter2PosMin = 10;
		this.Splitter2PosMax = 100;
		this.Splitter3PosMin = 40;
		this.Splitter3PosMax = 100 - HIDDEN_PANE_HEIGHT;

		this.Splitter1Pos    = 67.5;
		this.Splitter1PosSetUp = this.Splitter1Pos;
		this.Splitter2Pos = 0;
		this.Splitter3Pos = 0;

		this.OldSplitter1Pos = this.Splitter1Pos;
		this.OldSplitter2Pos = this.Splitter2Pos;
		this.OldSplitter3Pos = this.Splitter3Pos;

		var ScrollWidthMm  = this.ScrollWidthPx * g_dKoef_pix_to_mm;
		var ScrollWidthMm9 = 10 * g_dKoef_pix_to_mm;

		this.Thumbnails.m_oWordControl = this;

		// thumbnails -----
		this.m_oThumbnailsContainer = CreateControlContainer("id_panel_thumbnails");
		this.m_oThumbnailsContainer.Bounds.SetParams(0, 0, this.Splitter1Pos, 1000, false, false, true, false, this.Splitter1Pos, -1);
		this.m_oThumbnailsContainer.Anchor = (g_anchor_left | g_anchor_top | g_anchor_bottom);
		this.m_oBody.AddControl(this.m_oThumbnailsContainer);

		this.m_oThumbnailsSplit = CreateControlContainer("id_panel_thumbnails_split");
		this.m_oThumbnailsSplit.Bounds.SetParams(this.Splitter1Pos, 0, 1000, 1000, true, false, false, false, GlobalSkin.SplitterWidthMM, -1);
		this.m_oThumbnailsSplit.Anchor = (g_anchor_left | g_anchor_top | g_anchor_bottom);
		this.m_oBody.AddControl(this.m_oThumbnailsSplit);

		this.m_oThumbnailsBack = CreateControl("id_thumbnails_background");
		this.m_oThumbnailsBack.Bounds.SetParams(0, 0, ScrollWidthMm9, 1000, false, false, true, false, -1, -1);
		this.m_oThumbnailsBack.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right | g_anchor_bottom);
		this.m_oThumbnailsContainer.AddControl(this.m_oThumbnailsBack);

		this.m_oThumbnails = CreateControl("id_thumbnails");
		this.m_oThumbnails.Bounds.SetParams(0, 0, ScrollWidthMm9, 1000, false, false, true, false, -1, -1);
		this.m_oThumbnails.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right | g_anchor_bottom);
		this.m_oThumbnailsContainer.AddControl(this.m_oThumbnails);

		this.m_oThumbnails_scroll = CreateControl("id_vertical_scroll_thmbnl");
		this.m_oThumbnails_scroll.Bounds.SetParams(0, 0, 1000, 1000, false, false, false, false, ScrollWidthMm9, -1);
		this.m_oThumbnails_scroll.Anchor = (g_anchor_top | g_anchor_right | g_anchor_bottom);
		this.m_oThumbnailsContainer.AddControl(this.m_oThumbnails_scroll);
		// ----------------

		if (this.m_oApi.isMobileVersion)
		{
			this.m_oThumbnails_scroll.HtmlElement.style.display = "none";
		}

		// main content -------------------------------------------------------------
		this.m_oMainParent = CreateControlContainer("id_main_parent");
		this.m_oMainParent.Bounds.SetParams(this.Splitter1Pos + GlobalSkin.SplitterWidthMM, 0, g_dKoef_pix_to_mm, 1000, true, false, true, false, -1, -1);
		this.m_oBody.AddControl(this.m_oMainParent);

		this.m_oMainContent = CreateControlContainer("id_main");
		this.m_oMainContent.Bounds.SetParams(0, 0, g_dKoef_pix_to_mm, this.Splitter2Pos + GlobalSkin.SplitterWidthMM, true, false, true, true, -1, -1);

		this.m_oMainContent.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right | g_anchor_bottom);
		this.m_oMainParent.AddControl(this.m_oMainContent);

		// panel right --------------------------------------------------------------
		this.m_oPanelRight = CreateControlContainer("id_panel_right");
		this.m_oPanelRight.Bounds.SetParams(0, 0, 1000, ScrollWidthMm, false, false, false, true, ScrollWidthMm, -1);
		this.m_oPanelRight.Anchor = (g_anchor_top | g_anchor_right | g_anchor_bottom);

		this.m_oMainContent.AddControl(this.m_oPanelRight);

		this.m_oPanelRight_buttonRulers = CreateControl("id_buttonRulers");
		this.m_oPanelRight_buttonRulers.Bounds.SetParams(0, 0, 1000, 1000, false, false, false, false, -1, ScrollWidthMm);
		this.m_oPanelRight_buttonRulers.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right);
		this.m_oPanelRight.AddControl(this.m_oPanelRight_buttonRulers);

		var _vertScrollTop = ScrollWidthMm;
		if (GlobalSkin.RulersButton === false)
		{
			this.m_oPanelRight_buttonRulers.HtmlElement.style.display = "none";
			_vertScrollTop                                            = 0;
		}

		this.m_oPanelRight_buttonNextPage = CreateControl("id_buttonNextPage");
		this.m_oPanelRight_buttonNextPage.Bounds.SetParams(0, 0, 1000, 1000, false, false, false, false, -1, ScrollWidthMm);
		this.m_oPanelRight_buttonNextPage.Anchor = (g_anchor_left | g_anchor_bottom | g_anchor_right);
		this.m_oPanelRight.AddControl(this.m_oPanelRight_buttonNextPage);

		this.m_oPanelRight_buttonPrevPage = CreateControl("id_buttonPrevPage");
		this.m_oPanelRight_buttonPrevPage.Bounds.SetParams(0, 0, 1000, ScrollWidthMm, false, false, false, true, -1, ScrollWidthMm);
		this.m_oPanelRight_buttonPrevPage.Anchor = (g_anchor_left | g_anchor_bottom | g_anchor_right);
		this.m_oPanelRight.AddControl(this.m_oPanelRight_buttonPrevPage);

		var _vertScrollBottom = 2 * ScrollWidthMm;
		if (GlobalSkin.NavigationButtons == false)
		{
			this.m_oPanelRight_buttonNextPage.HtmlElement.style.display = "none";
			this.m_oPanelRight_buttonPrevPage.HtmlElement.style.display = "none";
			_vertScrollBottom                                           = 0;
		}

		this.m_oPanelRight_vertScroll = CreateControl("id_vertical_scroll");
		this.m_oPanelRight_vertScroll.Bounds.SetParams(0, _vertScrollTop, 1000, _vertScrollBottom, false, true, false, true, -1, -1);
		this.m_oPanelRight_vertScroll.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right | g_anchor_bottom);
		this.m_oPanelRight.AddControl(this.m_oPanelRight_vertScroll);
		// --------------------------------------------------------------------------

		// --- left ---
		this.m_oLeftRuler = CreateControlContainer("id_panel_left");
		this.m_oLeftRuler.Bounds.SetParams(0, 0, 1000, 1000, false, false, false, false, 5, -1);
		this.m_oLeftRuler.Anchor = (g_anchor_left | g_anchor_top | g_anchor_bottom);
		this.m_oMainContent.AddControl(this.m_oLeftRuler);

		this.m_oLeftRuler_buttonsTabs = CreateControl("id_buttonTabs");
		this.m_oLeftRuler_buttonsTabs.Bounds.SetParams(0, 0.8, 1000, 1000, false, true, false, false, -1, 5);
		this.m_oLeftRuler_buttonsTabs.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right);
		this.m_oLeftRuler.AddControl(this.m_oLeftRuler_buttonsTabs);

		this.m_oLeftRuler_vertRuler = CreateControl("id_vert_ruler");
		this.m_oLeftRuler_vertRuler.Bounds.SetParams(0, 7, 1000, 1000, false, true, false, false, -1, -1);
		this.m_oLeftRuler_vertRuler.Anchor = (g_anchor_left | g_anchor_right | g_anchor_top | g_anchor_bottom);
		this.m_oLeftRuler.AddControl(this.m_oLeftRuler_vertRuler);
		// ------------

		// --- top ----
		this.m_oTopRuler = CreateControlContainer("id_panel_top");
		this.m_oTopRuler.Bounds.SetParams(5, 0, 1000, 1000, true, false, false, false, -1, 7);
		this.m_oTopRuler.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right);
		this.m_oMainContent.AddControl(this.m_oTopRuler);

		this.m_oTopRuler_horRuler = CreateControl("id_hor_ruler");
		this.m_oTopRuler_horRuler.Bounds.SetParams(0, 0, 1000, 1000, false, false, false, false, -1, -1);
		this.m_oTopRuler_horRuler.Anchor = (g_anchor_left | g_anchor_right | g_anchor_top | g_anchor_bottom);
		this.m_oTopRuler.AddControl(this.m_oTopRuler_horRuler);
		// ------------

		// scroll hor --
		this.m_oScrollHor = CreateControlContainer("id_horscrollpanel");
		this.m_oScrollHor.Bounds.SetParams(0, 0, ScrollWidthMm, 1000, false, false, true, false, -1, ScrollWidthMm);
		this.m_oScrollHor.Anchor = (g_anchor_left | g_anchor_right | g_anchor_bottom);
		this.m_oMainContent.AddControl(this.m_oScrollHor);

		// ----------
		this.m_oMainView = CreateControlContainer("id_main_view");
		var useScrollW = (this.m_oApi.isMobileVersion) ? 0 : ScrollWidthMm;
		this.m_oMainView.Bounds.SetParams(5, 7, useScrollW, useScrollW, true, true, true, true, -1, -1);
		this.m_oMainView.Anchor = (g_anchor_left | g_anchor_right | g_anchor_top | g_anchor_bottom);
		this.m_oMainContent.AddControl(this.m_oMainView);

		// проблема с фокусом fixed-позиционированного элемента внутри (bug 63194)
		this.m_oMainView.HtmlElement.onscroll = function() {
			this.scrollTop = 0;
		};

		this.m_oEditor = CreateControl("id_viewer");
		this.m_oEditor.Bounds.SetParams(0, 0, 1000, 1000, false, false, false, false, -1, -1);
		this.m_oEditor.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right | g_anchor_bottom);
		this.m_oMainView.AddControl(this.m_oEditor);

		this.m_oOverlay = CreateControl("id_viewer_overlay");
		this.m_oOverlay.Bounds.SetParams(0, 0, 1000, 1000, false, false, false, false, -1, -1);
		this.m_oOverlay.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right | g_anchor_bottom);
		this.m_oMainView.AddControl(this.m_oOverlay);

		this.m_oBody.HtmlElement.oncontextmenu = function(e)
		{
			if (AscCommon.AscBrowser.isVivaldiLinux)
			{
				AscCommon.Window_OnMouseUp(e);
			}
			AscCommon.stopEvent(e);
			return false;
		};
		// --------------------------------------------------------------------------

		this.m_oDrawingDocument.TargetHtmlElement = document.getElementById('id_target_cursor');

		this.m_oOverlayApi.m_oControl  = this.m_oOverlay;
		this.m_oOverlayApi.m_oHtmlPage = this;
		this.m_oOverlayApi.Clear();
		this.ShowOverlay();

		this.m_oDrawingDocument.AutoShapesTrack = new AscCommon.CAutoshapeTrack();
		this.m_oDrawingDocument.AutoShapesTrack.init2(this.m_oOverlayApi);

		this.SlideDrawer.m_oWordControl = this;

		this.checkNeedRules();
		this.initEvents();
		this.OnResize(true);

		if (AscCommon.AscBrowser.isIE && !AscCommon.AscBrowser.isIeEdge)
		{
			var ie_hack = [
				this.m_oThumbnailsBack,
				this.m_oThumbnails,
				this.m_oMainContent,
				this.m_oEditor,
				this.m_oOverlay
			];

			for (var elem in ie_hack)
			{
				if (ie_hack[elem] && ie_hack[elem].HtmlElement)
					ie_hack[elem].HtmlElement.style.zIndex = 0;
			}
		}
	};

	this.CheckRetinaDisplay = function()
	{
		if (this.retinaScaling !== AscCommon.AscBrowser.retinaPixelRatio)
		{
			this.retinaScaling = AscCommon.AscBrowser.retinaPixelRatio;
			this.m_oDrawingDocument.ClearCachePages();
			//todo
			this.onButtonTabsDraw();
		}
	};

	this.CheckRetinaElement = function(htmlElem)
	{
		switch (htmlElem.id)
		{
			case "id_viewer":
			case "id_viewer_overlay":
			case "id_hor_ruler":
			case "id_vert_ruler":
			case "id_buttonTabs":
			case "id_thumbnails_background":
			case "id_thumbnails":
				return true;
			default:
				break;
		}
		return false;
	};
	this.CheckRetinaElement2 = function(htmlElem)
	{
		switch (htmlElem.id)
		{
			case "id_viewer":
			case "id_viewer_overlay":
				return true;
			default:
				break;
		}
		return false;
	};

	this.GetMainContentBounds = function()
	{
		return this.m_oMainParent.AbsolutePosition;
	};
	this.GetVertRulerLeft = function()
	{
		return 0;
	};

	// splitter
	this.createSplitterElement = function()
	{
		var Splitter            = document.createElement("div");
		Splitter.id             = "splitter_id";
		Splitter.style.position = "absolute";
		// skin
		//Splitter.style.backgroundColor = "#000000";
		//Splitter.style.backgroundImage = "url(Images/splitter.png)";
		Splitter.style.backgroundImage = "url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAYAAACp8Z5+AAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABh0RVh0U29mdHdhcmUAUGFpbnQuTkVUIHYzLjMxN4N3hgAAAB9JREFUGFdj+P//PwsDAwOQ+m8PooEYwQELwmRwqgAAbXwhnmjs9sgAAAAASUVORK5CYII=)";
		Splitter.style.overflow = 'hidden';
		Splitter.style.zIndex   = 1000;
		Splitter.setAttribute("contentEditable", false);
		return Splitter;
	};

	this.setVerticalSplitter = function(Splitter, dXPos)
	{
		Splitter.style.left   = parseInt(dXPos * g_dKoef_mm_to_pix) + "px";
		Splitter.style.top    = "0px";
		Splitter.style.width  = parseInt(GlobalSkin.SplitterWidthMM * g_dKoef_mm_to_pix) + "px";
		Splitter.style.height = this.Height + "px";
		Splitter.style.backgroundRepeat = "repeat-y";
	};

	this.setHorizontalSplitter = function(Splitter, dYPos)
	{
		Splitter.style.left   = parseInt((this.Splitter1Pos + GlobalSkin.SplitterWidthMM) * g_dKoef_mm_to_pix) + "px";
		Splitter.style.top    = (this.Height - parseInt((dYPos + GlobalSkin.SplitterWidthMM) * g_dKoef_mm_to_pix) + 1) + "px";
		Splitter.style.width  = this.Width - parseInt((this.Splitter1Pos + GlobalSkin.SplitterWidthMM) * g_dKoef_mm_to_pix) + "px";
		Splitter.style.height = parseInt(GlobalSkin.SplitterWidthMM * g_dKoef_mm_to_pix) + "px";
		Splitter.style.backgroundRepeat = "repeat-x";
	};

	this.createSplitterDiv = function(nIdx)
	{
		var Splitter = this.createSplitterElement();
		this.SplitterDiv = Splitter;
		this.SplitterType = nIdx;
		if (nIdx === 1)
		{
			this.setVerticalSplitter(Splitter, this.Splitter1Pos);
		}
		else if(nIdx === 2)
		{
			this.setHorizontalSplitter(Splitter, this.Splitter2Pos);
		}
		else if(nIdx === 3)
		{
			this.setHorizontalSplitter(Splitter, this.Splitter3Pos);
		}
		this.m_oBody.HtmlElement.appendChild(this.SplitterDiv);
	};

	this.hitToSplitter = function()
	{
		let nResult = 0;
		let oWordControl = oThis;

		let x1 = oWordControl.Splitter1Pos * g_dKoef_mm_to_pix;
		let x2 = (oWordControl.Splitter1Pos + GlobalSkin.SplitterWidthMM) * g_dKoef_mm_to_pix;

		let _x = global_mouseEvent.X - oWordControl.X;
		let _y = global_mouseEvent.Y - oWordControl.Y;

		if (_x >= x1 && _x <= x2 && _y >= 0 && _y <= oWordControl.Height && (oThis.IsUseNullThumbnailsSplitter || (oThis.Splitter1Pos != 0)))
		{
			nResult = 1;
		}
		else if (_x >= x2 && _x <= oWordControl.Width)
		{

			var y1 = oWordControl.Height - ((oWordControl.Splitter2Pos + GlobalSkin.SplitterWidthMM) * g_dKoef_mm_to_pix);
			var y2 = oWordControl.Height - (oWordControl.Splitter2Pos * g_dKoef_mm_to_pix);

			if(_y >= y1 && _y <= y2)
			{
				nResult = 2;
			}
		}

		//check event sender to prevent tracking after click on menu elements (bug 60586)
		if(nResult !== 0)
		{
			if(global_mouseEvent.Sender)
			{
				let oThContainer, oMainContainer, oBodyElement;
				let oSender = global_mouseEvent.Sender;
				if(this.m_oThumbnailsContainer)
				{
					oThContainer = this.m_oThumbnailsContainer.HtmlElement;
				}
				if(this.m_oMainParent)
				{
					oMainContainer = this.m_oMainParent.HtmlElement;
				}
				if(this.m_oBody)
				{
					oBodyElement = this.m_oBody.HtmlElement;
				}
				if(!(oThContainer && oThContainer.contains(oSender) ||
					oMainContainer && oMainContainer.contains(oSender) ||
					oBodyElement && oSender.contains(oBodyElement)))
				{
					return 0;
				}
			}

		}
		return nResult;
	};

	this.onBodyMouseDown = function(e)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		if (AscCommon.g_inputContext && AscCommon.g_inputContext.externalChangeFocus())
			return;

		if (oThis.SplitterType !== 0)
			return;

		let _isCatch = false;

		let downClick = global_mouseEvent.ClickCount;
		AscCommon.check_MouseDownEvent(e, true);
		global_mouseEvent.ClickCount = downClick;
		global_mouseEvent.LockMouse();

		let oWordControl = oThis;
		let nSplitter = 0;
		nSplitter = oThis.hitToSplitter();
		if(nSplitter > 0)
		{
			if(nSplitter === 1)
			{
				oWordControl.m_oBody.HtmlElement.style.cursor = "w-resize";
			}
			else
			{
				oWordControl.m_oBody.HtmlElement.style.cursor = "s-resize";
			}
			oWordControl.createSplitterDiv(nSplitter);
			_isCatch = true;
		}
		else
		{
			oWordControl.m_oBody.HtmlElement.style.cursor = "default";
		}
		if (_isCatch)
		{
			if (oWordControl.m_oMainParent && oWordControl.m_oMainParent.HtmlElement)
				oWordControl.m_oMainParent.HtmlElement.style.pointerEvents = "none";
			if (oWordControl.m_oThumbnailsContainer && oWordControl.m_oThumbnailsContainer.HtmlElement)
				oWordControl.m_oThumbnailsContainer.HtmlElement.style.pointerEvents = "none";
			AscCommon.stopEvent(e);
		}
	};

	this.onBodyMouseMove = function(e)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		var _isCatch = false;

		AscCommon.check_MouseMoveEvent(e, true);

		var oWordControl = oThis;
		if (null == oWordControl.SplitterDiv)
		{
			var nSplitter = oThis.hitToSplitter();
			if(nSplitter === 0)
			{
				oWordControl.m_oBody.HtmlElement.style.cursor = "default";
			}
			else if(nSplitter === 1)
			{
				oWordControl.m_oBody.HtmlElement.style.cursor = "w-resize";
			}
			else
			{
				oWordControl.m_oBody.HtmlElement.style.cursor = "s-resize";
			}
		}
		else
		{
			var _x = global_mouseEvent.X - oWordControl.X;
			var _y = global_mouseEvent.Y - oWordControl.Y;

			if (1 == oWordControl.SplitterType)
			{
				var isCanUnShowThumbnails = true;

				var _min = parseInt(oWordControl.Splitter1PosMin * g_dKoef_mm_to_pix);
				var _max = parseInt(oWordControl.Splitter1PosMax * g_dKoef_mm_to_pix);
				if (_x > _max)
					_x = _max;
				else if ((_x < (_min / 2)) && isCanUnShowThumbnails)
					_x = 0;
				else if (_x < _min)
					_x = _min;

				oWordControl.m_oBody.HtmlElement.style.cursor = "w-resize";
				oWordControl.SplitterDiv.style.left           = _x + "px";
			}
			else
			{
				var _max = oWordControl.Height - parseInt(oWordControl.Splitter2PosMin * g_dKoef_mm_to_pix);
				var _min = oWordControl.Height - parseInt(oWordControl.Splitter2PosMax * g_dKoef_mm_to_pix);

				if (_min < (30 * g_dKoef_mm_to_pix))
					_min = 30 * g_dKoef_mm_to_pix;

				var _c = parseInt(oWordControl.Splitter2PosMin * g_dKoef_mm_to_pix);
				if (_y > (_max + (_c / 2)))
					_y = oWordControl.Height;
				else if (_y > _max)
					_y = _max;
				else if (_y < _min)
					_y = _min;

				oWordControl.m_oBody.HtmlElement.style.cursor = "s-resize";
				oWordControl.SplitterDiv.style.top            = (_y - parseInt(GlobalSkin.SplitterWidthMM * g_dKoef_mm_to_pix)) + "px";
			}

			_isCatch = true;
		}

		if (_isCatch)
			AscCommon.stopEvent(e);
	};

	this.OnResizeSplitter = function(isNoNeedResize)
	{
		this.OldSplitter1Pos = this.Splitter1Pos;
		this.OldSplitter2Pos = this.Splitter2Pos;
		this.OldSplitter3Pos = this.Splitter3Pos;

		this.m_oThumbnailsContainer.Bounds.R    = this.Splitter1Pos;
		this.m_oThumbnailsContainer.Bounds.AbsW = this.Splitter1Pos;
		this.m_oThumbnailsSplit.Bounds.L 		= this.Splitter1Pos;

		this.Splitter3Pos = 0;

		this.Splitter2Pos = 0;

		if (this.IsUseNullThumbnailsSplitter || (0 != this.Splitter1Pos))
		{
			this.m_oMainParent.Bounds.L = this.Splitter1Pos + GlobalSkin.SplitterWidthMM;
			this.m_oMainContent.Bounds.B = 1000;
			this.m_oMainContent.Bounds.isAbsB = false;

			this.m_oThumbnailsContainer.HtmlElement.style.display = "block";
			this.m_oThumbnailsSplit.HtmlElement.style.display = "block";
			this.m_oMainParent.HtmlElement.style.borderLeft = ("1px solid " + GlobalSkin.BorderSplitterColor);
		}
		else
		{
			this.m_oMainParent.Bounds.L = 0;
			this.m_oMainContent.Bounds.B = 1000;
			this.m_oMainContent.Bounds.isAbsB = false;

			this.m_oThumbnailsContainer.HtmlElement.style.display = "none";
			this.m_oThumbnailsSplit.HtmlElement.style.display = "none";
			this.m_oMainParent.HtmlElement.style.borderLeft = "none";
		}

		if (true !== isNoNeedResize)
			this.OnResize2(true);
	};

	this.onBodyMouseUp = function(e)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		var _isCatch = false;

		AscCommon.check_MouseUpEvent(e, true);

		var oWordControl = oThis;
		oWordControl.UnlockCursorTypeOnMouseUp();

		if (oWordControl.m_oMainParent && oWordControl.m_oMainParent.HtmlElement)
			oWordControl.m_oMainParent.HtmlElement.style.pointerEvents = "";
		if (oWordControl.m_oThumbnailsContainer && oWordControl.m_oThumbnailsContainer.HtmlElement)
			oWordControl.m_oThumbnailsContainer.HtmlElement.style.pointerEvents = "";

		if (null != oWordControl.SplitterDiv)
		{
			var _x = parseInt(oWordControl.SplitterDiv.style.left);
			var _y = parseInt(oWordControl.SplitterDiv.style.top);

			if (oWordControl.SplitterType == 1)
			{
				var _posX = _x * g_dKoef_pix_to_mm;
				if (Math.abs(oWordControl.Splitter1Pos - _posX) > 1)
				{
					oWordControl.Splitter1Pos = _posX;
					oWordControl.Splitter1PosSetUp = oWordControl.Splitter1Pos;
					oWordControl.OnResizeSplitter();
					oWordControl.m_oApi.syncOnThumbnailsShow();
				}
			}
			else if(oWordControl.SplitterType == 2)
			{
				var _posY = (oWordControl.Height - _y) * g_dKoef_pix_to_mm - GlobalSkin.SplitterWidthMM;
				if (Math.abs(oWordControl.Splitter2Pos - _posY) > 1)
				{
					oWordControl.SetSplitter2Pos(_posY);
					oWordControl.OnResizeSplitter();
				}
			}
			else if(oWordControl.SplitterType == 3)
			{
				var _posY = (oWordControl.Height - _y) * g_dKoef_pix_to_mm - GlobalSkin.SplitterWidthMM;
				if (Math.abs(oWordControl.Splitter3Pos - _posY) > 1)
				{
					oWordControl.SetSplitter3Pos(_posY);
					oWordControl.OnResizeSplitter();
				}
			}

			oWordControl.m_oBody.HtmlElement.removeChild(oWordControl.SplitterDiv);
			oWordControl.SplitterDiv  = null;
			oWordControl.SplitterType = 0;

			_isCatch = true;
		}

		if (_isCatch)
			AscCommon.stopEvent(e);
	};

	this.SetSplitter3Pos = function(dPos)
	{
		var dDiff = this.Splitter2Pos - this.Splitter3Pos;
		this.Splitter3Pos = dPos;
		this.Splitter2Pos = Math.max(this.Splitter2Pos, this.Splitter3Pos + dDiff);
	};
	this.SetSplitter2Pos = function(dPos)
	{
		this.Splitter2Pos = Math.max(this.Splitter3Pos + HIDDEN_PANE_HEIGHT, dPos);
	};

	// events
	this.initEvents = function()
	{
		this.arrayEventHandlers[0] = new AscCommon.button_eventHandlers("", "0px 0px", "0px -16px", "0px -32px", this.m_oPanelRight_buttonRulers, this.onButtonRulersClick);
		this.arrayEventHandlers[1] = new AscCommon.button_eventHandlers("", "0px 0px", "0px -16px", "0px -32px", this.m_oPanelRight_buttonPrevPage, this.onPrevPage);
		this.arrayEventHandlers[2] = new AscCommon.button_eventHandlers("", "0px -48px", "0px -64px", "0px -80px", this.m_oPanelRight_buttonNextPage, this.onNextPage);

		this.m_oLeftRuler_buttonsTabs.HtmlElement.onclick = this.onButtonTabsClick;

		AscCommon.addMouseEvent(this.m_oEditor.HtmlElement, "down", this.onMouseDown);
		AscCommon.addMouseEvent(this.m_oEditor.HtmlElement, "move", this.onMouseMove);
		AscCommon.addMouseEvent(this.m_oEditor.HtmlElement, "up", this.onMouseUp);

		AscCommon.addMouseEvent(this.m_oOverlay.HtmlElement, "down", this.onMouseDown);
		AscCommon.addMouseEvent(this.m_oOverlay.HtmlElement, "move", this.onMouseMove);
		AscCommon.addMouseEvent(this.m_oOverlay.HtmlElement, "up", this.onMouseUp);

		document.getElementById('id_target_cursor').style.pointerEvents = "none";

		this.m_oMainContent.HtmlElement.onmousewheel = this.onMouseWhell;
		if (this.m_oMainContent.HtmlElement.addEventListener)
			this.m_oMainContent.HtmlElement.addEventListener("DOMMouseScroll", this.onMouseWhell, false);

		this.m_oBody.HtmlElement.onmousewheel = function(e)
		{
			e.preventDefault();
			return false;
		};

		AscCommon.addMouseEvent(this.m_oTopRuler_horRuler.HtmlElement, "down", this.horRulerMouseDown);
		AscCommon.addMouseEvent(this.m_oTopRuler_horRuler.HtmlElement, "move", this.horRulerMouseMove);
		AscCommon.addMouseEvent(this.m_oTopRuler_horRuler.HtmlElement, "up", this.horRulerMouseUp);

		AscCommon.addMouseEvent(this.m_oLeftRuler_vertRuler.HtmlElement, "down", this.verRulerMouseDown);
		AscCommon.addMouseEvent(this.m_oLeftRuler_vertRuler.HtmlElement, "move", this.verRulerMouseMove);
		AscCommon.addMouseEvent(this.m_oLeftRuler_vertRuler.HtmlElement, "up", this.verRulerMouseUp);

		if (!this.m_oApi.isMobileVersion)
		{
			AscCommon.addMouseEvent(this.m_oMainParent.HtmlElement, "down", this.onBodyMouseDown);
			AscCommon.addMouseEvent(this.m_oMainParent.HtmlElement, "move", this.onBodyMouseMove);
			AscCommon.addMouseEvent(this.m_oMainParent.HtmlElement, "up", this.onBodyMouseUp);

			AscCommon.addMouseEvent(this.m_oBody.HtmlElement, "down", this.onBodyMouseDown);
			AscCommon.addMouseEvent(this.m_oBody.HtmlElement, "move", this.onBodyMouseMove);
			AscCommon.addMouseEvent(this.m_oBody.HtmlElement, "up", this.onBodyMouseUp);
		}

		// в мобильной версии - при транзишне - не обновляется позиция/размер
		if (this.m_oApi.isMobileVersion)
		{
			var _t = this;
			document.addEventListener && document.addEventListener("transitionend", function() { _t.OnResize(false);  }, false);
			document.addEventListener && document.addEventListener("transitioncancel", function() { _t.OnResize(false); }, false);
		}

		this.Thumbnails.initEvents();
	};

	this.initEventsMobile = function()
	{
		if (this.m_oApi.isUseOldMobileVersion())
		{
			this.MobileTouchManager = new AscCommon.CMobileTouchManager( { eventsElement : "slides_mobile_element" } );
			this.MobileTouchManager.Init(this.m_oApi);

			this.MobileTouchManagerThumbnails = new AscCommon.CMobileTouchManagerThumbnails( { eventsElement : "slides_mobile_element" } );
			this.MobileTouchManagerThumbnails.Init(this.m_oApi);

			this.m_oThumbnailsContainer.HtmlElement.style.zIndex = "11";

			this.TextBoxBackground = CreateControl(AscCommon.g_inputContext.HtmlArea.id);
			this.TextBoxBackground.HtmlElement.parentNode.parentNode.style.zIndex = 10;

			this.MobileTouchManager.initEvents(AscCommon.g_inputContext.HtmlArea.id);
			this.MobileTouchManagerThumbnails.initEvents(this.m_oThumbnails.HtmlElement.id);

			if (AscCommon.AscBrowser.isAndroid)
			{
				this.TextBoxBackground.HtmlElement["oncontextmenu"] = function(e)
				{
					if (e.preventDefault)
						e.preventDefault();

					e.returnValue = false;
					return false;
				};

				this.TextBoxBackground.HtmlElement["onselectstart"] = function(e)
				{
					oThis.m_oLogicDocument.SelectAll();

					if (e.preventDefault)
						e.preventDefault();

					e.returnValue = false;
					return false;
				};
			}
		}
		else
		{
			this.MobileTouchManager = new AscCommon.CMobileTouchManager( { eventsElement : "slides_mobile_element", desktopMode : true } );
			this.MobileTouchManager.Init(this.m_oApi);

			this.MobileTouchManagerThumbnails = new AscCommon.CMobileTouchManagerThumbnails( { eventsElement : "slides_mobile_element", desktopMode : true } );
			this.MobileTouchManagerThumbnails.Init(this.m_oApi);
		}
	};

	// buttons
	this.onButtonTabsClick = function()
	{
		var oWordControl = oThis;
		if (oWordControl.m_nTabsType == tab_Left)
		{
			oWordControl.m_nTabsType = tab_Center;
			oWordControl.onButtonTabsDraw();
		}
		else if (oWordControl.m_nTabsType == tab_Center)
		{
			oWordControl.m_nTabsType = tab_Right;
			oWordControl.onButtonTabsDraw();
		}
		else
		{
			oWordControl.m_nTabsType = tab_Left;
			oWordControl
				.onButtonTabsDraw();
		}
	};

	this.onButtonTabsDraw = function()
	{
		var _ctx = this.m_oLeftRuler_buttonsTabs.HtmlElement.getContext('2d');
		_ctx.setTransform(1, 0, 0, 1, 0, 0);

		var dPR = AscCommon.AscBrowser.retinaPixelRatio;
		var _width  = Math.round(19 * dPR);
		var _height = Math.round(19 * dPR);

		_ctx.clearRect(0, 0, _width, _height);

		_ctx.lineWidth   = Math.round(dPR);
		_ctx.strokeStyle = GlobalSkin.RulerOutline;
		var rectSize = Math.round(14 * dPR);
		var lineWidth = _ctx.lineWidth;

		_ctx.strokeRect(2.5 * lineWidth, 3.5 * lineWidth, Math.round(14 * dPR), Math.round(14 * dPR));
		_ctx.beginPath();

		_ctx.strokeStyle = GlobalSkin.RulerTabsColor;

		_ctx.lineWidth = (dPR - Math.floor(dPR) === 0.5) ? 2 * Math.round(dPR) - 1 : 2 * Math.round(dPR);

		var tab_width = Math.round(5 * dPR);
		var offset = _ctx.lineWidth % 2 === 1 ? 0.5 : 0;

		var dx = Math.round((rectSize - 2 * Math.round(dPR) - tab_width) / 7 * 4);
		var dy = Math.round((rectSize - 2 * Math.round(dPR) - tab_width) / 7 * 4);
		var x = 4 * Math.round(dPR) + dx;
		var y = 4 * Math.round(dPR) + dy;

		if (this.m_nTabsType == tab_Left)
		{
			_ctx.moveTo(x + offset, y);
			_ctx.lineTo(x + offset, y + tab_width + offset);
			_ctx.lineTo(x + tab_width, y + tab_width + offset);
		}
		else if (this.m_nTabsType == tab_Center)
		{
			tab_width = Math.round(8 * dPR);
			tab_width = (tab_width % 2 === 1) ? tab_width - 1 : tab_width;
			var dx = Math.round((rectSize - Math.round(dPR) - tab_width) / 2);
			var x = 3 * Math.round(dPR) + dx;
			var vert_tab_width = Math.round(5 * dPR);
			_ctx.moveTo(x , y + vert_tab_width + offset);
			_ctx.lineTo(x + tab_width, y + vert_tab_width + offset);
			_ctx.moveTo(x - offset + tab_width / 2, y);
			_ctx.lineTo(x - offset + tab_width / 2, y + vert_tab_width);
		}
		else
		{
			var x = 3 * Math.round(dPR) + dx;
			_ctx.moveTo(x, tab_width + y + offset);
			_ctx.lineTo(x + tab_width + offset,  tab_width + y + offset);
			_ctx.lineTo(x + tab_width + offset, y);
		}

		_ctx.stroke();
		_ctx.beginPath();
	};

	this.onPrevPage = function()
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		let oWordControl = oThis;
		let nCurrentSlide = oWordControl.m_oDrawingDocument.SlideCurrent;
		if (0 < nCurrentSlide)
		{
			oWordControl.GoToPage(nCurrentSlide - 1);
		}
		else
		{
			oWordControl.GoToPage(0);
		}
	};
	this.onNextPage = function()
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		let oWordControl = oThis;
		let nCurrentSlide = oWordControl.m_oDrawingDocument.SlideCurrent;
		let SlidesCount = this.GetSlidesCount();
		if ((SlidesCount - 1) > nCurrentSlide)
		{
			oWordControl.GoToPage(nCurrentSlide + 1);
		}
		else if (SlidesCount > 0)
		{
			oWordControl.GoToPage(SlidesCount - 1);
		}
	};

	// rulers
	this.onButtonRulersClick       = function()
	{
		if (false === oThis.m_oApi.bInit_word_control || true === oThis.m_oApi.isViewMode)
			return;

		oThis.m_bIsRuler = !oThis.m_bIsRuler;
		oThis.checkNeedRules();
		oThis.OnResize(true);
	};

	this.HideRulers = function()
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		if (oThis.m_bIsRuler === false)
			return;

		oThis.m_bIsRuler = !oThis.m_bIsRuler;
		oThis.checkNeedRules();
		oThis.OnResize(true);
	};

	this.DisableRulerMarkers = function()
	{
		if (!this.IsEnabledRulerMarkers)
			return;

		this.IsEnabledRulerMarkers                 = false;
		this.m_oHorRuler.RepaintChecker.BlitAttack = true;
		this.m_oHorRuler.IsCanMoveAnyMarkers       = false;
		this.m_oHorRuler.IsDrawAnyMarkers          = false;
		this.m_oHorRuler.m_dMarginLeft             = 0;
		this.m_oHorRuler.m_dMarginRight            = this.m_oLogicDocument.GetWidthMM();

		this.m_oVerRuler.m_dMarginTop              = 0;
		this.m_oVerRuler.m_dMarginBottom           = this.m_oLogicDocument.GetHeightMM();
		this.m_oVerRuler.RepaintChecker.BlitAttack = true;

		if (this.m_bIsRuler)
		{
			this.UpdateHorRuler();
			this.UpdateVerRuler();
		}
	};

	this.EnableRulerMarkers = function()
	{
		if (this.IsEnabledRulerMarkers)
			return;

		this.IsEnabledRulerMarkers                 = true;
		this.m_oHorRuler.RepaintChecker.BlitAttack = true;
		this.m_oHorRuler.IsCanMoveAnyMarkers       = true;
		this.m_oHorRuler.IsDrawAnyMarkers          = true;

		if (this.m_bIsRuler)
		{
			this.UpdateHorRuler();
			this.UpdateVerRuler();
		}
	};

	this.horRulerMouseDown = function(e)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		if (e.preventDefault)
			e.preventDefault();
		else
			e.returnValue = false;

		var oWordControl = oThis;

		if (-1 != oWordControl.m_oDrawingDocument.SlideCurrent)
			oWordControl.m_oHorRuler.OnMouseDown(oWordControl.m_oDrawingDocument.SlideCurrectRect.left, 0, e);
	};
	this.horRulerMouseUp   = function(e)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		if (e.preventDefault)
			e.preventDefault();
		else
			e.returnValue = false;

		var oWordControl = oThis;

		if (-1 != oWordControl.m_oDrawingDocument.SlideCurrent)
			oWordControl.m_oHorRuler.OnMouseUp(oWordControl.m_oDrawingDocument.SlideCurrectRect.left, 0, e);
	};
	this.horRulerMouseMove = function(e)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		if (e.preventDefault)
			e.preventDefault();
		else
			e.returnValue = false;

		var oWordControl = oThis;

		if (-1 != oWordControl.m_oDrawingDocument.SlideCurrent)
			oWordControl.m_oHorRuler.OnMouseMove(oWordControl.m_oDrawingDocument.SlideCurrectRect.left, 0, e);
	};

	this.verRulerMouseDown = function(e)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		if (e.preventDefault)
			e.preventDefault();
		else
			e.returnValue = false;

		var oWordControl = oThis;

		if (-1 != oWordControl.m_oDrawingDocument.SlideCurrent)
			oWordControl.m_oVerRuler.OnMouseDown(0, oWordControl.m_oDrawingDocument.SlideCurrectRect.top, e);
	};
	this.verRulerMouseUp   = function(e)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		if (e.preventDefault)
			e.preventDefault();
		else
			e.returnValue = false;

		var oWordControl = oThis;

		if (-1 != oWordControl.m_oDrawingDocument.SlideCurrent)
			oWordControl.m_oVerRuler.OnMouseUp(0, oWordControl.m_oDrawingDocument.SlideCurrectRect.top, e);
	};
	this.verRulerMouseMove = function(e)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		if (e.preventDefault)
			e.preventDefault();
		else
			e.returnValue = false;

		var oWordControl = oThis;

		if (-1 != oWordControl.m_oDrawingDocument.SlideCurrent)
			oWordControl.m_oVerRuler.OnMouseMove(0, oWordControl.m_oDrawingDocument.SlideCurrectRect.top, e);
	};

	this.UpdateHorRuler = function(isattack)
	{
		if (!this.m_bIsRuler)
			return;

		if (!isattack && this.m_oDrawingDocument.SlideCurrent == -1)
			return;

		var drawRect = this.m_oDrawingDocument.SlideCurrectRect;
		var _left    = drawRect.left;
		this.m_oHorRuler.BlitToMain(_left, 0, this.m_oTopRuler_horRuler.HtmlElement);
	};
	this.UpdateVerRuler = function(isattack)
	{
		if (!this.m_bIsRuler)
			return;

		if (!isattack && this.m_oDrawingDocument.SlideCurrent == -1)
			return;

		var drawRect = this.m_oDrawingDocument.SlideCurrectRect;
		var _top     = drawRect.top;
		this.m_oVerRuler.BlitToMain(0, _top, this.m_oLeftRuler_vertRuler.HtmlElement);
	};

	this.UpdateHorRulerBack = function(isattack)
	{
		var drDoc = this.m_oDrawingDocument;
		if (0 <= drDoc.SlideCurrent && drDoc.SlideCurrent < drDoc.GetSlidesCount())
		{
			this.CreateBackgroundHorRuler(undefined, isattack);
		}
		this.UpdateHorRuler(isattack);
	};
	this.UpdateVerRulerBack = function(isattack)
	{
		var drDoc = this.m_oDrawingDocument;
		if (0 <= drDoc.SlideCurrent && drDoc.SlideCurrent < drDoc.GetSlidesCount())
		{
			this.CreateBackgroundVerRuler(undefined, isattack);
		}
		this.UpdateVerRuler(isattack);
	};

	this.CreateBackgroundHorRuler = function(margins, isattack)
	{
		var cachedPage       = {};
		cachedPage.width_mm  = this.m_oLogicDocument.GetWidthMM();
		cachedPage.height_mm = this.m_oLogicDocument.GetHeightMM();

		if (margins !== undefined)
		{
			cachedPage.margin_left   = margins.L;
			cachedPage.margin_top    = margins.T;
			cachedPage.margin_right  = margins.R;
			cachedPage.margin_bottom = margins.B;
		}
		else
		{
			cachedPage.margin_left   = 0;
			cachedPage.margin_top    = 0;
			cachedPage.margin_right  = this.m_oLogicDocument.GetWidthMM();
			cachedPage.margin_bottom = this.m_oLogicDocument.GetHeightMM();
		}

		this.m_oHorRuler.CreateBackground(cachedPage, isattack);
	};
	this.CreateBackgroundVerRuler = function(margins, isattack)
	{
		var cachedPage       = {};
		cachedPage.width_mm  = this.m_oLogicDocument.GetWidthMM();
		cachedPage.height_mm = this.m_oLogicDocument.GetHeightMM();

		if (margins !== undefined)
		{
			cachedPage.margin_left   = margins.L;
			cachedPage.margin_top    = margins.T;
			cachedPage.margin_right  = margins.R;
			cachedPage.margin_bottom = margins.B;
		}
		else
		{
			cachedPage.margin_left   = 0;
			cachedPage.margin_top    = 0;
			cachedPage.margin_right  = this.m_oLogicDocument.GetWidthMM();
			cachedPage.margin_bottom = this.m_oLogicDocument.GetHeightMM();
		}

		this.m_oVerRuler.CreateBackground(cachedPage, isattack);
	};

	this.checkNeedRules = function()
	{
		if (this.m_bIsRuler)
		{
			this.m_oLeftRuler.HtmlElement.style.display = 'block';
			this.m_oTopRuler.HtmlElement.style.display  = 'block';

			this.m_oMainView.Bounds.L = 5;
			this.m_oMainView.Bounds.T = 7;
		}
		else
		{
			this.m_oLeftRuler.HtmlElement.style.display = 'none';
			this.m_oTopRuler.HtmlElement.style.display  = 'none';

			this.m_oMainView.Bounds.L = 0;
			this.m_oMainView.Bounds.T = 0;
		}
	};

	this.FullRulersUpdate = function()
	{
		this.m_oHorRuler.RepaintChecker.BlitAttack = true;
		this.m_oVerRuler.RepaintChecker.BlitAttack = true;

		this.m_bIsUpdateHorRuler = true;
		this.m_bIsUpdateVerRuler = true;

		if (this.m_bIsRuler)
		{
			this.UpdateHorRulerBack(true);
			this.UpdateVerRulerBack(true);
		}
	};

	// zoom
	this.zoom_FitToWidth_value = function()
	{
		var _value = 100;
		if (!this.m_oLogicDocument)
			return _value;

		var w = this.m_oEditor.HtmlElement.width;
		w /= AscCommon.AscBrowser.retinaPixelRatio;

		var Zoom       = 100;
		var _pageWidth = this.m_oLogicDocument.GetWidthMM() * g_dKoef_mm_to_pix;
		if (0 != _pageWidth)
		{
			Zoom = 100 * (w - 2 * this.SlideDrawer.CONST_BORDER) / _pageWidth;

			if (Zoom < 5)
				Zoom = 5;
		}
		_value = Zoom >> 0;
		return _value;
	};
	this.zoom_FitToPage_value  = function(_canvas_height)
	{
		var _value = 100;
		if (!this.m_oLogicDocument)
			return _value;

		var w = this.m_oEditor.HtmlElement.width;
		var h = (undefined == _canvas_height) ? this.m_oEditor.HtmlElement.height : _canvas_height;

		w /= AscCommon.AscBrowser.retinaPixelRatio;
		h /= AscCommon.AscBrowser.retinaPixelRatio;

		var _pageWidth  = this.m_oLogicDocument.GetWidthMM() * g_dKoef_mm_to_pix;
		var _pageHeight = this.m_oLogicDocument.GetHeightMM() * g_dKoef_mm_to_pix;

		var _hor_Zoom = 100;
		if (0 != _pageWidth)
			_hor_Zoom = (100 * (w - 2 * this.SlideDrawer.CONST_BORDER)) / _pageWidth;
		var _ver_Zoom = 100;
		if (0 != _pageHeight)
			_ver_Zoom = (100 * (h - 2 * this.SlideDrawer.CONST_BORDER)) / _pageHeight;

		_value = (Math.min(_hor_Zoom, _ver_Zoom) - 0.5) >> 0;

		if (_value < 5)
			_value = 5;
		return _value;
	};

	this.zoom_FitToWidth = function()
	{
		this.m_nZoomType = 1;
		if (!this.m_oLogicDocument)
			return;

		var _new_value = this.zoom_FitToWidth_value();
		if (_new_value != this.m_nZoomValue)
		{
			this.m_nZoomValue = _new_value;
			this.zoom_Fire(1);
			return true;
		}
		else
		{
			this.m_oApi.sync_zoomChangeCallback(this.m_nZoomValue, 1);
		}
		return false;
	};
	this.zoom_FitToPage  = function()
	{
		this.m_nZoomType = 2;
		if (!this.m_oLogicDocument)
			return;

		var _new_value = this.zoom_FitToPage_value();

		if (_new_value != this.m_nZoomValue)
		{
			this.m_nZoomValue = _new_value;
			this.zoom_Fire(2);
			return true;
		}
		else
		{
			this.m_oApi.sync_zoomChangeCallback(this.m_nZoomValue, 2);
		}
		return false;
	};

	this.zoom_Fire = function(type)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		this.m_nZoomType = type;

		// нужно проверить режим и сбросить кеш грамотно (ie version)
		AscCommon.g_fontManager.ClearRasterMemory();

		var oWordControl = oThis;

		oWordControl.m_bIsRePaintOnScroll = false;
		var dPosition                     = 0;
		if (oWordControl.m_dScrollY_max != 0)
		{
			dPosition = oWordControl.m_dScrollY / oWordControl.m_dScrollY_max;
		}
		oWordControl.CheckZoom();
		oWordControl.CalculateDocumentSize();
		var lCurPage = oWordControl.m_oLogicDocument.getCurrentPage();
		//var lCurPage = oWordControl.m_oDrawingDocument.SlideCurrent;

		this.GoToPage(lCurPage, true);
		this.ZoomFreePageNum = lCurPage;

		if (-1 != lCurPage)
		{
			this.CreateBackgroundHorRuler();
			oWordControl.m_bIsUpdateHorRuler = true;
			this.CreateBackgroundVerRuler();
			oWordControl.m_bIsUpdateVerRuler = true;
		}
		var lPosition = parseInt(dPosition * oWordControl.m_oScrollVerApi.getMaxScrolledY());
		oWordControl.m_oScrollVerApi.scrollToY(lPosition);

		this.ZoomFreePageNum = -1;

		oWordControl.m_oApi.sync_zoomChangeCallback(this.m_nZoomValue, type);
		oWordControl.m_bIsUpdateTargetNoAttack = true;
		oWordControl.m_bIsRePaintOnScroll      = true;

		oWordControl.OnScroll();

		if (this.MobileTouchManager)
			this.MobileTouchManager.Resize_After();
	};

	this.zoom_Out = function()
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		var _zooms = oThis.zoom_values;
		var _count = _zooms.length;

		var _Zoom = _zooms[0];
		for (var i = (_count - 1); i >= 0; i--)
		{
			if (this.m_nZoomValue > _zooms[i])
			{
				_Zoom = _zooms[i];
				break;
			}
		}

		if (oThis.m_nZoomValue <= _Zoom)
			return;

		oThis.m_nZoomValue = _Zoom;
		oThis.zoom_Fire(0);
	};

	this.zoom_In = function()
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		var _zooms = oThis.zoom_values;
		var _count = _zooms.length;

		var _Zoom = _zooms[_count - 1];
		for (var i = 0; i < _count; i++)
		{
			if (this.m_nZoomValue < _zooms[i])
			{
				_Zoom = _zooms[i];
				break;
			}
		}

		if (oThis.m_nZoomValue >= _Zoom)
			return;

		oThis.m_nZoomValue = _Zoom;
		oThis.zoom_Fire(0);
	};

	this.CheckZoom = function()
	{
		if (!this.NoneRepaintPages) {
			// //todo
			// this.m_oLogicDocument.draw(this.m_nZoomValue, undefined, this.m_oLogicDocument.pageIndex);
			this.m_oDrawingDocument.ClearCachePages();
		}
	};

	// overlay
	this.ShowOverlay        = function()
	{
		this.m_oOverlayApi.Show();
	};
	this.UnShowOverlay      = function()
	{
		this.m_oOverlayApi.UnShow();
	};
	this.CheckUnShowOverlay = function()
	{
		var drDoc = this.m_oDrawingDocument;

		/*
		 if (!drDoc.m_bIsSearching && !drDoc.m_bIsSelection)
		 {
		 this.UnShowOverlay();
		 return false;
		 }
		 */

		return true;
	};
	this.CheckShowOverlay   = function()
	{
		var drDoc = this.m_oDrawingDocument;
		if (drDoc.m_bIsSearching || drDoc.m_bIsSelection)
			this.ShowOverlay();
	};

	// paint
	this.GetDrawingPageInfo = function(nPageIndex)
	{
		return {
			drawingPage : this.m_oDrawingDocument.SlideCurrectRect,
			width_mm    : this.m_oLogicDocument.GetWidthMM(),
			height_mm   : this.m_oLogicDocument.GetHeightMM()
		};
	};

	this.OnRePaintAttack = function()
	{
		this.m_bIsFullRepaint = true;
		this.OnScroll();
	};

	this.StartMainTimer = function()
	{
		this.paintMessageLoop.Start(this.onTimerScroll.bind(this));
	};

	this.onTimerScroll = function(isThUpdateSync)
	{
		var oWordControl                = oThis;

		if (oWordControl.m_oApi.isLongAction())
			return;

		var isRepaint = oWordControl.m_bIsScroll;
		if (oWordControl.m_bIsScroll)
		{
			oWordControl.m_bIsScroll = false;
			oWordControl.OnPaint();

			if (isThUpdateSync !== undefined)
			{
				oWordControl.Thumbnails.onCheckUpdate();
			}
		}
		else
		{
			oWordControl.Thumbnails.onCheckUpdate();
		}

		if (null != oWordControl.m_oLogicDocument)
		{
			oWordControl.m_oDrawingDocument.UpdateTargetFromPaint = true;
			oWordControl.m_oLogicDocument.CheckTargetUpdate();
			oWordControl.m_oDrawingDocument.CheckTargetShow();
			oWordControl.m_oDrawingDocument.UpdateTargetFromPaint = false;

			oWordControl.CheckFontCache();

			if (oWordControl.m_bIsUpdateTargetNoAttack)
			{
				oWordControl.m_oDrawingDocument.UpdateTargetNoAttack();
				oWordControl.m_bIsUpdateTargetNoAttack = false;
			}
		}

		oWordControl.m_oDrawingDocument.Collaborative_TargetsUpdate(isRepaint);
	};

	this.onTimerScroll_sync = function(isThUpdateSync)
	{
		var oWordControl                = oThis;
		var isRepaint                   = oWordControl.m_bIsScroll;
		if (oWordControl.m_bIsScroll)
		{
			oWordControl.m_bIsScroll = false;
			oWordControl.OnPaint();

			if (isThUpdateSync !== undefined)
			{
				oWordControl.Thumbnails.onCheckUpdate();
			}
			if (null != oWordControl.m_oLogicDocument && oWordControl.m_oApi.bInit_word_control)
				;//oWordControl.m_oLogicDocument.Viewer_OnChangePosition();//todo
		}
		else
		{
			oWordControl.Thumbnails.onCheckUpdate();
		}
		if (null != oWordControl.m_oLogicDocument)
		{
			oWordControl.m_oDrawingDocument.UpdateTargetFromPaint = true;
			// oWordControl.m_oLogicDocument.CheckTargetUpdate();//todo
			// oWordControl.m_oDrawingDocument.CheckTargetShow();
			oWordControl.m_oDrawingDocument.UpdateTargetFromPaint = false;

			oWordControl.CheckFontCache();
		}
		oWordControl.m_oDrawingDocument.Collaborative_TargetsUpdate(isRepaint);
	};

	this.OnPaint = function()
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		var canvas = this.m_oEditor.HtmlElement;
		if (null == canvas)
			return;



		var context = canvas.getContext("2d");
		var _width  = canvas.width;
		var _height = canvas.height;

		context.fillStyle = GlobalSkin.BackgroundColor;
		context.fillRect(0, 0, _width, _height);
		//context.clearRect(0, 0, _width, _height);

		/*
		 if (this.SlideDrawer.IsRecalculateSlide == true)
		 {
		 this.SlideDrawer.CheckSlide(this.m_oDrawingDocument.SlideCurrent);
		 this.SlideDrawer.IsRecalculateSlide = false;
		 }
		 */

		//todo
		// this.m_oLogicDocument.draw(this.m_nZoomValue, undefined, this.m_oLogicDocument.pageIndex);
		this.SlideDrawer.DrawSlide(context, this.m_dScrollX, this.m_dScrollX_max,
			this.m_dScrollY - this.SlideScrollMIN, this.m_dScrollY_max, this.m_oDrawingDocument.SlideCurrent);

		this.OnUpdateOverlay();

		if (this.m_bIsUpdateHorRuler)
		{
			this.UpdateHorRuler();
			this.m_bIsUpdateHorRuler = false;
		}
		if (this.m_bIsUpdateVerRuler)
		{
			this.UpdateVerRuler();
			this.m_bIsUpdateVerRuler = false;
		}
		if (this.m_bIsUpdateTargetNoAttack)
		{
			//todo
			//this.m_oDrawingDocument.UpdateTargetNoAttack();
			this.m_bIsUpdateTargetNoAttack = false;
		}
		this.m_oApi.clearEyedropperImgData();
	};

	this.OnScroll = function()
	{
		this.OnCalculatePagesPlace();
		this.m_bIsScroll = true;
	};

	// position
	this.CalculateDocumentSizeInternal = function(_canvas_height, _zoom_value, _check_bounds2)
	{
		var size = {
			m_dDocumentWidth		: 0,
			m_dDocumentHeight		: 0,
			m_dDocumentPageWidth	: 0,
			m_dDocumentPageHeight	: 0,
			SlideScrollMIN			: 0,
			SlideScrollMAX			: 0
		};

		var _zoom = (undefined == _zoom_value) ? this.m_nZoomValue : _zoom_value;
		var dKoef = (_zoom * g_dKoef_mm_to_pix / 100);

		var _bounds_slide = this.SlideBoundsOnCalculateSize;
		if (undefined == _check_bounds2)
		{
			this.SlideBoundsOnCalculateSize.fromBounds(this.SlideDrawer.BoundsChecker.Bounds);
		}
		else
		{
			_bounds_slide = new AscFormat.CBoundsController();
			this.SlideDrawer.CheckSlideSize(_zoom, this.m_oDrawingDocument.SlideCurrent);
			_bounds_slide.fromBounds(this.SlideDrawer.BoundsChecker2.Bounds);
		}

		var _srcW = this.m_oEditor.HtmlElement.width;
		var _srcH = (undefined !== _canvas_height) ? _canvas_height : this.m_oEditor.HtmlElement.height;
		if (AscCommon.AscBrowser.isCustomScaling())
		{
			_srcW = (_srcW / AscCommon.AscBrowser.retinaPixelRatio) >> 0;
			_srcH = (_srcH / AscCommon.AscBrowser.retinaPixelRatio) >> 0;

			_bounds_slide = {
				min_x : (_bounds_slide.min_x / AscCommon.AscBrowser.retinaPixelRatio) >> 0,
				min_y : (_bounds_slide.min_y / AscCommon.AscBrowser.retinaPixelRatio) >> 0,
				max_x : (_bounds_slide.max_x / AscCommon.AscBrowser.retinaPixelRatio) >> 0,
				max_y : (_bounds_slide.max_y / AscCommon.AscBrowser.retinaPixelRatio) >> 0
			};
		}

		var _centerX         = (_srcW / 2) >> 0;
		var _centerSlideX    = (dKoef * this.m_oLogicDocument.GetWidthMM() / 2) >> 0;
		var _hor_width_left  = Math.min(0, _centerX - (_centerSlideX - _bounds_slide.min_x) - this.SlideDrawer.CONST_BORDER);
		var _hor_width_right = Math.max(_srcW - 1, _centerX + (_bounds_slide.max_x - _centerSlideX) + this.SlideDrawer.CONST_BORDER);

		var _centerY           = (_srcH / 2) >> 0;
		var _centerSlideY      = (dKoef * this.m_oLogicDocument.GetHeightMM() / 2) >> 0;
		var _ver_height_top    = Math.min(0, _centerY - (_centerSlideY - _bounds_slide.min_y) - this.SlideDrawer.CONST_BORDER);
		var _ver_height_bottom = Math.max(_srcH - 1, _centerY + (_bounds_slide.max_y - _centerSlideY) + this.SlideDrawer.CONST_BORDER);

		var lWSlide = _hor_width_right - _hor_width_left + 1;
		var lHSlide = _ver_height_bottom - _ver_height_top + 1;

		var one_slide_width  = lWSlide;
		var one_slide_height = Math.max(lHSlide, _srcH);

		size.m_dDocumentPageWidth  = one_slide_width;
		size.m_dDocumentPageHeight = one_slide_height;

		size.m_dDocumentWidth  = one_slide_width;
		if (this.m_nVerticalSlideChangeOnScrollAllow) {
			size.m_dDocumentHeight = (one_slide_height * this.m_oDrawingDocument.GetSlidesCount()) >> 0;
		} else {
			size.m_dDocumentHeight = one_slide_height;
		}

		if (0 == this.m_oDrawingDocument.GetSlidesCount())
			size.m_dDocumentHeight = one_slide_height >> 0;

		if (this.m_nVerticalSlideChangeOnScrollAllow) {
			size.SlideScrollMIN = this.m_oDrawingDocument.SlideCurrent * one_slide_height;
		} else {
			size.SlideScrollMIN = 0;
		}
		size.SlideScrollMAX = size.SlideScrollMIN + one_slide_height - _srcH;

		if (0 == this.m_oDrawingDocument.GetSlidesCount())
		{
			size.SlideScrollMIN = 0;
			size.SlideScrollMAX = size.SlideScrollMIN + one_slide_height - _srcH;
		}

		return size;
	};

	this.CalculateDocumentSize = function(bIsAttack)
	{
		if (false === oThis.m_oApi.bInit_word_control)
		{
			oThis.UpdateScrolls();
			return;
		}

		var size = this.CalculateDocumentSizeInternal();

		this.m_dDocumentWidth      = size.m_dDocumentWidth;
		this.m_dDocumentHeight     = size.m_dDocumentHeight;
		this.m_dDocumentPageWidth  = size.m_dDocumentPageWidth;
		this.m_dDocumentPageHeight = size.m_dDocumentPageHeight;

		this.SlideScrollMIN = size.SlideScrollMIN;
		this.SlideScrollMAX = size.SlideScrollMAX;

		// теперь проверим необходимость перезуммирования
		if (1 == this.m_nZoomType)
		{
			if (true === this.zoom_FitToWidth())
				return;
		}
		if (2 == this.m_nZoomType)
		{
			if (true === this.zoom_FitToPage())
				return;
		}

		this.MainScrollLock();

		this.UpdateScrolls();

		this.MainScrollUnLock();

		this.Thumbnails.SlideWidth  = this.m_oLogicDocument.GetWidthMM();
		this.Thumbnails.SlideHeight = this.m_oLogicDocument.GetHeightMM();
		this.Thumbnails.CheckSizes();

		if (this.MobileTouchManager)
			this.MobileTouchManager.Resize();

		if (this.m_oApi.watermarkDraw)
		{
			this.m_oApi.watermarkDraw.zoom = this.m_nZoomValue / 100;
			this.m_oApi.watermarkDraw.Generate();
		}
	};

	this.CheckCalculateDocumentSize = function(_bounds)
	{
		if (false === oThis.m_oApi.bInit_word_control)
		{
			oThis.UpdateScrolls();
			return;
		}

		var size = this.CalculateDocumentSizeInternal();

		this.m_dDocumentWidth      = size.m_dDocumentWidth;
		this.m_dDocumentHeight     = size.m_dDocumentHeight;
		this.m_dDocumentPageWidth  = size.m_dDocumentPageWidth;
		this.m_dDocumentPageHeight = size.m_dDocumentPageHeight;

		this.SlideScrollMIN = size.SlideScrollMIN;
		this.SlideScrollMAX = size.SlideScrollMAX;

		this.MainScrollLock();

		this.UpdateScrolls();

		this.MainScrollUnLock();

		return false;
	};

	this.OnCalculatePagesPlace = function()
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		var canvas = this.m_oEditor.HtmlElement;
		if (null == canvas)
			return;

		var dKoef         = (this.m_nZoomValue * g_dKoef_mm_to_pix / 100);
		var _bounds_slide = this.SlideDrawer.BoundsChecker.Bounds;

		var _slideW = (dKoef * this.m_oLogicDocument.GetWidthMM()) >> 0;
		var _slideH = (dKoef * this.m_oLogicDocument.GetHeightMM()) >> 0;

		var _srcW = this.m_oEditor.HtmlElement.width;
		var _srcH = this.m_oEditor.HtmlElement.height;
		if (AscCommon.AscBrowser.isCustomScaling())
		{
			_srcW = (_srcW / AscCommon.AscBrowser.retinaPixelRatio) >> 0;
			_srcH = (_srcH / AscCommon.AscBrowser.retinaPixelRatio) >> 0;

			_bounds_slide = {
				min_x : (_bounds_slide.min_x / AscCommon.AscBrowser.retinaPixelRatio) >> 0,
				min_y : (_bounds_slide.min_y / AscCommon.AscBrowser.retinaPixelRatio) >> 0,
				max_x : (_bounds_slide.max_x / AscCommon.AscBrowser.retinaPixelRatio) >> 0,
				max_y : (_bounds_slide.max_y / AscCommon.AscBrowser.retinaPixelRatio) >> 0
			};
		}

		var _centerX         = (_srcW / 2) >> 0;
		var _centerSlideX    = (dKoef * this.m_oLogicDocument.GetWidthMM() / 2) >> 0;
		var _hor_width_left  = Math.min(0, _centerX - (_centerSlideX - _bounds_slide.min_x) - this.SlideDrawer.CONST_BORDER);
		var _hor_width_right = Math.max(_srcW - 1, _centerX + (_bounds_slide.max_x - _centerSlideX) + this.SlideDrawer.CONST_BORDER);

		var _centerY           = (_srcH / 2) >> 0;
		var _centerSlideY      = (dKoef * this.m_oLogicDocument.GetHeightMM() / 2) >> 0;
		var _ver_height_top    = Math.min(0, _centerY - (_centerSlideY - _bounds_slide.min_y) - this.SlideDrawer.CONST_BORDER);
		var _ver_height_bottom = Math.max(_srcH - 1, _centerX + (_bounds_slide.max_y - _centerSlideY) + this.SlideDrawer.CONST_BORDER);

		if (this.m_dScrollY <= this.SlideScrollMIN)
			this.m_dScrollY = this.SlideScrollMIN;
		if (this.m_dScrollY >= this.SlideScrollMAX)
			this.m_dScrollY = this.SlideScrollMAX;

		var _x = -this.m_dScrollX + _centerX - _centerSlideX - _hor_width_left;
		var _y = -(this.m_dScrollY - this.SlideScrollMIN) + _centerY - _centerSlideY - _ver_height_top;

		// теперь расчитаем какие нужны позиции, чтобы слайд находился по центру
		var _x_c                = _centerX - _centerSlideX;
		var _y_c                = _centerY - _centerSlideY;
		this.m_dScrollX_Central = _centerX - _centerSlideX - _hor_width_left - _x_c;
		this.m_dScrollY_Central = this.SlideScrollMIN + _centerY - _centerSlideY - _ver_height_top - _y_c;

		this.m_oDrawingDocument.SlideCurrectRect.left   = _x;
		this.m_oDrawingDocument.SlideCurrectRect.top    = _y;
		this.m_oDrawingDocument.SlideCurrectRect.right  = _x + _slideW;
		this.m_oDrawingDocument.SlideCurrectRect.bottom = _y + _slideH;

		if (this.m_oApi.isMobileVersion || this.m_oApi.isViewMode)
		{
			var lPage = this.m_oApi.GetCurrentVisiblePage();
			this.m_oApi.sendEvent("asc_onCurrentVisiblePage", this.m_oApi.GetCurrentVisiblePage());
		}

		if (this.m_bDocumentPlaceChangedEnabled)
			this.m_oApi.sendEvent("asc_onDocumentPlaceChanged");

		// update media control position
		AscCommon.g_specialPasteHelper.SpecialPasteButton_Update_Position();
	};

	// scrolls
	this.MainScrollLock   = function()
	{
		this.MainScrollsEnabledFlag++;
	};
	this.MainScrollUnLock = function()
	{
		this.MainScrollsEnabledFlag--;
		if (this.MainScrollsEnabledFlag < 0)
			this.MainScrollsEnabledFlag = 0;
	};

	this.GetVerticalScrollTo = function(y)
	{
		var dKoef = g_dKoef_mm_to_pix * this.m_nZoomValue / 100;
		return 5 + y * dKoef;
	};

	this.GetHorizontalScrollTo = function(x)
	{
		var dKoef = g_dKoef_mm_to_pix * this.m_nZoomValue / 100;
		return 5 + dKoef * x;
	};

	this.DeleteVerticalScroll = function()
	{
		this.m_oMainView.Bounds.R                    = 0;
		this.m_oPanelRight.HtmlElement.style.display = "none";
		this.OnResize();
	};

	this.verticalScrollCheckChangeSlide = function()
	{
		if (0 == this.m_nVerticalSlideChangeOnScrollInterval || !this.m_nVerticalSlideChangeOnScrollEnabled)
		{
			this.m_oScrollVer_.disableCurrentScroll = false;
			return true;
		}

		// защита от внутренних скроллах. мы превентим ТОЛЬКО самый верхний из onMouseWheel
		this.m_nVerticalSlideChangeOnScrollEnabled = false;

		var newTime = new Date().getTime();
		if (-1 == this.m_nVerticalSlideChangeOnScrollLast)
		{
			this.m_nVerticalSlideChangeOnScrollLast = newTime;
			this.m_oScrollVer_.disableCurrentScroll = false;
			return true;
		}

		var checkTime = this.m_nVerticalSlideChangeOnScrollLast + this.m_nVerticalSlideChangeOnScrollInterval;
		if (newTime < this.m_nVerticalSlideChangeOnScrollLast || newTime > checkTime)
		{
			this.m_nVerticalSlideChangeOnScrollLast = newTime;
			this.m_oScrollVer_.disableCurrentScroll = false;
			return true;
		}

		this.m_oScrollVer_.disableCurrentScroll = true;
		return false;
	};

	this.verticalScroll                = function(sender, scrollPositionY, maxY, isAtTop, isAtBottom)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		if (0 != this.MainScrollsEnabledFlag)
			return;

		if (this.StartVerticalScroll)
		{
			this.VerticalScrollOnMouseUp.ScrollY     = scrollPositionY;
			this.VerticalScrollOnMouseUp.ScrollY_max = maxY;

			if (this.m_nVerticalSlideChangeOnScrollAllow) {
				this.VerticalScrollOnMouseUp.SlideNum = (scrollPositionY * this.m_oDrawingDocument.GetSlidesCount() / Math.max(1, maxY)) >> 0;
				if (this.VerticalScrollOnMouseUp.SlideNum >= this.m_oDrawingDocument.GetSlidesCount())
					this.VerticalScrollOnMouseUp.SlideNum = this.m_oDrawingDocument.GetSlidesCount() - 1;
			} else {
				this.VerticalScrollOnMouseUp.SlideNum = this.m_oDrawingDocument.SlideCurrent;
			}

			this.m_oApi.sendEvent("asc_onPaintSlideNum", this.VerticalScrollOnMouseUp.SlideNum);
			return;
		}

		var lNumSlide;
		if (this.m_nVerticalSlideChangeOnScrollAllow) {
			lNumSlide = ((scrollPositionY / this.m_dDocumentPageHeight) + 0.01) >> 0; // 0.01 - ошибка округления!!;
		} else {
			lNumSlide = this.m_oDrawingDocument.SlideCurrent;
		}
		var _can_change_slide = true;
		if (-1 != this.ZoomFreePageNum && this.ZoomFreePageNum == this.m_oDrawingDocument.SlideCurrent)
			_can_change_slide = false;

		if (_can_change_slide)
		{
			if (lNumSlide != this.m_oDrawingDocument.SlideCurrent)
			{
				if (!this.verticalScrollCheckChangeSlide())
					return;

				if (this.IsGoToPageMAXPosition)
				{
					if (lNumSlide >= this.m_oDrawingDocument.SlideCurrent)
						this.IsGoToPageMAXPosition = false;
				}

				this.GoToPage(lNumSlide, false, true);
				return;
			}
			else if (this.SlideScrollMAX < scrollPositionY)
			{
				if (!this.verticalScrollCheckChangeSlide())
					return;

				this.IsGoToPageMAXPosition = false;
				this.GoToPage(this.m_oDrawingDocument.SlideCurrent + 1, false, true);
				return;
			}
		}
		else
		{
			if (!this.verticalScrollCheckChangeSlide())
				return;

			this.GoToPage(this.ZoomFreePageNum, false, true);
		}

		var oWordControl                       = oThis;
		oWordControl.m_dScrollY                = Math.max(0, Math.min(scrollPositionY, maxY));
		oWordControl.m_dScrollY_max            = maxY;
		oWordControl.m_bIsUpdateVerRuler       = true;
		oWordControl.m_bIsUpdateTargetNoAttack = true;

		oWordControl.IsGoToPageMAXPosition = false;

		if (oWordControl.m_bIsRePaintOnScroll === true)
			oWordControl.OnScroll();
	};
	this.verticalScrollMouseUp         = function(sender, e)
	{
		if (0 != this.MainScrollsEnabledFlag || !this.StartVerticalScroll)
			return;

		if (this.VerticalScrollOnMouseUp.SlideNum != this.m_oDrawingDocument.SlideCurrent)
			this.GoToPage(this.VerticalScrollOnMouseUp.SlideNum, false, true);
		else
		{
			this.StartVerticalScroll = false;
			this.m_oApi.sendEvent("asc_onEndPaintSlideNum");

			this.m_oScrollVerApi.scrollByY(0, true);
		}
	};
	this.CorrectSpeedVerticalScroll    = function(newScrollPos)
	{
		if (0 != this.MainScrollsEnabledFlag)
			return;

		this.StartVerticalScroll = true;

		var res = {isChange : false, Pos : newScrollPos};
		return res;
	};
	this.CorrectVerticalScrollByYDelta = function(delta)
	{
		if (0 != this.MainScrollsEnabledFlag)
			return;

		this.IsGoToPageMAXPosition = true;
		var res                    = {isChange : false, Pos : delta};

		if (this.m_dScrollY > this.SlideScrollMIN && (this.m_dScrollY + delta) < this.SlideScrollMIN)
		{
			res.Pos      = this.SlideScrollMIN - this.m_dScrollY;
			res.isChange = true;
		}
		else if (this.m_dScrollY < this.SlideScrollMAX && (this.m_dScrollY + delta) > this.SlideScrollMAX)
		{
			res.Pos      = this.SlideScrollMAX - this.m_dScrollY;
			res.isChange = true;
		}

		return res;
	};

	this.horizontalScroll = function(sender, scrollPositionX, maxX, isAtLeft, isAtRight)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		if (0 != this.MainScrollsEnabledFlag)
			return;

		var oWordControl                       = oThis;
		oWordControl.m_dScrollX                = scrollPositionX;
		oWordControl.m_dScrollX_max            = maxX;
		oWordControl.m_bIsUpdateHorRuler       = true;
		oWordControl.m_bIsUpdateTargetNoAttack = true;

		if (oWordControl.m_bIsRePaintOnScroll === true)
		{
			oWordControl.OnScroll();
		}
	};

	this.CreateScrollSettings = function()
	{
		var settings = new AscCommon.ScrollSettings();
		settings.screenW = this.m_oEditor.HtmlElement.width;
		settings.screenH = this.m_oEditor.HtmlElement.height;
		settings.vscrollStep = 45;
		settings.hscrollStep = 45;
		settings.contentH = this.m_dDocumentHeight;
		settings.contentW = this.m_dDocumentWidth;

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

		if (this.m_bIsRuler)
		{
			settings.screenAddH = this.m_oTopRuler_horRuler.HtmlElement.height;
		}

		settings.screenW = AscCommon.AscBrowser.convertToRetinaValue(settings.screenW);
		settings.screenH = AscCommon.AscBrowser.convertToRetinaValue(settings.screenH);
		settings.screenAddH = AscCommon.AscBrowser.convertToRetinaValue(settings.screenAddH);

		return settings;
	};

	this.UpdateScrolls = function()
	{
		var settings;
		if (window["NATIVE_EDITOR_ENJINE"])
			return;

		settings = this.CreateScrollSettings();
		settings.alwaysVisible = true;
		settings.isVerticalScroll = false;
		settings.isHorizontalScroll = true;

		if (this.m_oScrollHor_)
			this.m_oScrollHor_.Repos(settings, true, undefined);//unbind("scrollhorizontal")
		else
		{
			this.m_oScrollHor_ = new AscCommon.ScrollObject("id_horizontal_scroll", settings);
			this.m_oScrollHor_.bind("scrollhorizontal", function(evt)
			{
				oThis.horizontalScroll(this, evt.scrollD, evt.maxScrollX);
			});

			this.m_oScrollHor_.onLockMouse  = function(evt)
			{
				AscCommon.check_MouseDownEvent(evt, true);
				global_mouseEvent.LockMouse();
			};
			this.m_oScrollHor_.offLockMouse = function(evt)
			{
				AscCommon.check_MouseUpEvent(evt);
			};

			this.m_oScrollHorApi = this.m_oScrollHor_;
		}

		settings = this.CreateScrollSettings();
		settings.alwaysVisible = true;
		settings.isVerticalScroll = true;
		settings.isHorizontalScroll = false;

		if (this.m_oScrollVer_)
		{
			this.m_oScrollVer_.Repos(settings, undefined, true);//unbind("scrollvertical")
		}
		else
		{

			this.m_oScrollVer_ = new AscCommon.ScrollObject("id_vertical_scroll", settings);

			this.m_oScrollVer_.onLockMouse  = function(evt)
			{
				AscCommon.check_MouseDownEvent(evt, true);
				global_mouseEvent.LockMouse();
			};
			this.m_oScrollVer_.offLockMouse = function(evt)
			{
				AscCommon.check_MouseUpEvent(evt);
			};

			this.m_oScrollVer_.bind("scrollvertical", function(evt)
			{
				oThis.verticalScroll(this, evt.scrollD, evt.maxScrollY);
			});
			this.m_oScrollVer_.bind("mouseup.presentations", function(evt)
			{
				oThis.verticalScrollMouseUp(this, evt);
			});
			this.m_oScrollVer_.bind("correctVerticalScroll", function(yPos)
			{
				return oThis.CorrectSpeedVerticalScroll(yPos);
			});
			this.m_oScrollVer_.bind("correctVerticalScrollDelta", function(delta)
			{
				return oThis.CorrectVerticalScrollByYDelta(delta);
			});
			this.m_oScrollVerApi = this.m_oScrollVer_;
		}

		this.m_oApi.sendEvent("asc_onUpdateScrolls", this.m_dDocumentWidth, this.m_dDocumentHeight);

		this.m_dScrollX_max = this.m_oScrollHorApi.getMaxScrolledX();
		this.m_dScrollY_max = this.m_oScrollVerApi.getMaxScrolledY();

		if (this.m_dScrollX >= this.m_dScrollX_max)
			this.m_dScrollX = this.m_dScrollX_max;
		if (this.m_dScrollY >= this.m_dScrollY_max)
			this.m_dScrollY = this.m_dScrollY_max;
	};

	// search
	this.ToSearchResult = function()
	{
		var naviG = this.m_oDrawingDocument.CurrentSearchNavi;
		if (naviG.Page == -1)
			return;

		var navi = naviG.Place[0];
		var x    = navi.X;
		var y    = navi.Y;

		var rectSize = (navi.H * this.m_nZoomValue * g_dKoef_mm_to_pix / 100);
		var pos      = this.m_oDrawingDocument.ConvertCoordsToCursor2(x, y, navi.PageNum);

		if (true === pos.Error)
			return;

		var boxX = 0;
		var boxY = 0;

		var w = this.m_oEditor.HtmlElement.width;
		w /= AscCommon.AscBrowser.retinaPixelRatio;
		var h = this.m_oEditor.HtmlElement.height;
		h /= AscCommon.AscBrowser.retinaPixelRatio;

		var boxR = w - 2;
		var boxB = h - rectSize;

		var nValueScrollHor = 0;
		if (pos.X < boxX)
			nValueScrollHor = this.GetHorizontalScrollTo(x, navi.PageNum);

		if (pos.X > boxR)
		{
			var _mem        = x - g_dKoef_pix_to_mm * w * 100 / this.m_nZoomValue;
			nValueScrollHor = this.GetHorizontalScrollTo(_mem, navi.PageNum);
		}

		var nValueScrollVer = 0;
		if (pos.Y < boxY)
			nValueScrollVer = this.GetVerticalScrollTo(y, navi.PageNum);

		if (pos.Y > boxB)
		{
			var _mem        = y + navi.H + 10 - g_dKoef_pix_to_mm * h * 100 / this.m_nZoomValue;
			nValueScrollVer = this.GetVerticalScrollTo(_mem, navi.PageNum);
		}

		var isNeedScroll = false;
		if (0 != nValueScrollHor)
		{
			isNeedScroll                   = true;
			this.m_bIsUpdateTargetNoAttack = true;
			var temp                       = nValueScrollHor * this.m_dScrollX_max / (this.m_dDocumentWidth - w);
			this.m_oScrollHorApi.scrollToX(parseInt(temp), false);
		}
		if (0 != nValueScrollVer)
		{
			isNeedScroll                   = true;
			this.m_bIsUpdateTargetNoAttack = true;
			var temp                       = nValueScrollVer * this.m_dScrollY_max / (this.m_dDocumentHeight - h);
			this.m_oScrollVerApi.scrollToY(parseInt(temp), false);
		}

		if (true === isNeedScroll)
		{
			this.OnScroll();
			return;
		}

		this.OnUpdateOverlay();
	};

	// events
	this.onMouseDownTarget = function(e)
	{
		if (oThis.m_oDrawingDocument.TargetHtmlElementOnSlide)
			return oThis.onMouseDown(e);
	};
	this.onMouseMoveTarget = function(e)
	{
		if (oThis.m_oDrawingDocument.TargetHtmlElementOnSlide)
			return oThis.onMouseMove(e);
	};
	this.onMouseUpTarget = function(e)
	{
		if (oThis.m_oDrawingDocument.TargetHtmlElementOnSlide)
			return oThis.onMouseUp(e);
	};

	this.onMouseDown = function(e)
	{
		if (oThis.MobileTouchManager && oThis.MobileTouchManager.checkTouchEvent(e))
		{
			oThis.MobileTouchManager.startTouchingInProcess();
			let res = oThis.MobileTouchManager.mainOnTouchStart(e);
			oThis.MobileTouchManager.stopTouchingInProcess();
			return res;
		}

		oThis.m_oApi.checkInterfaceElementBlur();
		oThis.m_oApi.checkLastWork();

		if (false === oThis.m_oApi.bInit_word_control)
			return;

		var oWordControl = oThis;

		if (oWordControl.m_oDrawingDocument.TransitionSlide.IsPlaying())
			oWordControl.m_oDrawingDocument.TransitionSlide.End(true);

		// после fullscreen возможно изменение X, Y после вызова Resize.
		oWordControl.checkBodyOffset();

		if (!oThis.m_bIsIE)
		{
			if (e.preventDefault)
				e.preventDefault();
			else
				e.returnValue = false;
		}

		if (AscCommon.g_inputContext && AscCommon.g_inputContext.externalChangeFocus())
			return;

		oWordControl.Thumbnails.SetFocusElement(FOCUS_OBJECT_MAIN);

		var _xOffset = oWordControl.X;
		var _yOffset = oWordControl.Y;

		if (true === oWordControl.m_bIsRuler)
		{
			_xOffset += (5 * g_dKoef_mm_to_pix);
			_yOffset += (7 * g_dKoef_mm_to_pix);
		}

		if (window['closeDialogs'] != undefined)
			window['closeDialogs']();

		AscCommon.check_MouseDownEvent(e, true);
		global_mouseEvent.LockMouse();

		if ((0 == global_mouseEvent.Button) || (undefined == global_mouseEvent.Button))
		{
			global_mouseEvent.Button    = 0;
			oWordControl.m_bIsMouseLock = true;
		}

		if ((0 == global_mouseEvent.Button) || (undefined == global_mouseEvent.Button) || (2 == global_mouseEvent.Button))
		{
			var pos = oWordControl.m_oDrawingDocument.ConvertCoordsFromCursor2(global_mouseEvent.X, global_mouseEvent.Y);
			if (pos.Page == -1)
				return;

			var ret = oWordControl.m_oDrawingDocument.checkMouseDown_Drawing(pos);
			if (ret === true)
			{
				if (-1 == oWordControl.m_oTimerScrollSelect)
					oWordControl.m_oTimerScrollSelect = setInterval(oWordControl.SelectWheel, 20);

				AscCommon.stopEvent(e);
				return;
			}

			oWordControl.StartUpdateOverlay();
			oWordControl.m_oDrawingDocument.m_lCurrentPage = pos.Page;

			if(!oThis.m_oApi.isEyedropperStarted())
			{
				oWordControl.m_oLogicDocument.OnMouseDown(global_mouseEvent, pos.X, pos.Y, pos.Page);
			}
			oWordControl.EndUpdateOverlay();
		}
		else
		{
			var pos = oWordControl.m_oDrawingDocument.ConvertCoordsFromCursor2(global_mouseEvent.X, global_mouseEvent.Y);
			if (pos.Page == -1)
				return;

			oWordControl.m_oDrawingDocument.m_lCurrentPage = pos.Page;
		}

		if (-1 == oWordControl.m_oTimerScrollSelect)
		{
			oWordControl.m_oTimerScrollSelect = setInterval(oWordControl.SelectWheel, 20);
		}

		oWordControl.Thumbnails.SetFocusElement(FOCUS_OBJECT_MAIN);
	};

	this.onMouseMove  = function(e)
	{
		if (oThis.MobileTouchManager && oThis.MobileTouchManager.checkTouchEvent(e))
		{
			oThis.MobileTouchManager.startTouchingInProcess();
			let res = oThis.MobileTouchManager.mainOnTouchMove(e);
			oThis.MobileTouchManager.stopTouchingInProcess();
			return res;
		}

		oThis.m_oApi.checkLastWork();

		if (false === oThis.m_oApi.bInit_word_control)
			return;

		var oWordControl = oThis;

		if (e.preventDefault)
			e.preventDefault();
		else
			e.returnValue = false;

		AscCommon.check_MouseMoveEvent(e);
		var pos = oWordControl.m_oDrawingDocument.ConvertCoordsFromCursor2(global_mouseEvent.X, global_mouseEvent.Y);
		if (pos.Page == -1)
			return;

		if (oWordControl.m_oDrawingDocument.m_sLockedCursorType != "")
			oWordControl.m_oDrawingDocument.SetCursorType("default");

		if(oThis.m_oApi.isEyedropperStarted())
		{
			let oMainPos = oWordControl.m_oMainParent.AbsolutePosition;
			let oParentPos = oWordControl.m_oMainView.AbsolutePosition;
			let nX  = global_mouseEvent.X - oWordControl.X - (oMainPos.L + oParentPos.L) * AscCommon.g_dKoef_mm_to_pix;
			let nY  = global_mouseEvent.Y - oWordControl.Y - (oMainPos.T + oParentPos.T) * AscCommon.g_dKoef_mm_to_pix;
			nX = AscCommon.AscBrowser.convertToRetinaValue(nX, true);
			nY = AscCommon.AscBrowser.convertToRetinaValue(nY, true);
			oThis.m_oApi.checkEyedropperColor(nX, nY);
			oThis.m_oApi.sync_MouseMoveStartCallback();
			let MMData = new AscCommon.CMouseMoveData();
			let Coords = oWordControl.m_oDrawingDocument.ConvertCoordsToCursorWR(pos.X, pos.Y, pos.Page, null, true);
			MMData.X_abs = Coords.X;
			MMData.Y_abs = Coords.Y;
			const oEyedropperColor = oThis.m_oApi.getEyedropperColor();
			if(oEyedropperColor)
			{
				MMData.EyedropperColor = oEyedropperColor;
				MMData.Type = Asc.c_oAscMouseMoveDataTypes.Eyedropper;
				oWordControl.m_oDrawingDocument.SetCursorType(AscCommon.Cursors.Eyedropper, MMData);
			}
			else
			{
				oWordControl.m_oDrawingDocument.SetCursorType("default");
			}
			oThis.m_oApi.sync_MouseMoveEndCallback();
			return;
		}

		oWordControl.StartUpdateOverlay();
		var is_drawing = oWordControl.m_oDrawingDocument.checkMouseMove_Drawing(pos);
		if (is_drawing === true)
			return;

		oWordControl.m_oLogicDocument.OnMouseMove(global_mouseEvent, pos.X, pos.Y, pos.Page);

		oWordControl.EndUpdateOverlay();
	};
	this.onMouseMove2 = function()
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		var oWordControl = oThis;
		var pos          = oWordControl.m_oDrawingDocument.ConvertCoordsFromCursor2(global_mouseEvent.X, global_mouseEvent.Y);
		if (pos.Page == -1)
			return;

		oWordControl.StartUpdateOverlay();

		var is_drawing = oWordControl.m_oDrawingDocument.checkMouseMove_Drawing(pos);
		if (is_drawing === true)
			return;

		oWordControl.m_oLogicDocument.OnMouseMove(global_mouseEvent, pos.X, pos.Y, pos.Page);
		oWordControl.EndUpdateOverlay();
	};
	this.onMouseUp    = function(e, bIsWindow)
	{
		if (oThis.MobileTouchManager && oThis.MobileTouchManager.checkTouchEvent(e))
		{
			oThis.MobileTouchManager.startTouchingInProcess();
			let res = oThis.MobileTouchManager.mainOnTouchEnd(e);
			oThis.MobileTouchManager.stopTouchingInProcess();
			return res;
		}

		oThis.m_oApi.checkLastWork();

		if (false === oThis.m_oApi.bInit_word_control)
			return;

		var oWordControl = oThis;
		if (!global_mouseEvent.IsLocked)
			return;

		AscCommon.check_MouseUpEvent(e);

		var pos = oWordControl.m_oDrawingDocument.ConvertCoordsFromCursor2(global_mouseEvent.X, global_mouseEvent.Y);
		if (pos.Page == -1)
			return;

		oWordControl.UnlockCursorTypeOnMouseUp();
		oWordControl.m_bIsMouseLock = false;
		if (oWordControl.m_oDrawingDocument.TableOutlineDr.bIsTracked)
		{
			oWordControl.m_oDrawingDocument.TableOutlineDr.checkMouseUp(global_mouseEvent.X, global_mouseEvent.Y, oWordControl);
			oWordControl.m_oLogicDocument.Document_UpdateInterfaceState();
			oWordControl.m_oLogicDocument.Document_UpdateRulersState();

			if (-1 != oWordControl.m_oTimerScrollSelect)
			{
				clearInterval(oWordControl.m_oTimerScrollSelect);
				oWordControl.m_oTimerScrollSelect = -1;
			}
			oWordControl.OnUpdateOverlay();
			return;
		}

		if (-1 != oWordControl.m_oTimerScrollSelect)
		{
			clearInterval(oWordControl.m_oTimerScrollSelect);
			oWordControl.m_oTimerScrollSelect = -1;
		}

		oWordControl.m_bIsMouseUpSend = true;

		oWordControl.StartUpdateOverlay();

		var is_drawing = oWordControl.m_oDrawingDocument.checkMouseUp_Drawing(pos);
		if (is_drawing === true)
			return;

		if(!oThis.checkFinishEyedropper())
		{
			oWordControl.m_oLogicDocument.OnMouseUp(global_mouseEvent, pos.X, pos.Y, pos.Page);
		}

		oWordControl.m_bIsMouseUpSend = false;
		//        oWordControl.m_oLogicDocument.Document_UpdateInterfaceState();
		oWordControl.m_oLogicDocument.Document_UpdateRulersState();

		oWordControl.EndUpdateOverlay();
	};

	this.onMouseUpMainSimple = function()
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		var oWordControl = oThis;

		global_mouseEvent.Type = AscCommon.g_mouse_event_type_up;

		AscCommon.MouseUpLock.MouseUpLockedSend = true;

		global_mouseEvent.Sender = null;

		global_mouseEvent.UnLockMouse();

		global_mouseEvent.IsPressed = false;

		if (-1 != oWordControl.m_oTimerScrollSelect)
		{
			clearInterval(oWordControl.m_oTimerScrollSelect);
			oWordControl.m_oTimerScrollSelect = -1;
		}
	};

	this.onMouseUpExternal = function(x, y)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		var oWordControl = oThis;

		//---
		global_mouseEvent.X = x;
		global_mouseEvent.Y = y;

		global_mouseEvent.Type = AscCommon.g_mouse_event_type_up;

		AscCommon.MouseUpLock.MouseUpLockedSend = true;
		global_mouseEvent.Sender                = null;

		global_mouseEvent.UnLockMouse();

		global_mouseEvent.IsPressed = false;

		//---
		var pos = oWordControl.m_oDrawingDocument.ConvertCoordsFromCursor2(global_mouseEvent.X, global_mouseEvent.Y);
		if (pos.Page == -1)
			return;

		oWordControl.UnlockCursorTypeOnMouseUp();

		oWordControl.m_bIsMouseLock = false;

		if (-1 != oWordControl.m_oTimerScrollSelect)
		{
			clearInterval(oWordControl.m_oTimerScrollSelect);
			oWordControl.m_oTimerScrollSelect = -1;
		}

		oWordControl.StartUpdateOverlay();

		oWordControl.m_bIsMouseUpSend = true;
		if(!oThis.checkFinishEyedropper())
		{
			oWordControl.m_oLogicDocument.OnMouseUp(global_mouseEvent, pos.X, pos.Y, pos.Page);
		}
		oWordControl.m_bIsMouseUpSend = false;
		oWordControl.m_oLogicDocument.Document_UpdateInterfaceState();
		oWordControl.m_oLogicDocument.Document_UpdateRulersState();

		oWordControl.EndUpdateOverlay();
	};

	this.onMouseWhell = function(e)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		if (undefined !== window["AscDesktopEditor"])
		{
			if (false === window["AscDesktopEditor"]["CheckNeedWheel"]())
				return;
		}

		if (global_mouseEvent.IsLocked)
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

		let values = AscCommon.checkMouseWhell(e, {
			isSupportBidirectional : false,
			isAllowHorizontal : true,
			isUseMaximumDelta : true
		});

		oThis.m_nVerticalSlideChangeOnScrollEnabled = true;

		if (0 !== values.x)
			oThis.m_oScrollHorApi.scrollBy(values.x, 0, false);
		if (0 !== values.y)
			oThis.m_oScrollVerApi.scrollBy(0, values.y, false);

		oThis.m_nVerticalSlideChangeOnScrollEnabled = false;

		if (e.preventDefault)
			e.preventDefault();
		else
			e.returnValue = false;
		return false;
	};

	this.onKeyUp = function(e)
	{
		global_keyboardEvent.AltKey   = false;
		global_keyboardEvent.CtrlKey  = false;
		global_keyboardEvent.ShiftKey = false;
		global_keyboardEvent.AltGr    = false;
	};

	this.onKeyDown = function(e)
	{
		oThis.m_oApi.checkLastWork();

		if (oThis.m_oApi.isLongAction())
		{
			e.preventDefault();
			return;
		}

		var oWordControl = oThis;
		if (false === oWordControl.m_oApi.bInit_word_control)
		{
			AscCommon.check_KeyboardEvent2(e);
			e.preventDefault();
			return;
		}

		if (oWordControl.IsFocus === false && e.emulated !== true)
			return;

		if (oWordControl.m_oApi.isLongAction() || oWordControl.m_bIsMouseLock === true)
		{
			AscCommon.check_KeyboardEvent2(e);
			e.preventDefault();
			return;
		}

		if (oWordControl.Thumbnails.FocusObjType == FOCUS_OBJECT_THUMBNAILS)
		{
			if (0 == oWordControl.Splitter1Pos)
			{
				// табнейлы не видны. Чего же тогда обрабатывать им клавиатуру
				e.preventDefault();
				return false;
			}

			var ret = oWordControl.Thumbnails.onKeyDown(e);
			if (false === ret)
				return false;
			if (undefined === ret)
				return;
		}

		if (oWordControl.m_oDrawingDocument.TransitionSlide.IsPlaying())
			oWordControl.m_oDrawingDocument.TransitionSlide.End(true);

		var oWordControl = oThis;
		if (false === oWordControl.m_oApi.bInit_word_control || oWordControl.IsFocus === false && e.emulated !== true || oWordControl.m_oApi.isLongAction() || oWordControl.m_bIsMouseLock === true)
			return;

		AscCommon.check_KeyboardEvent(e);

		oWordControl.StartUpdateOverlay();

		oWordControl.IsKeyDownButNoPress = true;

		var _ret_mouseDown          = oWordControl.m_oLogicDocument.OnKeyDown(global_keyboardEvent);
		oWordControl.bIsUseKeyPress = ((_ret_mouseDown & keydownresult_PreventKeyPress) != 0) ? false : true;

		if ((_ret_mouseDown & keydownresult_PreventDefault) != 0)
		{
			// убираем превент с альтом. Уж больно итальянцы недовольны.
			e.preventDefault();
		}

		oWordControl.EndUpdateOverlay();
	};

	this.onKeyDownNoActiveControl = function(e)
	{
		var bSendToEditor = false;

		if (e.CtrlKey && !e.ShiftKey)
		{
			switch (e.KeyCode)
			{
				case 80: // P
				case 83: // S
					bSendToEditor = true;
					break;
				default:
					break;
			}
		}

		return bSendToEditor;
	};

	this.onKeyDownTBIM = function(e)
	{
		var oWordControl = oThis;
		if (false === oWordControl.m_oApi.bInit_word_control || oWordControl.IsFocus === false || oWordControl.m_oApi.isLongAction() || oWordControl.m_bIsMouseLock === true)
			return;

		AscCommon.check_KeyboardEvent(e);

		oWordControl.IsKeyDownButNoPress = true;

		oWordControl.StartUpdateOverlay();

		var _ret_mouseDown          = oWordControl.m_oLogicDocument.OnKeyDown(global_keyboardEvent);
		oWordControl.bIsUseKeyPress = ((_ret_mouseDown & keydownresult_PreventKeyPress) != 0) ? false : true;

		oWordControl.EndUpdateOverlay();

		if ((_ret_mouseDown & keydownresult_PreventDefault) != 0)
		{
			// убираем превент с альтом. Уж больно итальянцы недовольны.
			e.preventDefault();
			return false;
		}
	};

	this.onKeyPress = function(e)
	{
		if (window.GlobalPasteFlag || window.GlobalCopyFlag)
			return;

		if (oThis.Thumbnails.FocusObjType == FOCUS_OBJECT_THUMBNAILS)
		{
			return;
		}

		var oWordControl = oThis;
		if (false === oWordControl.m_oApi.bInit_word_control || oWordControl.IsFocus === false || oWordControl.m_oApi.isLongAction() || oWordControl.m_bIsMouseLock === true)
			return;

		if (window.opera && !oWordControl.IsKeyDownButNoPress)
		{
			oWordControl.onKeyDown(e);
		}

		oWordControl.IsKeyDownButNoPress = false;

		if (false === oWordControl.bIsUseKeyPress)
			return;

		if (null == oWordControl.m_oLogicDocument)
			return;

		AscCommon.check_KeyboardEvent(e);

		oWordControl.StartUpdateOverlay();

		var retValue = oWordControl.m_oLogicDocument.OnKeyPress(global_keyboardEvent);
		if (true === retValue)
		{
			e.preventDefault();
		}

		oWordControl.EndUpdateOverlay();
	};

	this.SelectWheel = function()
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		var oWordControl = oThis;
		var positionMinY = oWordControl.m_oMainContent.AbsolutePosition.T * g_dKoef_mm_to_pix + oWordControl.Y;
		if (oWordControl.m_bIsRuler)
			positionMinY = (oWordControl.m_oMainContent.AbsolutePosition.T + oWordControl.m_oTopRuler_horRuler.AbsolutePosition.B) * g_dKoef_mm_to_pix +
				oWordControl.Y;

		var positionMaxY = oWordControl.m_oMainContent.AbsolutePosition.B * g_dKoef_mm_to_pix + oWordControl.Y;

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
		var positionMinX = oWordControl.m_oMainParent.AbsolutePosition.L * g_dKoef_mm_to_pix + oWordControl.X;
		if (oWordControl.m_bIsRuler)
			positionMinX += oWordControl.m_oLeftRuler.AbsolutePosition.R * g_dKoef_mm_to_pix;

		var positionMaxX = oWordControl.m_oMainParent.AbsolutePosition.R * g_dKoef_mm_to_pix + oWordControl.X - oWordControl.ScrollWidthPx;

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

		if (0 != scrollYVal)
		{
			if (((oWordControl.m_dScrollY + scrollYVal) >= oWordControl.SlideScrollMIN) && ((oWordControl.m_dScrollY + scrollYVal) <= oWordControl.SlideScrollMAX))
				oWordControl.m_oScrollVerApi.scrollByY(scrollYVal, false);
		}
		if (0 != scrollXVal)
			oWordControl.m_oScrollHorApi.scrollByX(scrollXVal, false);

		if (scrollXVal != 0 || scrollYVal != 0)
			oWordControl.onMouseMove2();
	};

	// overlay
	this.StartUpdateOverlay = function()
	{
		this.IsUpdateOverlayOnlyEnd = true;
	};
	this.EndUpdateOverlay   = function()
	{
		this.IsUpdateOverlayOnlyEnd = false;
		if (this.IsUpdateOverlayOnEndCheck)
			this.OnUpdateOverlay();
		this.IsUpdateOverlayOnEndCheck = false;
	};

	this.OnUpdateOverlay = function()
	{
		if (this.IsUpdateOverlayOnlyEnd)
		{
			this.IsUpdateOverlayOnEndCheck = true;
			return false;
		}

		this.m_oApi.checkLastWork();

		var overlay = this.m_oOverlayApi;

		overlay.SetBaseTransform();
		overlay.Clear();
		var ctx = overlay.m_oContext;

		var drDoc = this.m_oDrawingDocument;
		drDoc.SelectionMatrix = null;

		if (drDoc.SlideCurrent >= drDoc.m_oLogicDocument.GetSlidesCount())
			drDoc.SlideCurrent = drDoc.m_oLogicDocument.GetSlidesCount() - 1;

		if (drDoc.m_bIsSearching)
		{
			ctx.fillStyle = "rgba(255,200,0,1)";
			ctx.beginPath();

			var drDoc = this.m_oDrawingDocument;
			drDoc.DrawSearch(overlay);

			ctx.globalAlpha = 0.5;
			ctx.fill();
			ctx.beginPath();
			ctx.globalAlpha = 1.0;

			if (null != drDoc.CurrentSearchNavi)
			{
				ctx.globalAlpha = 0.2;
				ctx.fillStyle   = "rgba(51,102,204,255)";

				var places = drDoc.CurrentSearchNavi.Place;
				for (var i = 0; i < places.length; i++)
				{
					var place = places[i];
					if (drDoc.SlideCurrent == place.PageNum)
					{
						drDoc.DrawSearchCur(overlay, place);
					}
				}

				ctx.fill();
				ctx.globalAlpha = 1.0;
			}
		}

		if (drDoc.m_bIsSelection)
		{
			ctx.fillStyle   = "rgba(51,102,204,255)";
			ctx.strokeStyle = "#9ADBFE";
			ctx.lineWidth = Math.round(AscCommon.AscBrowser.retinaPixelRatio);

			ctx.beginPath();

			if (drDoc.SlideCurrent != -1)
				this.m_oLogicDocument.GetCurrentSlide().drawSelect(1);

			ctx.globalAlpha = 0.2;
			ctx.fill();
			ctx.globalAlpha = 1.0;
			ctx.stroke();
			ctx.beginPath();
			ctx.globalAlpha = 1.0;

			if (this.MobileTouchManager)
				this.MobileTouchManager.CheckSelect(overlay);
		}

		if (drDoc.MathTrack.IsActive())
		{
			var dGlobalAplpha = ctx.globalAlpha;
			ctx.globalAlpha = 1.0;
			drDoc.DrawMathTrack(overlay);
			ctx.globalAlpha = dGlobalAplpha;
		}


		if (this.MobileTouchManager)
			this.MobileTouchManager.CheckTableRules(overlay);

		ctx.globalAlpha = 1.0;
		ctx             = null;

		if (this.m_oLogicDocument != null && drDoc.SlideCurrent >= 0)
		{
			//todo
			// let oSlide = this.m_oLogicDocument.GetCurrentSlide();
			// oSlide.drawSelect(2);
			//
			// var elements = oSlide.graphicObjects;
			// if (!elements.canReceiveKeyPress() && -1 != drDoc.SlideCurrent)
			// {
			// 	var drawPage = drDoc.SlideCurrectRect;
			// 	drDoc.AutoShapesTrack.init(overlay, drawPage.left, drawPage.top, drawPage.right, drawPage.bottom, this.m_oLogicDocument.GetWidthMM(), this.m_oLogicDocument.GetHeightMM());
			//
			// 	elements.DrawOnOverlay(drDoc.AutoShapesTrack);
			// 	drDoc.AutoShapesTrack.CorrectOverlayBounds();
			//
			// 	overlay.SetBaseTransform();
			// }
		}

		if (drDoc.InlineTextTrackEnabled && null != drDoc.InlineTextTrack)
		{
			var _oldPage        = drDoc.AutoShapesTrack.PageIndex;
			var _oldCurPageInfo = drDoc.AutoShapesTrack.CurrentPageInfo;

			drDoc.AutoShapesTrack.PageIndex = drDoc.InlineTextTrackPage;
			drDoc.AutoShapesTrack.DrawInlineMoveCursor(drDoc.InlineTextTrack.X, drDoc.InlineTextTrack.Y, drDoc.InlineTextTrack.Height, drDoc.InlineTextTrack.transform, null);

			drDoc.AutoShapesTrack.PageIndex       = _oldPage;
			drDoc.AutoShapesTrack.CurrentPageInfo = _oldCurPageInfo;
		}

		if (drDoc.placeholders.objects.length > 0 && drDoc.SlideCurrent >= 0)
		{
			var rectSlide = {};
			rectSlide.left = AscCommon.AscBrowser.convertToRetinaValue(drDoc.SlideCurrectRect.left, true);
			rectSlide.top = AscCommon.AscBrowser.convertToRetinaValue(drDoc.SlideCurrectRect.top, true);
			rectSlide.right = AscCommon.AscBrowser.convertToRetinaValue(drDoc.SlideCurrectRect.right, true);
			rectSlide.bottom = AscCommon.AscBrowser.convertToRetinaValue(drDoc.SlideCurrectRect.bottom, true);
			drDoc.placeholders.draw(overlay, drDoc.SlideCurrent, rectSlide, this.m_oLogicDocument.GetWidthMM(), this.m_oLogicDocument.GetHeightMM());
		}

		drDoc.DrawHorVerAnchor();

		if (this.MobileTouchManager)
		{
			let targetElement = (this.m_oDrawingDocument && this.m_oDrawingDocument.isDrawTargetGlass()) ? this.m_oDrawingDocument.TargetHtmlElement : null;
			this.MobileTouchManager.CheckGlass(overlay, this.m_oEditor.HtmlElement, targetElement);
		}

		return true;
	};

	// Eyedropper
	this.checkFinishEyedropper = function()
	{
		if(oThis.m_oApi.isEyedropperStarted())
		{
			oThis.m_oApi.finishEyedropper();
			const oPos = oThis.m_oDrawingDocument.ConvertCoordsFromCursor2(global_mouseEvent.X, global_mouseEvent.Y);
			if (oPos.Page !== -1)
			{
				oThis.m_oLogicDocument.OnMouseMove(global_mouseEvent, oPos.X, oPos.Y, oPos.Page);
			}
			return true;
		}
		return false;
	};

	this.UnlockCursorTypeOnMouseUp = function()
	{
		if (this.m_oApi.isInkDrawerOn())
			return;
		this.m_oDrawingDocument.UnlockCursorType();
	};


	// resize
	this.OnResize = function(isAttack)
	{
		AscCommon.AscBrowser.checkZoom();

		var isNewSize = this.checkBodySize();
		if (!isNewSize && this.retinaScaling === AscCommon.AscBrowser.retinaPixelRatio && false === isAttack)
		{
			return;
		}

		if (this.MobileTouchManager)
			this.MobileTouchManager.Resize_Before();

		if (this.m_oApi.printPreview)
			this.m_oApi.printPreview.resize();

		var isDesktopVersion = (undefined !== window["AscDesktopEditor"]) ? true : false;

		if (this.Splitter1Pos > 0.1 && !isDesktopVersion)
		{
			var maxSplitterThMax = g_dKoef_pix_to_mm * this.Width / 3;
			if (maxSplitterThMax > 80)
				maxSplitterThMax = 80;

			this.Splitter1PosMin = maxSplitterThMax >> 2;
			this.Splitter1PosMax = maxSplitterThMax >> 0;

			this.Splitter1Pos = this.Splitter1PosSetUp;
			if (this.Splitter1Pos < this.Splitter1PosMin)
				this.Splitter1Pos = this.Splitter1PosMin;
			if (this.Splitter1Pos > this.Splitter1PosMax)
				this.Splitter1Pos = this.Splitter1PosMax;

			this.OnResizeSplitter(true);
		}

		//console.log("resize");
		this.CheckRetinaDisplay();

		this.m_oBody.Resize(this.Width * g_dKoef_pix_to_mm, this.Height * g_dKoef_pix_to_mm, this);
		this.onButtonTabsDraw();

		if (AscCommon.g_inputContext)
			AscCommon.g_inputContext.onResize("id_main_parent");

		// теперь проверим необходимость перезуммирования
		if (1 == this.m_nZoomType && 0 != this.m_dDocumentPageWidth && 0 != this.m_dDocumentPageHeight)
		{
			if (true === this.zoom_FitToWidth())
			{
				this.m_oBoundsController.ClearNoAttack();
				this.onTimerScroll_sync();

				this.FullRulersUpdate();
				return;
			}
		}
		if (2 == this.m_nZoomType && 0 != this.m_dDocumentPageWidth && 0 != this.m_dDocumentPageHeight)
		{
			if (true === this.zoom_FitToPage())
			{
				this.m_oBoundsController.ClearNoAttack();
				this.onTimerScroll_sync();

				this.FullRulersUpdate();
				return;
			}
		}

		this.Thumbnails.m_bIsUpdate = true;
		this.CalculateDocumentSize();

		this.m_bIsUpdateTargetNoAttack = true;
		this.m_bIsRePaintOnScroll      = true;

		this.m_oBoundsController.ClearNoAttack();

		this.UpdateScrolls();

		this.OnScroll();
		this.onTimerScroll_sync(true);

		if (this.MobileTouchManager)
			this.MobileTouchManager.Resize_After();

		this.FullRulersUpdate();

		if (AscCommon.g_imageControlsStorage)
			AscCommon.g_imageControlsStorage.resize();
	};

	this.OnResize2 = function(isAttack)
	{
		this.m_oBody.Resize(this.Width * g_dKoef_pix_to_mm, this.Height * g_dKoef_pix_to_mm, this);
		this.onButtonTabsDraw();

		// теперь проверим необходимость перезуммирования
		if (1 == this.m_nZoomType)
		{
			if (true === this.zoom_FitToWidth())
			{
				this.m_oBoundsController.ClearNoAttack();
				this.onTimerScroll_sync();

				this.FullRulersUpdate();
				return;
			}
		}
		if (2 == this.m_nZoomType)
		{
			if (true === this.zoom_FitToPage())
			{
				this.m_oBoundsController.ClearNoAttack();
				this.onTimerScroll_sync();

				this.FullRulersUpdate();
				return;
			}
		}

		this.m_bIsUpdateHorRuler = true;
		this.m_bIsUpdateVerRuler = true;

		this.m_oHorRuler.RepaintChecker.BlitAttack = true;
		this.m_oVerRuler.RepaintChecker.BlitAttack = true;

		this.Thumbnails.m_bIsUpdate = true;
		this.CalculateDocumentSize();

		this.m_bIsUpdateTargetNoAttack = true;
		this.m_bIsRePaintOnScroll      = true;

		this.m_oBoundsController.ClearNoAttack();
		this.OnScroll();
		this.onTimerScroll_sync(true);

		this.FullRulersUpdate();
	};

	// current page
	this.SetCurrentPage = function()
	{
		var drDoc = this.m_oDrawingDocument;
		if (0 <= drDoc.SlideCurrent && drDoc.SlideCurrent < drDoc.GetSlidesCount())
		{
			this.CreateBackgroundHorRuler();
			this.CreateBackgroundVerRuler();
		}

		this.m_bIsUpdateHorRuler = true;
		this.m_bIsUpdateVerRuler = true;

		this.OnScroll();

		this.m_oApi.sync_currentPageCallback(drDoc.m_lCurrentPage);
	};
	this.GetSlidesCount = function()
	{
		//todo
		return this.m_oLogicDocument.getCountPages();
		//return this.m_oDrawingDocument.GetSlidesCount();
	};

	this.GoToPage = function(lPageNum, isFromZoom, isScroll, bIsAttack)
	{
		var drDoc = this.m_oDrawingDocument;

		if (!this.m_oScrollVerApi)
		{
			// сборка файлов
			return;
		}
		if(this.m_oApi.isEyedropperStarted() && drDoc.SlideCurrent !== lPageNum)
		{
			this.m_oApi.cancelEyedropper();
		}

		if (this.m_oDrawingDocument.TransitionSlide.IsPlaying())
			this.m_oDrawingDocument.TransitionSlide.End(true);


		if (lPageNum != -1 && (lPageNum < 0 || lPageNum >= drDoc.GetSlidesCount()))
			return;

		this.Thumbnails.LockMainObjType = true;
		this.StartVerticalScroll        = false;
		this.m_oApi.sendEvent("asc_onEndPaintSlideNum");

		var _bIsUpdate = (drDoc.SlideCurrent != lPageNum);

		this.ZoomFreePageNum = lPageNum;
		drDoc.SlideCurrent   = lPageNum;
		this.m_oLogicDocument.setCurrentPage(lPageNum);

		this.SlideDrawer.CheckSlide(drDoc.SlideCurrent);

		if (true !== isFromZoom)
		{
			this.m_oLogicDocument.Document_UpdateInterfaceState();
		}

		this.CalculateDocumentSize(false);

		this.Thumbnails.SelectPage(lPageNum);

		this.CreateBackgroundHorRuler();
		this.CreateBackgroundVerRuler();

		this.m_bIsUpdateHorRuler = true;
		this.m_bIsUpdateVerRuler = true;

		this.OnCalculatePagesPlace();

		//this.m_oScrollVerApi.scrollTo(0, drDoc.SlideCurrent * this.m_dDocumentPageHeight);
		if (this.IsGoToPageMAXPosition)
		{
			if (this.SlideScrollMAX > this.m_dScrollY_max)
				this.SlideScrollMAX = this.m_dScrollY_max;

			this.m_oScrollVerApi.scrollToY(this.SlideScrollMAX);
			this.IsGoToPageMAXPosition = false;
		}
		else
		{
			//this.m_oScrollVerApi.scrollToY(this.SlideScrollMIN);
			if (this.m_dScrollY_Central > this.m_dScrollY_max)
				this.m_dScrollY_Central = this.m_dScrollY_max;

			this.m_oScrollVerApi.scrollToY(this.m_dScrollY_Central);
		}

		if (this.m_dScrollX_Central > this.m_dScrollX_max)
			this.m_dScrollX_Central = this.m_dScrollX_max;

		this.m_oScrollHorApi.scrollToX(this.m_dScrollX_Central);

		this.ZoomFreePageNum = -1;

		if (this.m_oApi.isViewMode === false && null != this.m_oLogicDocument)
		{
			//this.m_oLogicDocument.Set_CurPage( drDoc.SlideCurrent );
			//this.m_oLogicDocument.MoveCursorToXY(0, 0, false);
			this.m_oLogicDocument.RecalculateCurPos();

			this.m_oApi.sync_currentPageCallback(drDoc.SlideCurrent);
		}
		else
		{
			this.m_oApi.sync_currentPageCallback(drDoc.SlideCurrent);
		}

		this.m_oLogicDocument.Document_UpdateSelectionState();

		this.Thumbnails.LockMainObjType = false;

		if (_bIsUpdate || bIsAttack === true)
			this.OnScroll();
	};

	// fonts cache checker
	this.CheckFontCache = function()
	{
		var _c = oThis;
		_c.m_nCurrentTimeClearCache++;
		if (_c.m_nCurrentTimeClearCache > 750) // 30 секунд. корректировать при смене интервала главного таймера!!!
		{
			_c.m_nCurrentTimeClearCache = 0;
			_c.m_oDrawingDocument.CheckFontCache();
		}
        oThis.m_oLogicDocument.ContinueSpellCheck();
	};

	// open/save
	this.InitDocument = function(bIsEmpty)
	{
		this.m_oDrawingDocument.m_oWordControl   = this;
		this.m_oDrawingDocument.m_oLogicDocument = this.m_oLogicDocument;

		if (false === bIsEmpty)
		{
			this.m_oLogicDocument.LoadTestDocument();
		}

		this.CalculateDocumentSize();
		this.StartMainTimer();

		this.CreateBackgroundHorRuler();
		this.CreateBackgroundVerRuler();
		this.UpdateHorRuler();
		this.UpdateVerRuler();
	};

	this.InitControl = function()
	{
		this.Thumbnails.Init();

		this.CalculateDocumentSize();
		this.StartMainTimer();

		this.CreateBackgroundHorRuler();
		this.CreateBackgroundVerRuler();
		this.UpdateHorRuler();
		this.UpdateVerRuler();

		this.m_oApi.syncOnThumbnailsShow();

		if (true)
		{
			AscCommon.InitBrowserInputContext(this.m_oApi, "id_target_cursor", "id_main_parent");
			if (AscCommon.g_inputContext)
				AscCommon.g_inputContext.onResize("id_main_parent");

			this.initEventsMobile();
		}
	};

	this.SaveDocument = function(noBase64)
	{
		var writer = new AscCommon.CBinaryFileWriter();
		this.m_oLogicDocument.CalculateComments();
		if (noBase64) {
			return writer.WriteDocument3(this.m_oLogicDocument);;
		} else {
			var str = writer.WriteDocument(this.m_oLogicDocument);
			return str;
			//console.log(str);
		}
	};

	this.UpdateViewMode = function ()
	{
		let nMode = Asc.editor.presentationViewMode;
		switch (nMode)
		{
			case Asc.c_oAscPresentationViewMode.normal:
			{
				this.GoToPage(0);
				this.m_oApi.asc_hideComments();
				this.m_oLogicDocument.Recalculate({Drawings:{All:true, Map:{}}});
				this.m_oLogicDocument.Document_UpdateInterfaceState();
				break;
			}
			case Asc.c_oAscPresentationViewMode.sorter:
			{
				break;
			}
		}
	};
}
//------------------------------------------------------------export----------------------------------------------------
window['AscCommon']                       = window['AscCommon'] || {};
window['AscVisio']                  = window['AscVisio'] || {};
window['AscVisio'].CEditorPage      = CEditorPage;

window['AscCommon'].Page_Width      = Page_Width;
window['AscCommon'].Page_Height     = Page_Height;
window['AscCommon'].X_Left_Margin   = X_Left_Margin;
window['AscCommon'].X_Right_Margin  = X_Right_Margin;
window['AscCommon'].Y_Bottom_Margin = Y_Bottom_Margin;
window['AscCommon'].Y_Top_Margin    = Y_Top_Margin;
