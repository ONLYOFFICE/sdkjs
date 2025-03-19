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
var g_fontApplication = AscFonts.g_fontApplication;

var AscBrowser             = AscCommon.AscBrowser;
var g_anchor_left          = AscCommon.g_anchor_left;
var g_anchor_top           = AscCommon.g_anchor_top;
var g_anchor_right         = AscCommon.g_anchor_right;
var g_anchor_bottom        = AscCommon.g_anchor_bottom;
var global_keyboardEvent   = AscCommon.global_keyboardEvent;
var global_mouseEvent      = AscCommon.global_mouseEvent;
var g_dKoef_pix_to_mm      = AscCommon.g_dKoef_pix_to_mm;
var g_dKoef_mm_to_pix      = AscCommon.g_dKoef_mm_to_pix;

var Page_Width  = 210;
var Page_Height = 297;

var X_Left_Margin   = 30;  // 3   cm
var X_Right_Margin  = 15;  // 1.5 cm
var Y_Bottom_Margin = 20;  // 2   cm
var Y_Top_Margin    = 20;  // 2   cm

var X_Right_Field  = Page_Width - X_Right_Margin;
var Y_Bottom_Field = Page_Height - Y_Bottom_Margin;

var tableSpacingMinValue = 0.02;//0.02мм

function CEditorPage(api)
{
	this.Name = "";

	// size
	this.X      = 0;
	this.Y      = 0;
	this.Width  = 10;
	this.Height = 10;

	// controls
	this.m_oBody        = null;
	this.m_oMenu        = null;
	this.m_oPanelRight  = null;
	this.m_oScrollHor   = null;
	this.m_oMainContent = null;
	this.m_oLeftRuler   = null;
	this.m_oTopRuler    = null;
	this.m_oMainView    = null;
	this.m_oEditor      = null;
	this.m_oOverlay     = null;

	this.m_oPanelRight_buttonRulers   = null;
	this.m_oPanelRight_vertScroll     = null;
	this.m_oPanelRight_buttonPrevPage = null;
	this.m_oPanelRight_buttonNextPage = null;

	this.m_oLeftRuler_buttonsTabs = null;
	this.m_oLeftRuler_vertRuler   = null;
	this.m_oTopRuler_horRuler = null;

	// reader mode
	this.ReaderModeDivWrapper = null;
	this.ReaderModeDiv        = null;

	this.ReaderFontSizeCur = 2;
	this.ReaderFontSizes   = [12, 14, 16, 18, 22, 28, 36, 48, 72];

	this.ReaderTouchManager = null;
	this.ReaderModeCurrent  = 0;

	this.isNewReaderMode = true;

	// overlay
	this.m_oOverlayApi = new AscCommon.COverlay();
	this.IsUpdateOverlayOnlyEnd       = false;
	this.IsUpdateOverlayOnlyEndReturn = false;
	this.IsUpdateOverlayOnEndCheck    = false;

	// rulers
	this.m_bIsHorScrollVisible          = false;
	this.m_bIsRuler                     = (api.isMobileVersion === true) ? false : true;

	this.m_oHorRuler = new CHorRuler();
	this.m_oVerRuler = new CVerRuler();
	this.m_oHorRuler.m_oWordControl = this;
	this.m_oVerRuler.m_oWordControl = this

	this.m_bIsUpdateHorRuler = false;
	this.m_bIsUpdateVerRuler = false;

	// document position
	this.m_oBoundsController = new AscFormat.CBoundsController();

	this.m_dScrollY           = 0;
	this.m_dScrollX           = 0;
	this.m_dScrollY_max       = 1;
	this.m_dScrollX_max       = 1;

	this.m_dDocumentWidth      = 0;
	this.m_dDocumentHeight     = 0;
	this.m_dDocumentPageWidth  = 0;
	this.m_dDocumentPageHeight = 0;

	this.m_oScrollHor_ = null;
	this.m_oScrollVer_ = null;

	this.m_oScrollHorApi = null;
	this.m_oScrollVerApi = null;

	this.zoom_values = [50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 320, 340, 360, 380, 400, 425, 450, 475, 500];
	this.m_nZoomType = 0; // 0 - custom, 1 - fitToWidth, 2 - fitToPage
	this.m_nZoomValue = 100;

	// текущий зум после резайза. чтобы например после zoomToWidth и zoomIn/Out можно было вернуться на значение меньше меньшего/больше большего
	this.m_nZoomValueMin = -1;
	this.m_nZoomValueMax = -1;

	// mobile
	this.MobileTouchManager = null;

	// сдвиг для редактора. используется для мобильной версии. меняется при скрытии тулбара.
	// изначально - высота тулбара. без учета deviceScale (css значение)
	this.offsetTop = 0;
	// запоминаем позицию скролла при начатии скролла. чтобы отсылать изменение в интерфейс (ТОЛЬКО при скролле).
	// чтобы иметь возможность убирать интерфейс
	this.mobileScrollStartPos = 0;

	// враппер для текстового элемента
	this.TextBoxBackground = null;

	//
	this.m_bIsIE = (AscBrowser.isIE || window.opera) ? true : false;

	this.m_bDocumentPlaceChangedEnabled = false;
	this.m_nTabsType = tab_Left;

	this.NoneRepaintPages = false;
	this.m_bIsScroll = false;
	this.m_bIsRePaintOnScroll = true;
	this.m_bIsFullRepaint = false;

	this.ScrollsWidthPx = 14;

	this.m_oDrawingDocument = Asc.editor.isPdfEditor() ? new AscPDF.CDrawingDocument() : new AscCommon.CDrawingDocument();
	this.m_oLogicDocument   = null;

	this.m_oDrawingDocument.m_oWordControl   = this;
	this.m_oDrawingDocument.m_oLogicDocument = this.m_oLogicDocument;

	this.m_bIsUpdateTargetNoAttack = false;

	this.arrayEventHandlers = [];

	this.m_oTimerScrollSelect = -1;
	this.IsFocus              = true;
	this.m_bIsMouseLock       = false;

	this.DrawingFreeze      = false;

	this.IsKeyDownButNoPress = false;

	this.MouseDownDocumentCounter = 0;
	this.MouseHandObject = null;

	this.bIsUseKeyPress = true;
	this.bIsEventPaste  = false;
	this.bIsDoublePx    = AscCommon.isSupportDoublePx();

	this.m_nCurrentTimeClearCache = 0;

	this.m_bIsMouseUpSend = false;

    this.retinaScaling = AscCommon.AscBrowser.retinaPixelRatio;
	this.IsRepaintOnCallbackLongAction = false;

	this.IsInitControl = false;

	// paint loop
	this.paintMessageLoop = new AscCommon.PaintMessageLoop(40, api);

	this.m_oApi = api;
	var oThis   = this;

	// methods ---
	this.checkBodyOffset = function()
	{
		var off = jQuery("#" + this.Name).offset();

		if (off)
		{
			this.X = off.left;
			this.Y = off.top;

			if (AscCommon.AscBrowser.isOffsetUsedZoom())
			{
				this.X *= AscCommon.AscBrowser.zoom;
				this.Y *= AscCommon.AscBrowser.zoom;
			}
		}
	};

	this.checkBodySize = function()
	{
		this.checkBodyOffset();

		var el = document.getElementById(this.Name);

		var _newW = el.offsetWidth;
		var _newH = el.offsetHeight;

		var _left_border_w = 0;
		if (window.getComputedStyle)
		{
			var _computed_style = window.getComputedStyle(el, null);
			if (_computed_style)
			{
				var _value = _computed_style.getPropertyValue("border-left-width");

				if (typeof _value == "string")
				{
					_left_border_w = parseInt(_value);
				}
			}
		}

		_newW -= _left_border_w;
		if (this.Width !== _newW || this.Height !== _newH)
		{
			this.Width  = _newW;
			this.Height = _newH;
			return true;
		}
		return false;
	};

	this.Init = function()
	{
		this.m_oBody = AscCommon.CreateControlContainer(this.Name);

		var scrollWidthMm = this.ScrollsWidthPx * g_dKoef_pix_to_mm;

		this.m_oScrollHor = AscCommon.CreateControlContainer("id_horscrollpanel");
		this.m_oScrollHor.Bounds.SetParams(0, 0, scrollWidthMm, 0, false, false, true, true, -1, scrollWidthMm);
		this.m_oScrollHor.Anchor = (g_anchor_left | g_anchor_right | g_anchor_bottom);
		this.m_oBody.AddControl(this.m_oScrollHor);

		// panel right --------------------------------------------------------------
		this.m_oPanelRight = AscCommon.CreateControlContainer("id_panel_right");
		if (!this.m_oApi.isRtlInterface)
		{
			this.m_oPanelRight.Bounds.SetParams(0, 0, 1000, 0, false, true, false, true, scrollWidthMm, -1);
			this.m_oPanelRight.Anchor = (g_anchor_top | g_anchor_right | g_anchor_bottom);
		}
		else
		{
			this.m_oPanelRight.Bounds.SetParams(0, 0, 0, 0, false, true, false, true, scrollWidthMm, -1);
			this.m_oPanelRight.Anchor = (g_anchor_top | g_anchor_left | g_anchor_bottom);
		}

		this.m_oBody.AddControl(this.m_oPanelRight);
		if (this.m_oApi.isMobileVersion)
		{
			this.m_oPanelRight.HtmlElement.style.zIndex = -1;

			var hor_scroll          = document.getElementById('id_horscrollpanel');
			hor_scroll.style.zIndex = -1;
		}

		this.m_oPanelRight_buttonRulers = AscCommon.CreateControl("id_buttonRulers");
		this.m_oPanelRight_buttonRulers.Bounds.SetParams(0, 0, 1000, 1000, false, false, false, false, -1, scrollWidthMm);
		this.m_oPanelRight_buttonRulers.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right);
		this.m_oPanelRight.AddControl(this.m_oPanelRight_buttonRulers);

		var _vertScrollTop = scrollWidthMm;
		if (GlobalSkin.RulersButton === false)
		{
			this.m_oPanelRight_buttonRulers.HtmlElement.style.display = "none";
			_vertScrollTop                                            = 0;
		}

		this.m_oPanelRight_buttonNextPage = AscCommon.CreateControl("id_buttonNextPage");
		this.m_oPanelRight_buttonNextPage.Bounds.SetParams(0, 0, 1000, 1000, false, false, false, false, -1, scrollWidthMm);
		this.m_oPanelRight_buttonNextPage.Anchor = (g_anchor_left | g_anchor_bottom | g_anchor_right);
		this.m_oPanelRight.AddControl(this.m_oPanelRight_buttonNextPage);

		this.m_oPanelRight_buttonPrevPage = AscCommon.CreateControl("id_buttonPrevPage");
		this.m_oPanelRight_buttonPrevPage.Bounds.SetParams(0, 0, 1000, scrollWidthMm, false, false, false, true, -1, scrollWidthMm);
		this.m_oPanelRight_buttonPrevPage.Anchor = (g_anchor_left | g_anchor_bottom | g_anchor_right);
		this.m_oPanelRight.AddControl(this.m_oPanelRight_buttonPrevPage);

		var _vertScrollBottom = 2 * scrollWidthMm;
		if (GlobalSkin.NavigationButtons == false)
		{
			this.m_oPanelRight_buttonNextPage.HtmlElement.style.display = "none";
			this.m_oPanelRight_buttonPrevPage.HtmlElement.style.display = "none";
			_vertScrollBottom                                           = 0;
		}

		this.m_oPanelRight_vertScroll = AscCommon.CreateControl("id_vertical_scroll");
		this.m_oPanelRight_vertScroll.Bounds.SetParams(0, _vertScrollTop, 1000, _vertScrollBottom, false, true, false, true, -1, -1);
		this.m_oPanelRight_vertScroll.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right | g_anchor_bottom);
		this.m_oPanelRight.AddControl(this.m_oPanelRight_vertScroll);
		// --------------------------------------------------------------------------

		// main content -------------------------------------------------------------
		this.m_oMainContent = AscCommon.CreateControlContainer("id_main");
		if (!this.m_oApi.isMobileVersion)
		{
			if (!this.m_oApi.isRtlInterface)
				this.m_oMainContent.Bounds.SetParams(0, 0, scrollWidthMm, 0, false, true, true, true, -1, -1);
			else
				this.m_oMainContent.Bounds.SetParams(scrollWidthMm, 0, 0, 0, true, true, true, true, -1, -1);
		}
		else
		{
			this.m_oMainContent.Bounds.SetParams(0, 0, 0, 0, false, true, true, true, -1, -1);
		}
		this.m_oMainContent.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right | g_anchor_bottom);
		this.m_oBody.AddControl(this.m_oMainContent);

		// --- left ---
		if (!this.m_oApi.isRtlInterface)
		{
			this.m_oLeftRuler = AscCommon.CreateControlContainer("id_panel_left");
			this.m_oLeftRuler.Bounds.SetParams(0, 0, 1000, 1000, false, false, false, false, 5, -1);
			this.m_oLeftRuler.Anchor = (g_anchor_left | g_anchor_top | g_anchor_bottom);
			this.m_oMainContent.AddControl(this.m_oLeftRuler);

			this.m_oMainView = AscCommon.CreateControlContainer("id_main_view");
			this.m_oMainView.Bounds.SetParams(5, 7, 1000, 1000, true, true, false, false, -1, -1);
			this.m_oMainView.Anchor = (g_anchor_left | g_anchor_right | g_anchor_top | g_anchor_bottom);
			this.m_oMainContent.AddControl(this.m_oMainView);
		}
		else
		{
			this.m_oLeftRuler = AscCommon.CreateControlContainer("id_panel_left");
			this.m_oLeftRuler.Bounds.SetParams(0, 0, 1000, 1000, false, false, false, false, 5, -1);
			this.m_oLeftRuler.Anchor = (g_anchor_right | g_anchor_top | g_anchor_bottom);
			this.m_oMainContent.AddControl(this.m_oLeftRuler);

			this.m_oMainView = AscCommon.CreateControlContainer("id_main_view");
			this.m_oMainView.Bounds.SetParams(0, 7, 5, 1000, false, true, true, false, -1, -1);
			this.m_oMainView.Anchor = (g_anchor_left | g_anchor_right | g_anchor_top | g_anchor_bottom);
			this.m_oMainContent.AddControl(this.m_oMainView);
		}
		this.m_oLeftRuler_buttonsTabs = AscCommon.CreateControl("id_buttonTabs");
		this.m_oLeftRuler_buttonsTabs.Bounds.SetParams(0, 0.8, 1000, 1000, false, true, false, false, -1, 5);
		this.m_oLeftRuler_buttonsTabs.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right);
		this.m_oLeftRuler.AddControl(this.m_oLeftRuler_buttonsTabs);

		this.m_oLeftRuler_vertRuler = AscCommon.CreateControl("id_vert_ruler");
		this.m_oLeftRuler_vertRuler.Bounds.SetParams(0, 7, 1000, 1000, false, true, false, false, -1, -1);
		this.m_oLeftRuler_vertRuler.Anchor = (g_anchor_left | g_anchor_right | g_anchor_top | g_anchor_bottom);
		this.m_oLeftRuler.AddControl(this.m_oLeftRuler_vertRuler);
		// ------------

		// --- top ----
		if (!this.m_oApi.isRtlInterface)
		{
			this.m_oTopRuler = AscCommon.CreateControlContainer("id_panel_top");
			this.m_oTopRuler.Bounds.SetParams(5, 0, 1000, 1000, true, false, false, false, -1, 7);
			this.m_oTopRuler.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right);
			this.m_oMainContent.AddControl(this.m_oTopRuler);
		}
		else
		{
			this.m_oTopRuler = AscCommon.CreateControlContainer("id_panel_top");
			this.m_oTopRuler.Bounds.SetParams(0, 0, 5, 1000, false, false, true, false, -1, 7);
			this.m_oTopRuler.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right);
			this.m_oMainContent.AddControl(this.m_oTopRuler);
		}

		this.m_oTopRuler_horRuler = AscCommon.CreateControl("id_hor_ruler");
		this.m_oTopRuler_horRuler.Bounds.SetParams(0, 0, 1000, 1000, false, false, false, false, -1, -1);
		this.m_oTopRuler_horRuler.Anchor = (g_anchor_left | g_anchor_right | g_anchor_top | g_anchor_bottom);
		this.m_oTopRuler.AddControl(this.m_oTopRuler_horRuler);
		// ------------

		// проблема с фокусом fixed-позиционированного элемента внутри (bug 63194)
		this.m_oMainView.HtmlElement.onscroll = function() {
			this.scrollTop = 0;
		};

		this.m_oEditor = AscCommon.CreateControl("id_viewer");
		this.m_oEditor.Bounds.SetParams(0, 0, 1000, 1000, false, false, false, false, -1, -1);
		this.m_oEditor.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right | g_anchor_bottom);
		this.m_oMainView.AddControl(this.m_oEditor);

		this.m_oOverlay = AscCommon.CreateControl("id_viewer_overlay");
		this.m_oOverlay.Bounds.SetParams(0, 0, 1000, 1000, false, false, false, false, -1, -1);
		this.m_oOverlay.Anchor = (g_anchor_left | g_anchor_top | g_anchor_right | g_anchor_bottom);
		this.m_oMainView.AddControl(this.m_oOverlay);
		// --------------------------------------------------------------------------

		this.m_oDrawingDocument.TargetHtmlElement = document.getElementById('id_target_cursor');

		this.checkNeedRules();
		this.initEvents();

		this.m_oOverlayApi.m_oControl  = this.m_oOverlay;
		this.m_oOverlayApi.m_oHtmlPage = this;
		this.m_oOverlayApi.Clear();
		this.ShowOverlay();

		this.m_oDrawingDocument.AutoShapesTrack = new AscCommon.CAutoshapeTrack();
		this.m_oDrawingDocument.AutoShapesTrack.init2(this.m_oOverlayApi);

		this.OnResize(true);

		// в мобильной версии - при транзишне - не обновляется позиция/размер
		if (this.m_oApi.isMobileVersion)
		{
			var _t = this;
			document.addEventListener && document.addEventListener("transitionend", function() { _t.OnResize(false);  }, false);
			document.addEventListener && document.addEventListener("transitioncancel", function() { _t.OnResize(false); }, false);
		}

		this.checkMouseHandMode();
	};

    this.CheckRetinaDisplay = function()
    {
        if (this.retinaScaling !== AscCommon.AscBrowser.retinaPixelRatio)
        {
            this.retinaScaling = AscCommon.AscBrowser.retinaPixelRatio;
            // сбросить кэш страниц
            this.onButtonTabsDraw();
        }
    };

	this.ShowOverlay        = function()
	{
		this.m_oOverlay.HtmlElement.style.display = "block";

		if (null == this.m_oOverlayApi.m_oContext)
			this.m_oOverlayApi.m_oContext = this.m_oOverlayApi.m_oControl.HtmlElement.getContext('2d');
	};
	this.UnShowOverlay      = function()
	{
		this.m_oOverlay.HtmlElement.style.display = "none";
	};
	this.CheckUnShowOverlay = function()
	{
		var drDoc = this.m_oDrawingDocument;
		if (!drDoc.m_bIsSearching && !drDoc.m_bIsSelection && !this.MobileTouchManager)
		{
			this.UnShowOverlay();
			return false;
		}
		return true;
	};
	this.CheckShowOverlay = function()
	{
		var drDoc = this.m_oDrawingDocument;
		if (drDoc.m_bIsSearching || drDoc.m_bIsSelection || this.MobileTouchManager)
			this.ShowOverlay();
	};

	// events ---
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
		{
			//this.m_oMainContent.HtmlElement.addEventListener("DOMMouseScroll", new Function("event", "return Editor_OnMouseWhell(event);"), false);
			this.m_oMainContent.HtmlElement.addEventListener("DOMMouseScroll", this.onMouseWhell, false);
		}

        AscCommon.addMouseEvent(this.m_oTopRuler_horRuler.HtmlElement, "down", this.horRulerMouseDown);
        AscCommon.addMouseEvent(this.m_oTopRuler_horRuler.HtmlElement, "move", this.horRulerMouseMove);
        AscCommon.addMouseEvent(this.m_oTopRuler_horRuler.HtmlElement, "up", this.horRulerMouseUp);

        AscCommon.addMouseEvent(this.m_oLeftRuler_vertRuler.HtmlElement, "down", this.verRulerMouseDown);
        AscCommon.addMouseEvent(this.m_oLeftRuler_vertRuler.HtmlElement, "move", this.verRulerMouseMove);
        AscCommon.addMouseEvent(this.m_oLeftRuler_vertRuler.HtmlElement, "up", this.verRulerMouseUp);

		this.m_oBody.HtmlElement.oncontextmenu = function(e)
		{
			if (AscCommon.AscBrowser.isVivaldiLinux)
			{
				AscCommon.Window_OnMouseUp(e);
			}
			AscCommon.stopEvent(e);
			return false;
		};
	};

	this.initEventsMobile = function()
	{
		if (this.m_oApi.isUseOldMobileVersion())
		{
			this.MobileTouchManager = new AscCommon.CMobileTouchManager( { eventsElement : "word_mobile_element" } );
			this.MobileTouchManager.Init(this.m_oApi);
			if (!this.MobileTouchManager.delegate.IsNativeViewer())
				this.MobileTouchManager.Resize();

			this.TextBoxBackground = AscCommon.CreateControl(AscCommon.g_inputContext.HtmlArea.id);
			this.TextBoxBackground.HtmlElement.parentNode.parentNode.style.zIndex = 10;

			this.MobileTouchManager.initEvents(AscCommon.g_inputContext.HtmlArea.id);

			if (AscBrowser.isAndroid)
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
			this.MobileTouchManager = new AscCommon.CMobileTouchManager( { eventsElement : "word_mobile_element", desktopMode : true } );
			this.MobileTouchManager.Init(this.m_oApi);
			this.MobileTouchManager.Resize();

			this.MobileTouchManager.addClickElement([this.m_oEditor.HtmlElement, this.m_oOverlay.HtmlElement]);
		}
	};

	// mouse hand mode ---
	this.checkMouseHandMode = function()
	{
		if (!this.m_oApi || !this.m_oApi.isTargetHandMode())
		{
			this.MouseHandObject = null;
			return;
		}

		this.MouseHandObject = {
			check : function(_this, _pos) {
				var logicDoc = _this.m_oLogicDocument;
				if (!logicDoc)
					return true;

				var isForms = (logicDoc.IsInForm(_pos.X, _pos.Y, _pos.Page) || logicDoc.IsInContentControl(_pos.X, _pos.Y, _pos.Page)) ? true : false;
				var isButtons = _this.m_oDrawingDocument.contentControls.checkPointerInButtons(_pos);

				if (isForms || isButtons)
					return false;

				return true;
			}
		};
	};

	// rulers ---
	this.onButtonRulersClick = function()
	{
		if (false === oThis.m_oApi.bInit_word_control || true === oThis.m_oApi.isViewMode)
			return;

		oThis.m_bIsRuler = !oThis.m_bIsRuler;
		oThis.checkNeedRules();
		oThis.OnResize(true);
	};

	this.HideRulers = function()
	{
		if (oThis.m_bIsRuler === false)
			return;

		oThis.m_bIsRuler = !oThis.m_bIsRuler;
		oThis.checkNeedRules();
		oThis.OnResize(true);
	};

	// rulers (events) ---
	this.horRulerMouseDown = function(e)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		if (e.preventDefault)
			e.preventDefault();
		else
			e.returnValue = false;

		var oWordControl = oThis;

		var _cur_page = oWordControl.m_oDrawingDocument.m_lCurrentPage;
		if (_cur_page < 0 || _cur_page >= oWordControl.m_oDrawingDocument.m_lPagesCount)
			return;

		oWordControl.m_oHorRuler.OnMouseDown(oWordControl.m_oDrawingDocument.m_arrPages[_cur_page].drawingPage.left, 0, e);
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

		var _cur_page = oWordControl.m_oDrawingDocument.m_lCurrentPage;
		if (_cur_page < 0 || _cur_page >= oWordControl.m_oDrawingDocument.m_lPagesCount)
			return;

		oWordControl.m_oHorRuler.OnMouseUp(oWordControl.m_oDrawingDocument.m_arrPages[_cur_page].drawingPage.left, 0, e);
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

		var _cur_page = oWordControl.m_oDrawingDocument.m_lCurrentPage;
		if (_cur_page < 0 || _cur_page >= oWordControl.m_oDrawingDocument.m_lPagesCount)
			return;

		oWordControl.m_oHorRuler.OnMouseMove(oWordControl.m_oDrawingDocument.m_arrPages[_cur_page].drawingPage.left, 0, e);
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

		var _cur_page = oWordControl.m_oDrawingDocument.m_lCurrentPage;
		if (_cur_page < 0 || _cur_page >= oWordControl.m_oDrawingDocument.m_lPagesCount)
			return;

		oWordControl.m_oVerRuler.OnMouseDown(0, oWordControl.m_oDrawingDocument.m_arrPages[_cur_page].drawingPage.top, e);
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

		var _cur_page = oWordControl.m_oDrawingDocument.m_lCurrentPage;
		if (_cur_page < 0 || _cur_page >= oWordControl.m_oDrawingDocument.m_lPagesCount)
			return;

		oWordControl.m_oVerRuler.OnMouseUp(0, oWordControl.m_oDrawingDocument.m_arrPages[_cur_page].drawingPage.top, e);
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

		var _cur_page = oWordControl.m_oDrawingDocument.m_lCurrentPage;
		if (_cur_page < 0 || _cur_page >= oWordControl.m_oDrawingDocument.m_lPagesCount)
			return;

		oWordControl.m_oVerRuler.OnMouseMove(0, oWordControl.m_oDrawingDocument.m_arrPages[_cur_page].drawingPage.top, e);
	};

	this.checkNeedRules     = function()
	{
		if (this.m_bIsRuler)
		{
			this.m_oLeftRuler.HtmlElement.style.display = 'block';
			this.m_oTopRuler.HtmlElement.style.display  = 'block';

			if (!this.m_oApi.isRtlInterface)
				this.m_oMainView.Bounds.L = 5;
			else
				this.m_oMainView.Bounds.R = 5;

			this.m_oMainView.Bounds.T = 7;
		}
		else
		{
			this.m_oLeftRuler.HtmlElement.style.display = 'none';
			this.m_oTopRuler.HtmlElement.style.display  = 'none';

			if (!this.m_oApi.isRtlInterface)
				this.m_oMainView.Bounds.L = 0;
			else
				this.m_oMainView.Bounds.R = 0;

			this.m_oMainView.Bounds.T = 0;
		}
	};

	this.UpdateHorRuler = function()
	{
		if (!this.m_bIsRuler || this.m_oApi.isDocumentRenderer())
			return;

		var _left = 0;
		var lPage = this.m_oDrawingDocument.m_lCurrentPage;
		if (0 <= lPage && lPage < this.m_oDrawingDocument.m_lPagesCount)
		{
			_left = this.m_oDrawingDocument.m_arrPages[lPage].drawingPage.left;
		}
		else if (this.m_oDrawingDocument.m_lPagesCount != 0)
		{
			_left = this.m_oDrawingDocument.m_arrPages[this.m_oDrawingDocument.m_lPagesCount - 1].drawingPage.left;
		}

		this.m_oHorRuler.BlitToMain(_left, 0, this.m_oTopRuler_horRuler.HtmlElement);
	};
	this.UpdateVerRuler = function()
	{
		if (!this.m_bIsRuler || this.m_oApi.isDocumentRenderer())
			return;

		var _top  = 0;
		var lPage = this.m_oDrawingDocument.m_lCurrentPage;
		if (0 <= lPage && lPage < this.m_oDrawingDocument.m_lPagesCount)
		{
			_top = this.m_oDrawingDocument.m_arrPages[lPage].drawingPage.top;
		}
		else if (this.m_oDrawingDocument.m_lPagesCount != 0)
		{
			_top = this.m_oDrawingDocument.m_arrPages[this.m_oDrawingDocument.m_lPagesCount - 1].drawingPage.top;
		}

		this.m_oVerRuler.BlitToMain(0, _top, this.m_oLeftRuler_vertRuler.HtmlElement);
	};

	this.UpdateHorRulerBack = function(isattack)
	{
		var drDoc = this.m_oDrawingDocument;
		if (0 <= drDoc.m_lCurrentPage && drDoc.m_lCurrentPage < drDoc.m_lPagesCount)
		{
			this.m_oHorRuler.CreateBackground(drDoc.m_arrPages[drDoc.m_lCurrentPage], isattack);
		}
		this.UpdateHorRuler();
	};
	this.UpdateVerRulerBack = function(isattack)
	{
		var drDoc = this.m_oDrawingDocument;
		if (0 <= drDoc.m_lCurrentPage && drDoc.m_lCurrentPage < drDoc.m_lPagesCount)
		{
			this.m_oVerRuler.CreateBackground(drDoc.m_arrPages[drDoc.m_lCurrentPage], isattack);
		}
		this.UpdateVerRuler();
	};

	// zoom ---
	this.calculate_zoom_FitToWidth = function()
	{
		var w = this.m_oEditor.AbsolutePosition.R - this.m_oEditor.AbsolutePosition.L;

		var Zoom = 100;
		if (0 != this.m_dDocumentPageWidth)
		{
			Zoom = 100 * (w - 10) / this.m_dDocumentPageWidth;

			if (Zoom < 5)
				Zoom = 5;

			if (this.m_oApi.isMobileVersion)
			{
				var _w = this.m_oEditor.HtmlElement.width;
				_w /= AscCommon.AscBrowser.retinaPixelRatio;
				Zoom = 100 * _w * g_dKoef_pix_to_mm / this.m_dDocumentPageWidth;
			}
		}
		return (Zoom - 0.5) >> 0;
	};
	this.zoom_FitToWidth = function()
	{
		if (this.m_oApi.isUseNativeViewer && this.m_oDrawingDocument.m_oDocumentRenderer)
		{
			this.m_oDrawingDocument.m_oDocumentRenderer.setZoomMode(AscCommon.ViewerZoomMode.Width);
			return;
		}

		var _new_value = this.calculate_zoom_FitToWidth();

		this.m_nZoomType = 1;
		if (_new_value != this.m_nZoomValue)
		{
			var _old_val      = this.m_nZoomValue;
			this.m_nZoomValue = _new_value;
			this.zoom_Fire(1, _old_val);

			if (this.MobileTouchManager)
				this.MobileTouchManager.CheckZoomCriticalValues(this.m_nZoomValue);

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
		if (this.m_oApi.isUseNativeViewer && this.m_oDrawingDocument.m_oDocumentRenderer)
		{
			this.m_oDrawingDocument.m_oDocumentRenderer.setZoomMode(AscCommon.ViewerZoomMode.Page);
			return;
		}

		var w = parseInt(this.m_oEditor.HtmlElement.width) * g_dKoef_pix_to_mm;
		var h = parseInt(this.m_oEditor.HtmlElement.height) * g_dKoef_pix_to_mm;

		w = AscCommon.AscBrowser.convertToRetinaValue(w);
		h = AscCommon.AscBrowser.convertToRetinaValue(h);

		var _hor_Zoom = 100;
		if (0 != this.m_dDocumentPageWidth)
			_hor_Zoom = (100 * (w - 10)) / this.m_dDocumentPageWidth;
		var _ver_Zoom = 100;
		if (0 != this.m_dDocumentPageHeight)
			_ver_Zoom = (100 * (h - 10)) / this.m_dDocumentPageHeight;

		var _new_value = parseInt(Math.min(_hor_Zoom, _ver_Zoom) - 0.5);

		if (_new_value < 5)
			_new_value = 5;

		this.m_nZoomType = 2;
		if (_new_value != this.m_nZoomValue)
		{
			var _old_val      = this.m_nZoomValue;
			this.m_nZoomValue = _new_value;
			this.zoom_Fire(2, _old_val);
			return true;
		}
		else
		{
			this.m_oApi.sync_zoomChangeCallback(this.m_nZoomValue, 2);
		}
		return false;
	};

	this.zoom_Fire = function(type, old_zoom)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		if (0 === type)
		{
			if (oThis.m_oApi.isUseNativeViewer && oThis.m_oDrawingDocument.m_oDocumentRenderer)
			{
				oThis.m_oDrawingDocument.m_oDocumentRenderer.setZoom(oThis.m_nZoomValue / 100);
				return;
			}
		}

		// нужно проверить режим и сбросить кеш грамотно (ie version)
		AscCommon.g_fontManager.ClearRasterMemory();

		if (AscCommon.g_fontManager2)
			AscCommon.g_fontManager2.ClearRasterMemory();

		var oWordControl = oThis;

		oWordControl.m_bIsRePaintOnScroll = false;

		var xScreen1 = oWordControl.m_oEditor.AbsolutePosition.R - oWordControl.m_oEditor.AbsolutePosition.L;
		var yScreen1 = oWordControl.m_oEditor.AbsolutePosition.B - oWordControl.m_oEditor.AbsolutePosition.T;
		xScreen1 *= g_dKoef_mm_to_pix;
		yScreen1 *= g_dKoef_mm_to_pix;

		xScreen1 >>= 1;
		yScreen1 >>= 1;

		var posDoc = oWordControl.m_oDrawingDocument.ConvertCoordsFromCursor2(xScreen1, yScreen1, old_zoom, true);

		oWordControl.CheckZoom();
		oWordControl.CalculateDocumentSize();

		var lCurPage = oWordControl.m_oDrawingDocument.m_lCurrentPage;
		if (-1 != lCurPage)
		{
			oWordControl.m_oHorRuler.CreateBackground(oWordControl.m_oDrawingDocument.m_arrPages[lCurPage]);
			oWordControl.m_bIsUpdateHorRuler = true;
			oWordControl.m_oVerRuler.CreateBackground(oWordControl.m_oDrawingDocument.m_arrPages[lCurPage]);
			oWordControl.m_bIsUpdateVerRuler = true;
		}

		oWordControl.OnCalculatePagesPlace();
		var posScreenNew = oWordControl.m_oDrawingDocument.ConvertCoordsToCursor(posDoc.X, posDoc.Y, posDoc.Page);

		var _x_pos = oWordControl.m_oScrollHorApi.getCurScrolledX() + posScreenNew.X - xScreen1;
		var _y_pos = oWordControl.m_oScrollVerApi.getCurScrolledY() + posScreenNew.Y - yScreen1;

		_x_pos = Math.max(0, Math.min(_x_pos, oWordControl.m_dScrollX_max));
		_y_pos = Math.max(0, Math.min(_y_pos, oWordControl.m_dScrollY_max));

		// TODO: заглушка под открытие. мы любим открывать файлы с зумом. И тогда документ не в начале открывается, а с
		// малениким проскролливанием
		if (oWordControl.m_dScrollY == 0)
			_y_pos = 0;

		oWordControl.m_oScrollVerApi.scrollToY(_y_pos);
		oWordControl.m_oScrollHorApi.scrollToX(_x_pos);

		if (this.MobileTouchManager)
			this.MobileTouchManager.Resize();

		oWordControl.m_oApi.sync_zoomChangeCallback(this.m_nZoomValue, type);
		oWordControl.m_bIsUpdateTargetNoAttack = true;
		oWordControl.m_bIsRePaintOnScroll      = true;


		oWordControl.OnScroll();

		if (this.MobileTouchManager)
			this.MobileTouchManager.Resize_After();

		if (this.m_oApi.watermarkDraw)
		{
			this.m_oApi.watermarkDraw.zoom = this.m_nZoomValue / 100;
			this.m_oApi.watermarkDraw.Generate();
		}
	};

	this.zoom_Out = function()
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		oThis.m_nZoomType = 0;

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

		if (_Zoom >= oThis.m_nZoomValue && (oThis.m_nZoomValueMin > 0 && _Zoom > oThis.m_nZoomValueMin))
			_Zoom = oThis.m_nZoomValueMin;

		if (oThis.m_nZoomValue <= _Zoom)
			return;

		var _old_val       = oThis.m_nZoomValue;
		oThis.m_nZoomValue = _Zoom;
		oThis.zoom_Fire(0, _old_val);
	};
	this.zoom_In = function()
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		oThis.m_nZoomType = 0;

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

		if (_Zoom <= oThis.m_nZoomValue && (oThis.m_nZoomValueMax > 0 && _Zoom < oThis.m_nZoomValueMax))
			_Zoom = oThis.m_nZoomValueMax;

		if (oThis.m_nZoomValue >= _Zoom)
			return;

		var _old_val       = oThis.m_nZoomValue;
		oThis.m_nZoomValue = _Zoom;
		oThis.zoom_Fire(0, _old_val);
	};

	this.CheckZoom = function()
	{
		if (!this.NoneRepaintPages)
			this.m_oDrawingDocument.ClearCachePages();
	};

	// reader ---
	this.ChangeReaderMode = function()
	{
		if (!this.m_oLogicDocument)
			return;

		if (this.ReaderModeCurrent)
			this.DisableReaderMode();
		else
			this.EnableReaderMode();
	};

	this.SetNewMobileMode = function()
	{
		if (this.m_oLogicDocument)
		{
			if (this.m_oDrawingDocument)
			{
				this.m_nZoomType = 1;
				this.m_oDrawingDocument.m_bUpdateAllPagesOnFirstRecalculate = true;
			}
			let sectPr = this.m_oLogicDocument.GetSectionsInfo().Get(0).SectPr;
			const nPageW = sectPr.GetPageWidth() / AscCommon.AscBrowser.retinaPixelRatio;
			const nPageH = sectPr.GetPageHeight() / AscCommon.AscBrowser.retinaPixelRatio;
			const nScale = this.ReaderFontSizes[this.ReaderFontSizeCur] / 16;
			this.m_oLogicDocument.SetDocumentReadMode(nPageW, nPageH, nScale);
			return true;
		}
		return false;
	};

	this.IncreaseReaderFontSize = function()
	{
		if (this.ReaderFontSizeCur >= (this.ReaderFontSizes.length - 1))
		{
			this.ReaderFontSizeCur = this.ReaderFontSizes.length - 1;
			return;
		}
		this.ReaderFontSizeCur++;

		if (this.isNewReaderMode && this.SetNewMobileMode())
			return;

		if (null == this.ReaderModeDiv)
			return;

		this.ReaderModeDiv.style.fontSize = this.ReaderFontSizes[this.ReaderFontSizeCur] + "pt";
		this.ReaderTouchManager.ChangeFontSize();
	};
	this.DecreaseReaderFontSize = function()
	{
		if (this.ReaderFontSizeCur <= 0)
		{
			this.ReaderFontSizeCur = 0;
			return;
		}
		this.ReaderFontSizeCur--;

		if (this.isNewReaderMode && this.SetNewMobileMode())
			return;

		if (null == this.ReaderModeDiv)
			return;

		this.ReaderModeDiv.style.fontSize = this.ReaderFontSizes[this.ReaderFontSizeCur] + "pt";
		this.ReaderTouchManager.ChangeFontSize();
	};

	this.IsReaderMode        = function()
	{
		return (this.ReaderModeCurrent == 1);
	};
	this.UpdateReaderContent = function()
	{
		if (this.ReaderModeCurrent == 1 && this.ReaderModeDivWrapper != null)
		{
			this.ReaderModeDivWrapper.innerHTML = "<div id=\"reader_id\" style=\"width:100%;display:block;z-index:9;font-family:arial;font-size:" +
				this.ReaderFontSizes[this.ReaderFontSizeCur] + "pt;position:absolute;resize:none;-webkit-box-sizing:border-box;box-sizing:border-box;padding-left:5%;padding-right:5%;padding-top:10%;padding-bottom:10%;background-color:#FFFFFF;\">" +
				this.m_oApi.ContentToHTML(true) + "</div>";
		}
	};

	this.EnableReaderMode = function()
	{
		if (this.isNewReaderMode && this.SetNewMobileMode())
		{
			this.ReaderModeCurrent = 1;
			return;
		}

		this.ReaderModeCurrent = 1;
		if (this.ReaderTouchManager)
		{
			this.TransformDivUseAnimation(this.ReaderModeDivWrapper, 0);
			return;
		}

		this.ReaderModeDivWrapper = document.createElement('div');
		this.ReaderModeDivWrapper.setAttribute("style", "z-index:11;font-family:arial;font-size:12pt;position:absolute;\
            resize:none;padding:0px;display:block;\
            margin:0px;left:0px;top:0px;background-color:#FFFFFF");

		var _c_h                               = parseInt(oThis.m_oMainView.HtmlElement.style.height);
		this.ReaderModeDivWrapper.style.top    = _c_h + "px";
		this.ReaderModeDivWrapper.style.width  = this.m_oMainView.HtmlElement.style.width;
		this.ReaderModeDivWrapper.style.height = this.m_oMainView.HtmlElement.style.height;

		this.ReaderModeDivWrapper.id        = "wrapper_reader_id";
		this.ReaderModeDivWrapper.innerHTML = "<div id=\"reader_id\" style=\"width:100%;display:block;z-index:9;font-family:arial;font-size:" +
			this.ReaderFontSizes[this.ReaderFontSizeCur] + "pt;position:absolute;resize:none;-webkit-box-sizing:border-box;box-sizing:border-box;padding-left:5%;padding-right:5%;padding-top:10%;padding-bottom:10%;background-color:#FFFFFF;\">" +
			this.m_oApi.ContentToHTML(true) + "</div>";

		this.m_oMainView.HtmlElement.appendChild(this.ReaderModeDivWrapper);

		this.ReaderModeDiv = document.getElementById("reader_id");

		if (this.MobileTouchManager)
		{
			this.MobileTouchManager.Destroy();
			this.MobileTouchManager = null;
		}

		this.ReaderTouchManager = new AscCommon.CReaderTouchManager();
		this.ReaderTouchManager.Init(this.m_oApi);

		this.TransformDivUseAnimation(this.ReaderModeDivWrapper, 0);

		var hasPointer = AscCommon.AscBrowser.isIE ? ((!('ontouchstart' in window)) &&  (!!(window.PointerEvent || window.MSPointerEvent))) : false;
		if (AscCommon.AscBrowser.isAppleDevices && AscCommon.AscBrowser.iosVersion >= 13)
			hasPointer = true;

		var eventNames = hasPointer ? ["onpointerdown", "onpointermove", "onpointerup", "onpointercancel"] : ["ontouchstart", "ontouchmove", "ontouchend", "ontouchcancel"];

		this.ReaderModeDivWrapper[eventNames[0]] = function(e)
		{
			return oThis.ReaderTouchManager.onTouchStart(e);
		};
		this.ReaderModeDivWrapper[eventNames[1]]  = function(e)
		{
			return oThis.ReaderTouchManager.onTouchMove(e);
		};
		this.ReaderModeDivWrapper[eventNames[2]] = function(e)
		{
			return oThis.ReaderTouchManager.onTouchEnd(e);
		};
		this.ReaderModeDivWrapper[eventNames[3]] = function(e)
		{
			return oThis.ReaderTouchManager.onTouchEnd(e);
		};

		//this.m_oEditor.HtmlElement.style.display = "none";
		//this.m_oOverlay.HtmlElement.style.display = "none";
	};

	this.DisableReaderMode = function()
	{
		if (this.isNewReaderMode && this.m_oLogicDocument)
		{
			this.ReaderModeCurrent = 0;
			if (this.m_oDrawingDocument)
			{
				this.m_nZoomType = 1;
				this.m_oDrawingDocument.m_bUpdateAllPagesOnFirstRecalculate = true;
			}
			this.m_oLogicDocument.SetDocumentPrintMode();
			return;
		}

		this.ReaderModeCurrent = 0;
		if (null == this.ReaderModeDivWrapper)
			return;

		this.TransformDivUseAnimation(this.ReaderModeDivWrapper, parseInt(this.ReaderModeDivWrapper.style.height) + 10);
		setTimeout(this.CheckDestroyReader, 500);
		return;
	};

	this.CheckDestroyReader = function()
	{
		if (oThis.ReaderModeDivWrapper != null)
		{
			if (parseInt(oThis.ReaderModeDivWrapper.style.top) > parseInt(oThis.ReaderModeDivWrapper.style.height))
			{
				oThis.m_oMainView.HtmlElement.removeChild(oThis.ReaderModeDivWrapper);

				oThis.ReaderModeDivWrapper = null;
				oThis.ReaderModeDiv        = null;

				oThis.ReaderTouchManager.Destroy();
				oThis.ReaderTouchManager = null;

				oThis.ReaderModeCurrent = 0;

				if (oThis.m_oApi.isMobileVersion)
				{
					oThis.MobileTouchManager = new AscCommon.CMobileTouchManager({ eventsElement : "word_mobile_element" });
					oThis.MobileTouchManager.Init(oThis.m_oApi);

					if (AscCommon.g_inputContext && AscCommon.g_inputContext.HtmlArea)
						oThis.MobileTouchManager.initEvents(AscCommon.g_inputContext.HtmlArea.id);

					oThis.MobileTouchManager.Resize();

					oThis.MobileTouchManager.scrollTo(oThis.m_dScrollX, oThis.m_dScrollY);
				}

				return;
			}

			if (oThis.ReaderModeCurrent == 0)
			{
				setTimeout(oThis.CheckDestroyReader, 200);
			}
		}
	};

	// ui buttons
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

		var oWordControl = oThis;
		if (0 < oWordControl.m_oDrawingDocument.m_lCurrentPage)
		{
			oWordControl.GoToPage(oWordControl.m_oDrawingDocument.m_lCurrentPage - 1);
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

		var oWordControl = oThis;
		if ((oWordControl.m_oDrawingDocument.m_lPagesCount - 1) > oWordControl.m_oDrawingDocument.m_lCurrentPage)
		{
			oWordControl.GoToPage(oWordControl.m_oDrawingDocument.m_lCurrentPage + 1);
		}
		else if (oWordControl.m_oDrawingDocument.m_lPagesCount > 0)
		{
			oWordControl.GoToPage(oWordControl.m_oDrawingDocument.m_lPagesCount - 1);
		}
	};

	// scroll
	this.verticalScroll = function(sender, scrollPositionY, maxY, isAtTop, isAtBottom)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		var oWordControl                       = oThis;
		oWordControl.m_dScrollY                = Math.max(0, Math.min(scrollPositionY, maxY));
		oWordControl.m_dScrollY_max            = maxY;
		oWordControl.m_bIsUpdateVerRuler       = true;
		oWordControl.m_bIsUpdateTargetNoAttack = true;

		if (oWordControl.m_bIsRePaintOnScroll === true)
			oWordControl.OnScroll();

		if (oWordControl.MobileTouchManager && oWordControl.MobileTouchManager.iScroll)
		{
			oWordControl.MobileTouchManager.iScroll.y = -oWordControl.m_dScrollY;

			if ((oWordControl.MobileTouchManager.Mode === AscCommon.MobileTouchMode.None || oWordControl.MobileTouchManager.Mode === AscCommon.MobileTouchMode.Scroll) &&
				(oWordControl.MobileTouchManager.iScroll.initiated !== 0 || oWordControl.MobileTouchManager.iScroll.isAnimating))
			{
				oThis.m_oApi.sendEvent("onMobileScrollDelta", oWordControl.m_dScrollY - oWordControl.mobileScrollStartPos);
			}
		}
	};
	this.CorrectSpeedVerticalScroll = function(newScrollPos)
	{
		// без понтов.
		//return pos;

		// понты
		var res = {isChange : false, Pos : newScrollPos};
		if (newScrollPos <= 0 || newScrollPos >= this.m_dScrollY_max)
			return res;

		var _heightPageMM = Page_Height;
		if (this.m_oDrawingDocument.m_arrPages.length > 0)
			_heightPageMM = this.m_oDrawingDocument.m_arrPages[0].height_mm;
		var del = 20 + (g_dKoef_mm_to_pix * _heightPageMM * this.m_nZoomValue / 100 + 0.5) >> 0;

		var delta = Math.abs(newScrollPos - this.m_dScrollY);
		if (this.m_oDrawingDocument.m_lPagesCount <= 10)
			return res;
		else if (this.m_oDrawingDocument.m_lPagesCount <= 100 && (delta < del * 0.3))
			return res;
		else if (delta < del * 0.2)
			return res;

		var canvas = this.m_oEditor.HtmlElement;
		if (null == canvas)
			return;

		var _height      = canvas.height;
		var documentMaxY = this.m_dDocumentHeight - _height;
		if (documentMaxY <= 0)
			return res;

		var lCurrentTopInDoc = parseInt(newScrollPos * documentMaxY / this.m_dScrollY_max);
		var lCount           = parseInt(lCurrentTopInDoc / del);

		res.isChange = true;
		res.Pos      = parseInt((lCount * del) * this.m_dScrollY_max / documentMaxY);

		if (res.Pos < 0)
			res.Pos = 0;
		if (res.Pos > this.m_dScrollY_max)
			res.Pos = this.m_dScrollY_max;

		return res;
	};

	this.horizontalScroll = function(sender, scrollPositionX, maxX, isAtLeft, isAtRight)
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		var oWordControl                       = oThis;
		oWordControl.m_dScrollX                = Math.max(0, Math.min(scrollPositionX, maxX));
		oWordControl.m_dScrollX_max            = maxX;
		oWordControl.m_bIsUpdateHorRuler       = true;
		oWordControl.m_bIsUpdateTargetNoAttack = true;

		if (oWordControl.m_bIsRePaintOnScroll === true)
		{
			oWordControl.OnScroll();
		}

		if (oWordControl.MobileTouchManager && oWordControl.MobileTouchManager.iScroll)
		{
			oWordControl.MobileTouchManager.iScroll.x = -oWordControl.m_dScrollX;
		}
	};

	this.CreateScrollSettings = function()
	{
		var settings = new AscCommon.ScrollSettings();
		settings.screenW = this.m_oEditor.HtmlElement.width;
		settings.screenH = this.m_oEditor.HtmlElement.height;
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

		settings.screenW = AscCommon.AscBrowser.convertToRetinaValue(settings.screenW);
		settings.screenH = AscCommon.AscBrowser.convertToRetinaValue(settings.screenH);

		if (!this.MobileTouchManager)
			settings.screenH -= this.offsetTop;

		return settings;
	};

	this.UpdateScrolls = function()
	{
		var settings;
		if (window["NATIVE_EDITOR_ENJINE"])
			return;

		settings = this.CreateScrollSettings();
		settings.isHorizontalScroll = true;
		settings.isVerticalScroll = false;
		settings.contentW = this.m_dDocumentWidth;
		if (this.m_oScrollHor_)
			this.m_oScrollHor_.Repos(settings, this.m_bIsHorScrollVisible);
		else
		{
			this.m_oScrollHor_ = new AscCommon.ScrollObject("id_horizontal_scroll", settings);

			this.m_oScrollHor_.onLockMouse  = function(evt)
			{
				AscCommon.check_MouseDownEvent(evt, true);
				global_mouseEvent.LockMouse();
			};
			this.m_oScrollHor_.offLockMouse = function(evt)
			{
				AscCommon.check_MouseUpEvent(evt);
			};
			this.m_oScrollHor_.bind("scrollhorizontal", function(evt)
			{
				oThis.horizontalScroll(this, evt.scrollD, evt.maxScrollX);
			});
			this.m_oScrollHorApi = this.m_oScrollHor_;
		}

		settings = this.CreateScrollSettings();
		settings.isHorizontalScroll = false;
		settings.isVerticalScroll = true;
		settings.contentH = this.m_dDocumentHeight;

		if (this.MobileTouchManager)
			settings.contentH += this.offsetTop;

		if (this.m_oScrollVer_)
		{
			this.m_oScrollVer_.Repos(settings, undefined, true);
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
			this.m_oScrollVer_.bind("correctVerticalScroll", function(yPos)
			{
				return oThis.CorrectSpeedVerticalScroll(yPos);
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

	this.checkNeedHorScroll = function()
	{
		if (this.m_oApi.isMobileVersion)
		{
			this.m_oPanelRight.Bounds.B  = 0;
			this.m_oMainContent.Bounds.B = 0;

			// этот флаг для того, чтобы не делался лишний зум и т.д.
			this.m_bIsHorScrollVisible = false;
			return false;
		}

		var _editor_width = this.m_oEditor.HtmlElement.width;
		_editor_width /= AscCommon.AscBrowser.retinaPixelRatio;

		var oldVisible = this.m_bIsHorScrollVisible;
		if (this.m_dDocumentWidth <= _editor_width)
		{
			this.m_bIsHorScrollVisible = false;
		}
		else
		{
			this.m_bIsHorScrollVisible = true;
		}

		if (this.m_bIsHorScrollVisible)
		{
			this.m_oScrollHor.HtmlElement.style.display = 'block';

			this.m_oPanelRight.Bounds.B  = this.ScrollsWidthPx * g_dKoef_pix_to_mm;
			this.m_oMainContent.Bounds.B = this.ScrollsWidthPx * g_dKoef_pix_to_mm;
		}
		else
		{
			this.m_oPanelRight.Bounds.B                 = 0;
			this.m_oMainContent.Bounds.B                = 0;
			this.m_oScrollHor.HtmlElement.style.display = 'none';
		}

		if (this.m_bIsHorScrollVisible != oldVisible)
		{
			this.m_dScrollX = 0;
			this.OnResize(true);
			return true;
		}
		return false;
	};

	this.getScrollMaxX = function(size)
	{
		if (size >= this.m_dDocumentWidth)
			return 1;
		return this.m_dDocumentWidth - size;
	};
	this.getScrollMaxY = function(size)
	{
		if (size >= this.m_dDocumentHeight)
			return 1;
		return this.m_dDocumentHeight - size;
	};

	this.GetVerticalScrollTo = function(y, page)
	{
		var dKoef = g_dKoef_mm_to_pix * this.m_nZoomValue / 100;

		var lYPos = 0;
		for (var i = 0; i < page; i++)
		{
			lYPos += (20 + (this.m_oDrawingDocument.m_arrPages[i].height_mm * dKoef + 0.5) >> 0);
		}

		lYPos += y * dKoef;
		return lYPos;
	};

	this.GetHorizontalScrollTo = function(x, page)
	{
		var dKoef = g_dKoef_mm_to_pix * this.m_nZoomValue / 100;
		return 5 + dKoef * x;
	};

	this.ScrollToPosition = function(x, y, PageNum, height)
	{
		if (PageNum < 0 || PageNum >= this.m_oDrawingDocument.m_lCountCalculatePages)
			return;

		var _h       = (undefined === height) ? 5 : height;
		var rectSize = (_h * this.m_nZoomValue * g_dKoef_mm_to_pix / 100);
		var pos      = this.m_oDrawingDocument.ConvertCoordsToCursor2(x, y, PageNum);

		if (true === pos.Error)
			return;

		var _ww = this.m_oEditor.HtmlElement.width;
		var _hh = this.m_oEditor.HtmlElement.height;
		_ww /= AscCommon.AscBrowser.retinaPixelRatio;
		_hh /= AscCommon.AscBrowser.retinaPixelRatio;

		var boxX = 0;
		var boxY = 0;
		var boxR = _ww - 2;
		var boxB = _hh - rectSize;

		var nValueScrollHor = 0;
		if (pos.X < boxX)
		{
			nValueScrollHor = this.GetHorizontalScrollTo(x, PageNum);
		}
		if (pos.X > boxR)
		{
			var _mem        = x - g_dKoef_pix_to_mm * _ww * 100 / this.m_nZoomValue;
			nValueScrollHor = this.GetHorizontalScrollTo(_mem, PageNum);
		}

		var nValueScrollVer = 0;
		if (pos.Y < boxY)
		{
			nValueScrollVer = this.GetVerticalScrollTo(y, PageNum);
		}
		if (pos.Y > boxB)
		{
			var _mem        = y + _h + 10 - g_dKoef_pix_to_mm * _hh * 100 / this.m_nZoomValue;
			nValueScrollVer = this.GetVerticalScrollTo(_mem, PageNum);
		}

		var isNeedScroll = false;
		if (0 != nValueScrollHor)
		{
			isNeedScroll                   = true;
			this.m_bIsUpdateTargetNoAttack = true;
			var temp                       = nValueScrollHor * this.m_dScrollX_max / (this.m_dDocumentWidth - _ww);
			this.m_oScrollHorApi.scrollToX(parseInt(temp), false);
		}
		if (0 != nValueScrollVer)
		{
			isNeedScroll                   = true;
			this.m_bIsUpdateTargetNoAttack = true;
			var temp                       = nValueScrollVer * this.m_dScrollY_max / (this.m_dDocumentHeight - _hh);
			this.m_oScrollVerApi.scrollToY(parseInt(temp), false);
		}

		if (true === isNeedScroll)
		{
			this.OnScroll();
			return;
		}
	};

	this.ScrollToAbsolutePosition = function(x, y, page, isBottom)
	{
		let pos = this.m_oDrawingDocument.ConvertCoordsToCursor(x, y, page);
		if (pos.Error)
			return;

		if (true === isBottom)
			pos.Y -= AscCommon.AscBrowser.convertToRetinaValue(this.m_oEditor.HtmlElement.height);

		// TODO: X position?
		if (0 !== pos.Y)
			this.m_oScrollVerApi.scrollToY(pos.Y + this.m_dScrollY);
	};

	// events ---
	this.onMouseDown = function(e, isTouch)
	{
		if (oThis.MobileTouchManager && oThis.MobileTouchManager.checkTouchEvent(e))
		{
			oThis.MobileTouchManager.startTouchingInProcess();
			let res = oThis.MobileTouchManager.mainOnTouchStart(e);
			oThis.MobileTouchManager.stopTouchingInProcess();
			return res;
		}

		if (oThis.MobileTouchManager)
			oThis.MobileTouchManager.checkMouseFocus(e);

		oThis.m_oApi.checkInterfaceElementBlur();
		oThis.m_oApi.checkLastWork();

		//console.log("down: " + isTouch + ", " + AscCommon.isTouch);
		if (false === oThis.m_oApi.bInit_word_control || (AscCommon.isTouch && undefined === isTouch) || oThis.m_oApi.isLongAction())
			return;

		if (!oThis.m_bIsIE)
		{
			if (e.preventDefault)
				e.preventDefault();
			else
				e.returnValue = false;
		}

		if (AscCommon.g_inputContext && AscCommon.g_inputContext.externalChangeFocus())
			return;

		var oWordControl = oThis;

		if (this.id == "id_viewer" && oThis.m_oOverlay.HtmlElement.style.display == "block")
			return;

		var _xOffset = oWordControl.X;
		var _yOffset = oWordControl.Y;

		let offsets = oWordControl.m_oDrawingDocument.GetMainOffset();
		_xOffset += offsets.x;
		_yOffset += offsets.y;

		if (window['closeDialogs'] != undefined)
			window['closeDialogs']();

		AscCommon.check_MouseDownEvent(e, true);
		global_mouseEvent.LockMouse();

		// у Илюхи есть проблема при вводе с клавы, пока нажата кнопка мыши
		if ((0 == global_mouseEvent.Button) || (undefined == global_mouseEvent.Button))
		{
			oWordControl.m_bIsMouseLock = true;
		}

		var pos = null;

		if (oWordControl.MouseHandObject)
		{
			pos = oWordControl.m_oDrawingDocument.ConvertCoordsFromCursor2(global_mouseEvent.X, global_mouseEvent.Y);
			if (oWordControl.MouseHandObject.check(oWordControl, pos))
			{
				oWordControl.MouseHandObject.X = global_mouseEvent.X;
				oWordControl.MouseHandObject.Y = global_mouseEvent.Y;
				oWordControl.MouseHandObject.Active = true;
				oWordControl.MouseHandObject.ScrollX = oWordControl.m_dScrollX;
				oWordControl.MouseHandObject.ScrollY = oWordControl.m_dScrollY;
				oWordControl.m_oDrawingDocument.SetCursorType(AscCommon.Cursors.Grabbing);

				oWordControl.m_oLogicDocument && oWordControl.m_oLogicDocument.EndFormEditing();
				return;
			}
		}

		oWordControl.StartUpdateOverlay();
		var bIsSendSelectWhell = false;

		if ((0 == global_mouseEvent.Button) || (undefined == global_mouseEvent.Button))
		{
			if (oWordControl.m_oDrawingDocument.AutoShapesTrackLockPageNum == -1)
				pos = oWordControl.m_oDrawingDocument.ConvertCoordsFromCursor2(global_mouseEvent.X, global_mouseEvent.Y);
			else
				pos = oWordControl.m_oDrawingDocument.ConvetToPageCoords(global_mouseEvent.X, global_mouseEvent.Y, oWordControl.m_oDrawingDocument.AutoShapesTrackLockPageNum);

			if (pos.Page == -1)
			{
				oWordControl.EndUpdateOverlay();
				return;
			}

			if (oWordControl.m_oDrawingDocument.IsFreezePage(pos.Page))
			{
				oWordControl.EndUpdateOverlay();
				return;
			}

			if (null == oWordControl.m_oDrawingDocument.m_oDocumentRenderer)
			{
				// теперь проверить трек таблиц
				var ret = oWordControl.m_oDrawingDocument.checkMouseDown_Drawing(pos);
				if (ret === true)
				{
					if (-1 == oWordControl.m_oTimerScrollSelect && AscCommon.global_mouseEvent.IsLocked)
						oWordControl.m_oTimerScrollSelect = setInterval(oWordControl.SelectWheel, 20);

					if (oWordControl.MobileTouchManager && oWordControl.MobileTouchManager.iScroll)
						oWordControl.MobileTouchManager.iScroll.disableLongTapAction();

					if (!oWordControl.m_oApi.getHandlerOnClick())
						AscCommon.stopEvent(e);
					oWordControl.EndUpdateOverlay();
					return;
				}

				if (-1 == oWordControl.m_oTimerScrollSelect && AscCommon.global_mouseEvent.IsLocked)
				{
					// добавим это и здесь, чтобы можно было отменять во время LogicDocument.OnMouseDown
					oWordControl.m_oTimerScrollSelect = setInterval(oWordControl.SelectWheel, 20);
					bIsSendSelectWhell                = true;
				}

				oWordControl.m_oDrawingDocument.NeedScrollToTargetFlag = true;
				if(!oThis.m_oApi.isEyedropperStarted())
				{
					oWordControl.m_oLogicDocument.OnMouseDown(global_mouseEvent, pos.X, pos.Y, pos.Page);
				}
				oWordControl.m_oDrawingDocument.NeedScrollToTargetFlag = false;

				oWordControl.MouseDownDocumentCounter++;
			}
			else
			{
				oWordControl.m_oDrawingDocument.m_oDocumentRenderer.OnMouseDown(pos.Page, pos.X, pos.Y);
				oWordControl.MouseDownDocumentCounter++;
			}
		}
		else if (global_mouseEvent.Button == 2)
			oWordControl.MouseDownDocumentCounter++;

		if (!bIsSendSelectWhell && -1 == oWordControl.m_oTimerScrollSelect && AscCommon.global_mouseEvent.IsLocked)
		{
			oWordControl.m_oTimerScrollSelect = setInterval(oWordControl.SelectWheel, 20);
		}

		oWordControl.EndUpdateOverlay();

		oWordControl.m_oDrawingDocument.isHideTargetBeforeFirstClick = false;
	};

	this.onMouseMove  = function(e, isTouch)
	{
		if (oThis.MobileTouchManager && oThis.MobileTouchManager.checkTouchEvent(e))
		{
			oThis.MobileTouchManager.startTouchingInProcess();
			let res = oThis.MobileTouchManager.mainOnTouchMove(e);
			oThis.MobileTouchManager.stopTouchingInProcess();
			return res;
		}

		oThis.m_oApi.checkLastWork();

		if (false === oThis.m_oApi.bInit_word_control || (AscCommon.isTouch && undefined === isTouch) || oThis.m_oApi.isLongAction())
			return;

		if (e)
		{
			if (!oThis.m_bIsMouseLock)
			{
				if (e.preventDefault)
					e.preventDefault();
				else
					e.returnValue = false;
			}
			else
			{
				AscCommon.stopEvent(e);
			}

			AscCommon.check_MouseMoveEvent(e);
		}

		var oWordControl = oThis;

		var pos = null;
		if (oWordControl.MouseHandObject)
		{
			if (oWordControl.MouseHandObject.Active)
			{
				oWordControl.m_oDrawingDocument.SetCursorType(AscCommon.Cursors.Grabbing);

				var scrollX = global_mouseEvent.X - oWordControl.MouseHandObject.X;
				var scrollY = global_mouseEvent.Y - oWordControl.MouseHandObject.Y;

				if (0 != scrollX && oWordControl.m_bIsHorScrollVisible)
					oWordControl.m_oScrollHorApi.scrollToX(oWordControl.MouseHandObject.ScrollX - scrollX);
				if (0 != scrollY)
					oWordControl.m_oScrollVerApi.scrollToY(oWordControl.MouseHandObject.ScrollY - scrollY);

				return;
			}
			else if (!global_mouseEvent.IsLocked)
			{
				pos = oWordControl.m_oDrawingDocument.ConvertCoordsFromCursor2(global_mouseEvent.X, global_mouseEvent.Y);
				if (oWordControl.MouseHandObject.check(oWordControl, pos))
				{
					oThis.m_oApi.sync_MouseMoveStartCallback();
					oThis.m_oApi.sync_MouseMoveCallback(new AscCommon.CMouseMoveData());
					oThis.m_oApi.sync_MouseMoveEndCallback();

					oWordControl.m_oDrawingDocument.SetCursorType(AscCommon.Cursors.Grab);

					oWordControl.m_oLogicDocument && oWordControl.m_oLogicDocument.UpdateCursorType();
					oWordControl.StartUpdateOverlay();
					oWordControl.m_oDrawingDocument.contentControls.onPointerLeave();
					oWordControl.OnUpdateOverlay();
					oWordControl.EndUpdateOverlay();
					return;
				}
			}
		}

		if (oWordControl.m_oDrawingDocument.AutoShapesTrackLockPageNum == -1)
			pos = oWordControl.m_oDrawingDocument.ConvertCoordsFromCursor2(global_mouseEvent.X, global_mouseEvent.Y);
		else
			pos = oWordControl.m_oDrawingDocument.ConvetToPageCoords(global_mouseEvent.X, global_mouseEvent.Y, oWordControl.m_oDrawingDocument.AutoShapesTrackLockPageNum);

		if (pos.Page == -1)
			return;

		if (oWordControl.m_oDrawingDocument.IsFreezePage(pos.Page))
			return;

		if (oWordControl.m_oDrawingDocument.m_sLockedCursorType != "")
			oWordControl.m_oDrawingDocument.SetCursorType("default");
		if (oWordControl.m_oDrawingDocument.m_oDocumentRenderer != null)
		{
			oWordControl.m_oDrawingDocument.m_oDocumentRenderer.OnMouseMove(pos.Page, pos.X, pos.Y);
			return;
		}

		if(oThis.m_oApi.isEyedropperStarted())
		{
			let oParentPos = oWordControl.m_oMainView.AbsolutePosition;
			let nX  = global_mouseEvent.X - oWordControl.X - oParentPos.L * AscCommon.g_dKoef_mm_to_pix;
			let nY  = global_mouseEvent.Y - oWordControl.Y - oParentPos.T * AscCommon.g_dKoef_mm_to_pix;
			nX = AscCommon.AscBrowser.convertToRetinaValue(nX, true);
			nY = AscCommon.AscBrowser.convertToRetinaValue(nY, true);
			oThis.m_oApi.checkEyedropperColor(nX, nY);
			oThis.m_oApi.sync_MouseMoveStartCallback();
			let MMData = new AscCommon.CMouseMoveData();
			let Coords = oWordControl.m_oDrawingDocument.ConvertCoordsToCursorWR(pos.X, pos.Y, pos.Page, null);
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
		var is_drawing = oWordControl.m_oDrawingDocument.checkMouseMove_Drawing(pos, e === undefined ? true : false);
		if (is_drawing === true)
			return;

		oWordControl.m_oDrawingDocument.TableOutlineDr.bIsNoTable = true;
		oWordControl.m_oLogicDocument.OnMouseMove(global_mouseEvent, pos.X, pos.Y, pos.Page);

		if (oWordControl.m_oDrawingDocument.TableOutlineDr.bIsNoTable === false)
		{
			// TODO: нужно посмотреть, может в ЭТОМ же месте трек для таблицы уже нарисован
			oWordControl.ShowOverlay();
			oWordControl.OnUpdateOverlay();
		}

		if (!oWordControl.IsUpdateOverlayOnEndCheck)
		{
			if (oWordControl.m_oDrawingDocument.contentControls && oWordControl.m_oDrawingDocument.contentControls.ContentControlsCheckLast())
				oWordControl.OnUpdateOverlay();
		}

		oWordControl.EndUpdateOverlay();
	};
	this.onMouseMove2 = function()
	{
		if (false === oThis.m_oApi.bInit_word_control)
			return;

		var oWordControl = oThis;
		var pos          = null;
		if (oWordControl.m_oDrawingDocument.AutoShapesTrackLockPageNum == -1)
			pos = oWordControl.m_oDrawingDocument.ConvertCoordsFromCursor2(global_mouseEvent.X, global_mouseEvent.Y);
		else
			pos = oWordControl.m_oDrawingDocument.ConvetToPageCoords(global_mouseEvent.X, global_mouseEvent.Y, oWordControl.m_oDrawingDocument.AutoShapesTrackLockPageNum);

		if (pos.Page == -1)
			return;

		if (null != oWordControl.m_oDrawingDocument.m_oDocumentRenderer)
		{
			oWordControl.m_oDrawingDocument.m_oDocumentRenderer.OnMouseMove(pos.Page, pos.X, pos.Y);
			return;
		}

		if (oWordControl.m_oDrawingDocument.IsFreezePage(pos.Page))
			return;

		oWordControl.StartUpdateOverlay();

		var is_drawing = oWordControl.m_oDrawingDocument.checkMouseMove_Drawing(pos);
		if (is_drawing === true)
			return;

		oWordControl.m_oLogicDocument.OnMouseMove(global_mouseEvent, pos.X, pos.Y, pos.Page);
		oWordControl.EndUpdateOverlay();
	};

	this.onMouseUp    = function(e, bIsWindow, isTouch)
	{
		if (oThis.MobileTouchManager && oThis.MobileTouchManager.checkTouchEvent(e))
		{
			oThis.MobileTouchManager.startTouchingInProcess();
			let res = oThis.MobileTouchManager.mainOnTouchEnd(e);
			oThis.MobileTouchManager.stopTouchingInProcess();
			return res;
		}

		oThis.m_oApi.checkLastWork();

		//console.log("up: " + isTouch + ", " + AscCommon.isTouch);
		if (false === oThis.m_oApi.bInit_word_control || (AscCommon.isTouch && undefined === isTouch) || oThis.m_oApi.isLongAction())
			return;
		//if (true == global_mouseEvent.IsLocked)
		//    return;

		var oWordControl = oThis;

		if (oWordControl.MouseHandObject && oWordControl.MouseHandObject.Active)
		{
			AscCommon.check_MouseUpEvent(e);
			oWordControl.MouseHandObject.Active = false;
			oWordControl.m_oDrawingDocument.SetCursorType(AscCommon.Cursors.Grab);
			oWordControl.m_bIsMouseLock = false;
			return;
		}

		if (!global_mouseEvent.IsLocked && 0 == oWordControl.MouseDownDocumentCounter)
			return;

		if (this.id == "id_viewer" && oThis.m_oOverlay.HtmlElement.style.display == "block" && undefined == bIsWindow)
			return;

		if ((global_mouseEvent.Sender != oThis.m_oEditor.HtmlElement &&
				global_mouseEvent.Sender != oThis.m_oOverlay.HtmlElement &&
				global_mouseEvent.Sender != oThis.m_oDrawingDocument.TargetHtmlElement) &&
			(oThis.TextBoxBackground && oThis.TextBoxBackground.HtmlElement != global_mouseEvent.Sender))
			return;

		if (-1 != oWordControl.m_oTimerScrollSelect)
		{
			clearInterval(oWordControl.m_oTimerScrollSelect);
			oWordControl.m_oTimerScrollSelect = -1;
		}

		if (oWordControl.m_oHorRuler.m_bIsMouseDown)
			oWordControl.m_oHorRuler.OnMouseUpExternal();

		if (oWordControl.m_oVerRuler.DragType != 0)
			oWordControl.m_oVerRuler.OnMouseUpExternal();

		if (oWordControl.m_oScrollVerApi.getIsLockedMouse())
		{
			oWordControl.m_oScrollVerApi.evt_mouseup(e);
		}
		if (oWordControl.m_oScrollHorApi.getIsLockedMouse())
		{
			oWordControl.m_oScrollHorApi.evt_mouseup(e);
		}

		AscCommon.check_MouseUpEvent(e);
		var pos = null;
		if (oWordControl.m_oDrawingDocument.AutoShapesTrackLockPageNum == -1)
			pos = oWordControl.m_oDrawingDocument.ConvertCoordsFromCursor2(global_mouseEvent.X, global_mouseEvent.Y);
		else
			pos = oWordControl.m_oDrawingDocument.ConvetToPageCoords(global_mouseEvent.X, global_mouseEvent.Y, oWordControl.m_oDrawingDocument.AutoShapesTrackLockPageNum);

		if (pos.Page == -1)
			return;

		if (oWordControl.m_oDrawingDocument.IsFreezePage(pos.Page))
			return;

		oWordControl.UnlockCursorTypeOnMouseUp();

		oWordControl.StartUpdateOverlay();

		// восстанавливаем фокус
		oWordControl.m_bIsMouseLock = false;

		var is_drawing = oWordControl.m_oDrawingDocument.checkMouseUp_Drawing(pos);
		if (is_drawing === true)
			return;

		if (null != oWordControl.m_oDrawingDocument.m_oDocumentRenderer)
		{
			oWordControl.m_oDrawingDocument.m_oDocumentRenderer.OnMouseUp();

			oWordControl.MouseDownDocumentCounter--;
			if (oWordControl.MouseDownDocumentCounter < 0)
				oWordControl.MouseDownDocumentCounter = 0;

			oWordControl.EndUpdateOverlay();
			return;
		}

		oWordControl.m_bIsMouseUpSend = true;

		if (2 == global_mouseEvent.Button)
		{
			// пошлем сначала моусдаун
			//oWordControl.m_oLogicDocument.OnMouseDown(global_mouseEvent, pos.X, pos.Y, pos.Page);
		}
		oWordControl.m_oDrawingDocument.NeedScrollToTargetFlag = true;
		if(!oThis.checkFinishEyedropper())
		{
			oWordControl.m_oLogicDocument.OnMouseUp(global_mouseEvent, pos.X, pos.Y, pos.Page);
		}
		oWordControl.m_oDrawingDocument.NeedScrollToTargetFlag = false;

		oWordControl.MouseDownDocumentCounter--;
		if (oWordControl.MouseDownDocumentCounter < 0)
			oWordControl.MouseDownDocumentCounter = 0;

		oWordControl.m_bIsMouseUpSend = false;

		oWordControl.EndUpdateOverlay();

		if (AscCommon.check_MouseClickOnUp())
		{
			if (window.g_asc_plugins)
				window.g_asc_plugins.onPluginEvent("onClick", oWordControl.m_oLogicDocument.IsSelectionUse());
		}
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

		if (oWordControl.MouseHandObject && oWordControl.MouseHandObject.Active)
		{
			oWordControl.MouseHandObject.Active = false;
			oWordControl.m_oDrawingDocument.SetCursorType(AscCommon.Cursors.Grab);
			return;
		}

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

		let bIsSendedToEditor = true;
		if (!global_mouseEvent.IsLocked && 0 == oWordControl.MouseDownDocumentCounter)
			bIsSendedToEditor = false;

		AscCommon.MouseUpLock.MouseUpLockedSend = true;

		if (oWordControl.m_oHorRuler.m_bIsMouseDown)
			oWordControl.m_oHorRuler.OnMouseUpExternal();

		if (oWordControl.m_oVerRuler.DragType != 0)
			oWordControl.m_oVerRuler.OnMouseUpExternal();

		if (oWordControl.m_oScrollVerApi.getIsLockedMouse())
		{
			var __e = {clientX : x, clientY : y};
			oWordControl.m_oScrollVerApi.evt_mouseup(__e);
		}
		if (oWordControl.m_oScrollHorApi.getIsLockedMouse())
		{
			var __e = {clientX : x, clientY : y};
			oWordControl.m_oScrollHorApi.evt_mouseup(__e);
		}

		if (window.g_asc_plugins)
			window.g_asc_plugins.onExternalMouseUp();

		global_mouseEvent.Sender = null;

		global_mouseEvent.UnLockMouse();

		global_mouseEvent.IsPressed = false;

		if (oWordControl.MouseHandObject && oWordControl.MouseHandObject.Active)
		{
			oWordControl.MouseHandObject.Active = false;
			oWordControl.m_oDrawingDocument.SetCursorType(AscCommon.Cursors.Grab);
			return;
		}

		if (-1 != oWordControl.m_oTimerScrollSelect)
		{
			clearInterval(oWordControl.m_oTimerScrollSelect);
			oWordControl.m_oTimerScrollSelect = -1;
		}

		//---
		var pos = null;
		if (oWordControl.m_oDrawingDocument.AutoShapesTrackLockPageNum == -1)
			pos = oWordControl.m_oDrawingDocument.ConvertCoordsFromCursor2(global_mouseEvent.X, global_mouseEvent.Y);
		else
			pos = oWordControl.m_oDrawingDocument.ConvetToPageCoords(global_mouseEvent.X, global_mouseEvent.Y, oWordControl.m_oDrawingDocument.AutoShapesTrackLockPageNum);

		if (pos.Page == -1)
			return;

		if (oWordControl.m_oDrawingDocument.IsFreezePage(pos.Page))
			return;

		oWordControl.UnlockCursorTypeOnMouseUp();

		oWordControl.StartUpdateOverlay();

		// восстанавливаем фокус
		oWordControl.m_bIsMouseLock = false;

		var is_drawing = oWordControl.m_oDrawingDocument.checkMouseUp_Drawing(pos);
		if (is_drawing === true)
			return;

		if (null != oWordControl.m_oDrawingDocument.m_oDocumentRenderer)
		{
			oWordControl.m_oDrawingDocument.m_oDocumentRenderer.OnMouseUp();

			oWordControl.MouseDownDocumentCounter--;
			if (oWordControl.MouseDownDocumentCounter < 0)
				oWordControl.MouseDownDocumentCounter = 0;

			oWordControl.EndUpdateOverlay();
			return;
		}

		oWordControl.m_bIsMouseUpSend = true;

		if (2 == global_mouseEvent.Button)
		{
			// пошлем сначала моусдаун
			//oWordControl.m_oLogicDocument.OnMouseDown(global_mouseEvent, pos.X, pos.Y, pos.Page);
		}
		if (!oThis.checkFinishEyedropper())
		{
			if (bIsSendedToEditor)
				oWordControl.m_oLogicDocument.OnMouseUp(global_mouseEvent, pos.X, pos.Y, pos.Page);
		}
		oWordControl.MouseDownDocumentCounter--;
		if (oWordControl.MouseDownDocumentCounter < 0)
			oWordControl.MouseDownDocumentCounter = 0;

		oWordControl.m_bIsMouseUpSend = false;

		if (bIsSendedToEditor)
		{
			oWordControl.m_oLogicDocument.Document_UpdateInterfaceState();
			oWordControl.m_oLogicDocument.Document_UpdateRulersState();
		}

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

		let values = AscCommon.checkMouseWhell(e, {
			isSupportBidirectional : false,
			isAllowHorizontal : oThis.m_bIsHorScrollVisible,
			isUseMaximumDelta : true
		});

		if (0 !== values.x)
			oThis.m_oScrollHorApi.scrollBy(values.x, 0, false);
		if (0 !== values.y)
			oThis.m_oScrollVerApi.scrollBy(0, values.y, false);

		// здесь - имитируем моус мув ---------------------------
		var _e   = {};
		_e.pageX = global_mouseEvent.X;
		_e.pageY = global_mouseEvent.Y;

		_e.clientX = global_mouseEvent.X;
		_e.clientY = global_mouseEvent.Y;

		_e.altKey   = global_mouseEvent.AltKey;
		_e.shiftKey = global_mouseEvent.ShiftKey;
		_e.ctrlKey  = global_mouseEvent.CtrlKey;
		_e.metaKey  = global_mouseEvent.CtrlKey;

		_e.srcElement = global_mouseEvent.Sender;

		oThis.onMouseMove(_e);
		// ------------------------------------------------------

		if (e.preventDefault)
			e.preventDefault();
		else
			e.returnValue = false;

		return false;
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

		// если находимся в самом верху (без тулбара) - то наверх не будем скроллиться
		// делаем заглушку
		var minPosY = 20;
		minPosY *= AscCommon.AscBrowser.retinaPixelRatio;
		if (positionMinY < minPosY)
			positionMinY = minPosY;

		var positionMaxY = oWordControl.m_oMainContent.AbsolutePosition.B * g_dKoef_mm_to_pix + oWordControl.Y;

		var scrollYVal = 0;
		if (global_mouseEvent.Y < positionMinY)
		{
			var delta = 30;
			if (20 > (positionMinY - global_mouseEvent.Y))
				delta = 10;

			scrollYVal = -delta;
			//oWordControl.m_oScrollVerApi.scrollByY(-delta, false);
			//oWordControl.onMouseMove2();
			//return;
		}
		else if (global_mouseEvent.Y > positionMaxY)
		{
			var delta = 30;
			if (20 > (global_mouseEvent.Y - positionMaxY))
				delta = 10;

			scrollYVal = delta;
			//oWordControl.m_oScrollVerApi.scrollByY(delta, false);
			//oWordControl.onMouseMove2();
			//return;
		}

		var scrollXVal = 0;
		if (oWordControl.m_bIsHorScrollVisible)
		{
			var positionMinX = oWordControl.m_oMainContent.AbsolutePosition.L * g_dKoef_mm_to_pix + oWordControl.X;
			if (oWordControl.m_bIsRuler)
				positionMinX += oWordControl.m_oLeftRuler.AbsolutePosition.R * g_dKoef_mm_to_pix;

			var positionMaxX = oWordControl.m_oMainContent.AbsolutePosition.R * g_dKoef_mm_to_pix + oWordControl.X;

			if (global_mouseEvent.X < positionMinX)
			{
				var delta = 30;
				if (20 > (positionMinX - global_mouseEvent.X))
					delta = 10;

				scrollXVal = -delta;
				//oWordControl.m_oScrollHorApi.scrollByX(-delta, false);
				//oWordControl.onMouseMove2();
				//return;
			}
			else if (global_mouseEvent.X > positionMaxX)
			{
				var delta = 30;
				if (20 > (global_mouseEvent.X - positionMaxX))
					delta = 10;

				scrollXVal = delta;
				//oWordControl.m_oScrollVerApi.scrollByX(delta, false);
				//oWordControl.onMouseMove2();
				//return;
			}
		}

		if (0 != scrollYVal)
			oWordControl.m_oScrollVerApi.scrollByY(scrollYVal, false);
		if (0 != scrollXVal)
			oWordControl.m_oScrollHorApi.scrollByX(scrollXVal, false);

		if (scrollXVal != 0 || scrollYVal != 0)
			oWordControl.onMouseMove2();
	};

	this.checkViewerModeKeys = function(e)
	{
		var isSendEditor = false;
		if (e.KeyCode == 33) // PgUp
		{
			//
		}
		else if (e.KeyCode == 34) // PgDn
		{
			//
		}
		else if (e.KeyCode == 35) // клавиша End
		{
			if (true === e.CtrlKey) // Ctrl + End - переход в конец документа
			{
				oThis.m_oScrollVerApi.scrollTo(0, oThis.m_dScrollY_max);
			}
		}
		else if (e.KeyCode == 36) // клавиша Home
		{
			if (true === e.CtrlKey) // Ctrl + Home - переход в начало документа
			{
				oThis.m_oScrollVerApi.scrollTo(0, 0);
			}
		}
		else if (e.KeyCode == 37) // Left Arrow
		{
			if (oThis.m_bIsHorScrollVisible)
			{
				oThis.m_oScrollHorApi.scrollBy(-30, 0, false);
			}
		}
		else if (e.KeyCode == 38) // Top Arrow
		{
			oThis.m_oScrollVerApi.scrollBy(0, -30, false);
		}
		else if (e.KeyCode == 39) // Right Arrow
		{
			if (oThis.m_bIsHorScrollVisible)
			{
				oThis.m_oScrollHorApi.scrollBy(30, 0, false);
			}
		}
		else if (e.KeyCode == 40) // Bottom Arrow
		{
			oThis.m_oScrollVerApi.scrollBy(0, 30, false);
		}
		else if (e.KeyCode == 65 && true === e.CtrlKey) // Ctrl + A - выделяем все
		{
			isSendEditor = true;
		}
		else if (e.KeyCode == 67 && true === e.CtrlKey) // Ctrl + C + ...
		{
			if (false === e.ShiftKey)
			{
				AscCommon.Editor_Copy(oThis.m_oApi);
				//не возвращаем true чтобы не было preventDefault
			}
		}
		return isSendEditor;
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

		if (oWordControl.m_bIsRuler && oWordControl.m_oHorRuler.m_bIsMouseDown)
		{
			AscCommon.check_KeyboardEvent2(e);
			e.preventDefault();
			return;
		}

		// Esc даем делать с клавы, даже когда мышка зажата, чтобы можно было сбросить drag-n-drop, но если у нас
		// идет работа с автофигурами (любые движения), тогда не пропускаем.
		if (oWordControl.m_bIsMouseLock === true && (27 !== e.keyCode || true === oWordControl.m_oLogicDocument.Is_TrackingDrawingObjects()))
		{
			if (!window.USER_AGENT_MACOS)
			{
				AscCommon.check_KeyboardEvent2(e);
				e.preventDefault();
				return;
			}

			// на масОс есть мега выделение на трекпаде. там моусАп приходит с задержкой.
			// нужно лдибо копить команды клавиатуры, либо насильно заранее сделать моусАп самому.
			// мы выбараем второе

			oWordControl.onMouseUpExternal(global_mouseEvent.X, global_mouseEvent.Y);
		}

		AscCommon.check_KeyboardEvent(e);
		if (oWordControl.IsFocus === false && e.emulated !== true)
		{
			// некоторые команды нужно продолжать обрабатывать
			if (!oWordControl.onKeyDownNoActiveControl(global_keyboardEvent))
				return;
		}

		/*
		 if (oThis.ChangeHintProps())
		 {
		 e.preventDefault();
		 oThis.OnScroll();
		 return;
		 }
		 */
		if (Asc.editor.isPdfEditor())
		{
			let ret = oWordControl.m_oDrawingDocument.m_oDocumentRenderer.OnKeyDown(global_keyboardEvent);
			if (ret === keydownresult_PreventAll || ret === true)
			{
				e.preventDefault();
			}
			return;
		}
		/*
		 if (oWordControl.m_oDrawingDocument.IsFreezePage(oWordControl.m_oDrawingDocument.m_lCurrentPage))
		 {
		 e.preventDefault();
		 return;
		 }
		 */
		/*
		 if (oWordControl.m_oApi.isViewMode)
		 {
		 var isSendToEditor = oWordControl.checkViewerModeKeys(global_keyboardEvent);
		 if (false === isSendToEditor)
		 return;
		 }
		 */

		oWordControl.StartUpdateOverlay();

		oWordControl.IsKeyDownButNoPress = true;
		var _ret_mouseDown               = oWordControl.m_oLogicDocument.OnKeyDown(global_keyboardEvent);
		oWordControl.bIsUseKeyPress      = ((_ret_mouseDown & keydownresult_PreventKeyPress) != 0) ? false : true;

		oWordControl.EndUpdateOverlay();

		if ((_ret_mouseDown & keydownresult_PreventDefault) != 0)// || (true === global_keyboardEvent.AltKey && !AscBrowser.isMacOs))
		{
			// убираем превент с альтом. Уж больно итальянцы недовольны.
			e.preventDefault();
		}

		/*
		 if (false === oWordControl.TextboxUsedForSpecials)
		 {
		 if (false === oWordControl.bIsUseKeyPress || true === global_keyboardEvent.AltKey)
		 {
		 e.preventDefault();
		 }
		 }
		 else
		 {
		 if (true !== global_keyboardEvent.AltKey && true !== global_keyboardEvent.CtrlKey)
		 {
		 if (false === oWordControl.bIsUseKeyPress)
		 e.preventDefault();
		 }
		 }
		 */
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

	this.onKeyUp    = function(e)
	{
		global_keyboardEvent.AltKey   = false;
		global_keyboardEvent.CtrlKey  = false;
		global_keyboardEvent.ShiftKey = false;
		global_keyboardEvent.AltGr    = false;
	};
	this.onKeyPress = function(e)
	{
		if (AscCommon.g_clipboardBase.IsWorking())
			return;

		if (oThis.m_oApi.isLongAction())
		{
			e.preventDefault();
			return;
		}

		var oWordControl = oThis;
		if (false === oWordControl.m_oApi.bInit_word_control || oWordControl.IsFocus === false || oWordControl.m_bIsMouseLock === true)
			return;

		if (window.opera && !oWordControl.IsKeyDownButNoPress)
		{
			oWordControl.StartUpdateOverlay();
			oWordControl.onKeyDown(e);
			oWordControl.EndUpdateOverlay();
		}
		oWordControl.IsKeyDownButNoPress = false;

		if (false === oWordControl.bIsUseKeyPress)
			return;

		if (null == oWordControl.m_oLogicDocument)
			return;

		AscCommon.check_KeyboardEvent(e);

		oWordControl.StartUpdateOverlay();
		var retValue = oWordControl.m_oLogicDocument.OnKeyPress(global_keyboardEvent);
		oWordControl.EndUpdateOverlay();
		if (true === retValue)
			e.preventDefault();
	};

	// paint loop
	this.StartMainTimer = function()
	{
		this.paintMessageLoop.Start(this.onTimerScroll.bind(this));
	};

	this.onTimerScroll_internal = function()
	{
		var oWordControl = oThis;

		if (!oWordControl.m_oApi.bInit_word_control)
			return;

		var isLongAction = oWordControl.m_oApi.isLongAction();
		if (isLongAction && oWordControl.m_oDrawingDocument && oWordControl.m_oDrawingDocument.isDisableEditBeforeCalculateLA)
			isLongAction = false;

		if (isLongAction)
			return;

		var isRepaint = oWordControl.m_bIsScroll;
		if (null != oWordControl.m_oLogicDocument && !oWordControl.m_oApi.isLockTargetUpdate)
		{
			oWordControl.m_oDrawingDocument.UpdateTargetFromPaint = true;
			oWordControl.m_oLogicDocument.CheckTargetUpdate();
			oWordControl.m_oDrawingDocument.CheckTargetShow();
			oWordControl.m_oDrawingDocument.UpdateTargetFromPaint = false;

			oWordControl.CheckFontCache();
			oWordControl.m_oDrawingDocument.CheckTrackTable();
		}
		if (oWordControl.m_bIsScroll)
		{
			isRepaint = true;
			oWordControl.m_bIsScroll = false;
			oWordControl.OnPaint();

			if (null != oWordControl.m_oLogicDocument && oWordControl.m_oApi.bInit_word_control)
				oWordControl.m_oLogicDocument.Viewer_OnChangePosition();
		}

		oWordControl.m_oDrawingDocument.Collaborative_TargetsUpdate(isRepaint);

		oWordControl.m_oApi.sendEvent("asc_onPaintTimer");
	};

	this.onTimerScroll = function()
	{
		try
		{
			oThis.onTimerScroll_internal();
		}
		catch (err)
		{
		}
	};

	this.onTimerScroll_sync = function()
	{
		var oWordControl = oThis;

		if (!oWordControl.m_oApi.bInit_word_control || oWordControl.m_oApi.isOnlyReaderMode)
			return;

		var isRepaint                   = oWordControl.m_bIsScroll;
		if (oWordControl.m_bIsScroll)
		{
			oWordControl.m_bIsScroll = false;
			oWordControl.OnPaint();

			if (null != oWordControl.m_oLogicDocument && oWordControl.m_oApi.bInit_word_control)
				oWordControl.m_oLogicDocument.Viewer_OnChangePosition();
		}
		if (null != oWordControl.m_oLogicDocument)
		{
			oWordControl.m_oDrawingDocument.UpdateTargetFromPaint = true;
			oWordControl.m_oLogicDocument.CheckTargetUpdate();
			oWordControl.m_oDrawingDocument.CheckTargetShow();
			oWordControl.m_oDrawingDocument.UpdateTargetFromPaint = false;

			oWordControl.CheckFontCache();
			oWordControl.m_oDrawingDocument.CheckTrackTable();
		}
		oWordControl.m_oDrawingDocument.Collaborative_TargetsUpdate(isRepaint);
	};

	// paint & position ---
	this.private_RefreshAll = function()
	{
		AscCommon.g_fontManager.ClearFontsRasterCache();

		if (AscCommon.g_fontManager2)
			AscCommon.g_fontManager2.ClearFontsRasterCache();

		this.OnRePaintAttack();
	};

	this.OnRePaintAttack = function()
	{
		this.m_bIsFullRepaint = true;
		this.OnScroll();

		if (this.m_oApi.isUseNativeViewer)
		{
			var oViewer = this.m_oDrawingDocument.m_oDocumentRenderer;
			if (oViewer)
			{
				oViewer.isClearPages = true;
				oViewer.paint();
			}
		}
	};

	this.OnPaint = function()
	{
		var isNoPaint = this.m_oApi.isLongAction();
		if (isNoPaint && this.m_oDrawingDocument && this.m_oDrawingDocument.isDisableEditBeforeCalculateLA)
			isNoPaint = false;

		if (isNoPaint)
		{
			if (false === this.IsRepaintOnCallbackLongAction)
			{
				this.m_oApi.checkLongActionCallback(this.OnScroll.bind(this), true);
			}
			this.IsRepaintOnCallbackLongAction = true;
			return;
		}

		if (this.DrawingFreeze || true === window["DisableVisibleComponents"])
		{
			this.m_oApi.checkLastWork();
			return;
		}
		//console.log("paint");

		var canvas = this.m_oEditor.HtmlElement;
		if (null == canvas)
		{
			this.m_oApi.checkLastWork();
			return;
		}

		var context       = canvas.getContext("2d");
		context.fillStyle = GlobalSkin.BackgroundColor;

		if (AscCommon.AscBrowser.isSailfish)
			context.fillRect(0, 0, canvas.width, canvas.height);
		else if (true === canvas.fullRepaint)
		{
			context.fillRect(0, 0, canvas.width, canvas.height);
			delete canvas.fullRepaint;
		}

		if (this.m_oApi.isDocumentRenderer() || this.m_oDrawingDocument.m_lDrawingFirst < 0 || this.m_oDrawingDocument.m_lDrawingEnd < 0)
			return;

		// сначала посморим, изменились ли ректы страниц
		var rectsPages = [];
		for (var i = this.m_oDrawingDocument.m_lDrawingFirst; i <= this.m_oDrawingDocument.m_lDrawingEnd; i++)
		{
			var drawPage = this.m_oDrawingDocument.m_arrPages[i].drawingPage;

			var _cur_page_rect = new AscCommon._rect();
			_cur_page_rect.x   = (drawPage.left * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
			_cur_page_rect.y   = (drawPage.top * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
			_cur_page_rect.w   = ((drawPage.right * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - _cur_page_rect.x;
			_cur_page_rect.h   = ((drawPage.bottom * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - _cur_page_rect.y;

			rectsPages.push(_cur_page_rect);
		}
		this.m_oBoundsController.CheckPageRects(rectsPages, context);

		if (this.m_oDrawingDocument.m_bIsSelection)
		{
			this.m_oOverlayApi.Clear();
			this.m_oOverlayApi.m_oContext.beginPath();
			this.m_oOverlayApi.m_oContext.fillStyle   = "rgba(51,102,204,255)";
			this.m_oOverlayApi.m_oContext.globalAlpha = 0.2;
		}

		if (this.NoneRepaintPages)
		{
			this.m_bIsFullRepaint = false;

			for (let i = this.m_oDrawingDocument.m_lDrawingFirst; i <= this.m_oDrawingDocument.m_lDrawingEnd; i++)
			{
				let drawPage = this.m_oDrawingDocument.m_arrPages[i].drawingPage;

				let _x = (drawPage.left * AscCommon.AscBrowser.retinaPixelRatio) >> 0;
				let _y = (drawPage.top * AscCommon.AscBrowser.retinaPixelRatio) >> 0;

				this.m_oDrawingDocument.m_arrPages[i].Draw(context,
					_x,
					_y,
					((drawPage.right * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - _x,
					((drawPage.bottom * AscCommon.AscBrowser.retinaPixelRatio) >> 0) - _y,
					this.m_oApi);
			}
		}
		else
		{
			for (let i = 0; i < this.m_oDrawingDocument.m_lDrawingFirst; i++)
				this.m_oDrawingDocument.StopRenderingPage(i);

			for (let i = this.m_oDrawingDocument.m_lDrawingEnd + 1; i < this.m_oDrawingDocument.m_lPagesCount; i++)
				this.m_oDrawingDocument.StopRenderingPage(i);

			for (let i = this.m_oDrawingDocument.m_lDrawingFirst; i <= this.m_oDrawingDocument.m_lDrawingEnd; i++)
			{
				let drawPage = this.m_oDrawingDocument.m_arrPages[i].drawingPage;

				if (this.m_bIsFullRepaint === true)
				{
					this.m_oDrawingDocument.StopRenderingPage(i);
				}

				var __x = drawPage.left;
				var __y = drawPage.top;
				var __w = drawPage.right - __x;
				var __h = drawPage.bottom - __y;

				__x = AscCommon.AscBrowser.convertToRetinaValue(__x, true);
				__y = AscCommon.AscBrowser.convertToRetinaValue(__y, true);
				__w = AscCommon.AscBrowser.convertToRetinaValue(__w, true);
				__h = AscCommon.AscBrowser.convertToRetinaValue(__h, true);

				this.m_oDrawingDocument.CheckRecalculatePage(__w, __h, i);
				if (null == drawPage.cachedImage)
				{
					this.m_oDrawingDocument.StartRenderingPage(i);
				}

				this.m_oDrawingDocument.m_arrPages[i].Draw(context, __x, __y, __w, __h, this.m_oApi);
			}
		}

		this.m_bIsFullRepaint = false;

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
			this.m_oDrawingDocument.UpdateTargetNoAttack();
			this.m_bIsUpdateTargetNoAttack = false;
		}
		this.m_oApi.clearEyedropperImgData();
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

	this.GetDrawingPageInfo = function(nPageIndex)
	{
		return this.m_oDrawingDocument.m_arrPages[nPageIndex];
	};

	this.CalculateDocumentSize = function()
	{
		this.m_dDocumentWidth      = 0;
		this.m_dDocumentHeight     = 0;
		this.m_dDocumentPageWidth  = 0;
		this.m_dDocumentPageHeight = 0;

		var dKoef = (this.m_nZoomValue * g_dKoef_mm_to_pix / 100);

		for (var i = 0; i < this.m_oDrawingDocument.m_lPagesCount; i++)
		{
			var mm_w = this.m_oDrawingDocument.m_arrPages[i].width_mm;
			var mm_h = this.m_oDrawingDocument.m_arrPages[i].height_mm;

			if (mm_w > this.m_dDocumentPageWidth)
				this.m_dDocumentPageWidth = mm_w;
			if (mm_h > this.m_dDocumentPageHeight)
				this.m_dDocumentPageHeight = mm_h;

			var _pageWidth  = (mm_w * dKoef) >> 0;
			var _pageHeight = (mm_h * dKoef + 0.5) >> 0;

			if (_pageWidth > this.m_dDocumentWidth)
				this.m_dDocumentWidth = _pageWidth;

			this.m_dDocumentHeight += 20;
			this.m_dDocumentHeight += _pageHeight;
		}

		this.m_dDocumentHeight += 20;

		// теперь увеличим ширину документа, чтобы он не был плотно к краям
		if (!this.m_oApi.isMobileVersion)
			this.m_dDocumentWidth += 40;

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

		// теперь нужно выставить размеры для скроллов
		this.checkNeedHorScroll();

		this.UpdateScrolls();

		if (this.MobileTouchManager)
			this.MobileTouchManager.Resize();

		if (this.ReaderTouchManager)
			this.ReaderTouchManager.Resize();

		if (this.m_bIsRePaintOnScroll === true)
			this.OnScroll();
	};

	this.OnCalculatePagesPlace = function()
	{
		if (this.m_oApi.isUseNativeViewer && this.m_oDrawingDocument.m_oDocumentRenderer)
		{
			// там все свое
			return;
		}

		if (this.MobileTouchManager && !this.MobileTouchManager.IsWorkedPosition())
			this.MobileTouchManager.ClearContextMenu();

		var canvas = this.m_oEditor.HtmlElement;
		if (null == canvas)
			return;

		var _width  = canvas.width;
		var _height = canvas.height;

		_width = AscCommon.AscBrowser.convertToRetinaValue(_width);
		_height = AscCommon.AscBrowser.convertToRetinaValue(_height);

		var bIsFoundFirst = false;
		var bIsFoundEnd   = false;

		var hor_pos_median = parseInt(_width / 2);
		if (0 != this.m_dScrollX || (this.m_dDocumentWidth > _width))
		{
			//var part = this.m_dScrollX / Math.max(this.m_dScrollX_max, 1);
			//hor_pos_median = parseInt(this.m_dDocumentWidth / 2 + part * (_width - this.m_dDocumentWidth));
			hor_pos_median = parseInt(this.m_dDocumentWidth / 2 - this.m_dScrollX);
		}

		let lCurrentTopInDoc = parseInt(this.m_dScrollY);
		//let offsetTop = AscCommon.AscBrowser.convertToRetinaValue(this.offsetTop, true);
		let offsetTop = this.offsetTop;

		var dKoef  = (this.m_nZoomValue * g_dKoef_mm_to_pix / 100);
		var lStart = offsetTop;
		for (var i = 0; i < this.m_oDrawingDocument.m_lPagesCount; i++)
		{
			var _pageWidth  = (this.m_oDrawingDocument.m_arrPages[i].width_mm * dKoef + 0.5) >> 0;
			var _pageHeight = (this.m_oDrawingDocument.m_arrPages[i].height_mm * dKoef + 0.5) >> 0;

			if (false === bIsFoundFirst)
			{
				if (lStart + 20 + _pageHeight > lCurrentTopInDoc)
				{
					this.m_oDrawingDocument.m_lDrawingFirst = i;
					bIsFoundFirst                           = true;
				}
			}

			var xDst = hor_pos_median - parseInt(_pageWidth / 2);
			var wDst = _pageWidth;
			var yDst = lStart + 20 - lCurrentTopInDoc;
			var hDst = _pageHeight;

			if (false === bIsFoundEnd)
			{
				if (yDst > _height)
				{
					this.m_oDrawingDocument.m_lDrawingEnd = i - 1;
					bIsFoundEnd                           = true;
				}
			}

			var drawRect = this.m_oDrawingDocument.m_arrPages[i].drawingPage;

			drawRect.left      = xDst;
			drawRect.top       = yDst;
			drawRect.right     = xDst + wDst;
			drawRect.bottom    = yDst + hDst;
			drawRect.pageIndex = i;

			lStart += (20 + _pageHeight);
		}

		if (false === bIsFoundEnd)
		{
			this.m_oDrawingDocument.m_lDrawingEnd = this.m_oDrawingDocument.m_lPagesCount - 1;
		}

		if ((-1 == this.m_oDrawingDocument.m_lPagesCount) && (0 != this.m_oDrawingDocument.m_lPagesCount))
		{
			this.m_oDrawingDocument.m_lCurrentPage = 0;
			this.SetCurrentPage();
		}

		// отправляем евент о текущей странице. только в мобильной версии
		if ((this.m_oApi.isMobileVersion || this.m_oApi.isViewMode) && (!window["NATIVE_EDITOR_ENJINE"]))
		{
			var lPage = this.m_oApi.GetCurrentVisiblePage();
			this.m_oApi.sendEvent("asc_onCurrentVisiblePage", this.m_oApi.GetCurrentVisiblePage());

			if (null != this.m_oDrawingDocument.m_oDocumentRenderer)
			{
				this.m_oDrawingDocument.m_lCurrentPage = lPage;
				this.m_oApi.sendEvent("asc_onCurrentPage", lPage);
			}
		}

		if (this.m_bDocumentPlaceChangedEnabled)
			this.m_oApi.sendEvent("asc_onDocumentPlaceChanged");
	};

	this.OnResize = function(isAttack)
	{
		AscBrowser.checkZoom();

		var isNewSize = this.checkBodySize();
		if (!isNewSize && this.retinaScaling === AscCommon.AscBrowser.retinaPixelRatio && false === isAttack)
			return;

		this.m_nZoomValueMin = -1;
		this.m_nZoomValueMax = -1;
		if (this.MobileTouchManager)
			this.MobileTouchManager.Resize_Before();

		if (this.m_oApi.printPreview)
			this.m_oApi.printPreview.resize();

		this.CheckRetinaDisplay();
		this.m_oBody.Resize(this.Width * g_dKoef_pix_to_mm, this.Height * g_dKoef_pix_to_mm, this);
		this.onButtonTabsDraw();

		let areaParent = "id_main_view";
		if (this.m_oApi.isUseNativeViewer)
		{
			areaParent = "id_main";
			var oViewer = this.m_oDrawingDocument.m_oDocumentRenderer;
			if (oViewer)
				oViewer.resize();
		}

		if (AscCommon.g_inputContext)
			AscCommon.g_inputContext.onResize(areaParent);

		if (this.TextBoxBackground != null)
		{
			// это мега заглушка. чтобы не показывалась клавиатура при тыкании на тулбар
			this.TextBoxBackground.HtmlElement.style.top = "10px";
		}

		if (this.checkNeedHorScroll())
			return;

		// теперь проверим необходимость перезуммирования
		if (1 == this.m_nZoomType)
		{
			if (true === this.zoom_FitToWidth())
			{
				this.m_oBoundsController.ClearNoAttack();
				this.onTimerScroll_sync();
				return;
			}
		}
		if (2 == this.m_nZoomType)
		{
			if (true === this.zoom_FitToPage())
			{
				this.m_oBoundsController.ClearNoAttack();
				this.onTimerScroll_sync();
				return;
			}
		}

		this.m_bIsUpdateHorRuler = true;
		this.m_bIsUpdateVerRuler = true;

		if (this.m_bIsRuler)
		{
			this.UpdateHorRulerBack(true);
			this.UpdateVerRulerBack(true);
		}

		this.m_oHorRuler.RepaintChecker.BlitAttack = true;
		this.m_oVerRuler.RepaintChecker.BlitAttack = true;

		this.UpdateScrolls();

		if (this.MobileTouchManager)
			this.MobileTouchManager.Resize();

		if (this.ReaderTouchManager)
			this.ReaderTouchManager.Resize();

		this.m_bIsUpdateTargetNoAttack = true;
		this.m_bIsRePaintOnScroll      = true;

		this.m_oBoundsController.ClearNoAttack();

		this.OnScroll();
		this.onTimerScroll_sync();

		if (this.MobileTouchManager)
			this.MobileTouchManager.Resize_After();

		if (AscCommon.g_imageControlsStorage)
			AscCommon.g_imageControlsStorage.resize();
	};

	// overlay ---
	this.StartUpdateOverlay = function()
	{
		this.IsUpdateOverlayOnlyEnd = true;
	};
	this.EndUpdateOverlay   = function()
	{
		if (this.IsUpdateOverlayOnlyEndReturn)
			return;

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

		//console.log("update_overlay");
		var overlay = this.m_oOverlayApi;
		//if (!overlay.m_bIsShow)
		//    return;

		overlay.Clear();
		var ctx = overlay.m_oContext;

		var drDoc = this.m_oDrawingDocument;
		drDoc.SelectionMatrix = null;
		if (drDoc.m_lDrawingFirst < 0 || drDoc.m_lDrawingEnd < 0)
			return true;

		if (drDoc.m_bIsSearching)
		{
			ctx.fillStyle = "rgba(255,200,0,1)";
			ctx.beginPath();

			var drDoc = this.m_oDrawingDocument;
			for (var i = drDoc.m_lDrawingFirst; i <= drDoc.m_lDrawingEnd; i++)
			{
				var drawPage                  = drDoc.m_arrPages[i].drawingPage;
				drDoc.m_arrPages[i].pageIndex = i;
				drDoc.m_arrPages[i].DrawSearch(overlay, drawPage.left, drawPage.top, drawPage.right - drawPage.left, drawPage.bottom - drawPage.top, drDoc);
			}

			ctx.globalAlpha = 0.5;
			ctx.fill();
			ctx.beginPath();
			ctx.globalAlpha = 1.0;
		}

		if (null == drDoc.m_oDocumentRenderer)
		{
			if (drDoc.m_bIsSelection)
			{
				this.CheckShowOverlay();
				drDoc.private_StartDrawSelection(overlay);

				if (!this.MobileTouchManager)
				{
					for (var i = drDoc.m_lDrawingFirst; i <= drDoc.m_lDrawingEnd; i++)
					{
						if (!drDoc.IsFreezePage(i))
							this.m_oLogicDocument.DrawSelectionOnPage(i);
					}
				}
				else
				{
					for (var i = 0; i <= drDoc.m_lPagesCount; i++)
					{
						if (!drDoc.IsFreezePage(i))
							this.m_oLogicDocument.DrawSelectionOnPage(i);
					}
				}

				drDoc.private_EndDrawSelection();

				drDoc.DrawPageSelection2(overlay);

				if (this.MobileTouchManager)
					this.MobileTouchManager.CheckSelect(overlay);
			}

			if (this.MobileTouchManager)
				this.MobileTouchManager.CheckTableRules(overlay);

			ctx.globalAlpha = 1.0;

			// drawShapes (+ track)
			if (this.m_oLogicDocument.DrawingObjects)
			{
				for (var indP = drDoc.m_lDrawingFirst; indP <= drDoc.m_lDrawingEnd; indP++)
				{
					this.m_oDrawingDocument.AutoShapesTrack.PageIndex = indP;
					this.m_oLogicDocument.DrawingObjects.drawSelect(indP);
				}

				if (this.m_oLogicDocument.DrawingObjects.needUpdateOverlay())
				{
					overlay.Show();
					this.m_oDrawingDocument.AutoShapesTrack.PageIndex = -1;
					this.m_oLogicDocument.DrawingObjects.drawOnOverlay(this.m_oDrawingDocument.AutoShapesTrack);
					this.m_oDrawingDocument.AutoShapesTrack.CorrectOverlayBounds();
				}
			}

			var _table_outline = drDoc.TableOutlineDr.TableOutline;
			let isTouchMode = this.MobileTouchManager ? this.MobileTouchManager.isTouchMode() : false;
			if (_table_outline != null && !isTouchMode)
			{
				var _page = _table_outline.PageNum;
				if (_page >= drDoc.m_lDrawingFirst && _page <= drDoc.m_lDrawingEnd)
				{
					var drawPage = drDoc.m_arrPages[_page].drawingPage;
					drDoc.m_arrPages[_page].DrawTableOutline(overlay,
						drawPage.left, drawPage.top, drawPage.right - drawPage.left, drawPage.bottom - drawPage.top, drDoc.TableOutlineDr);
				}
				if (true)
				{
					var _lastBounds = drDoc.TableOutlineDr.getLastPageBounds();
					_page = _lastBounds.Page;
					if (_page >= drDoc.m_lDrawingFirst && _page <= drDoc.m_lDrawingEnd)
					{
						var drawPage = drDoc.m_arrPages[_page].drawingPage;
						drDoc.m_arrPages[_page].DrawTableOutline(overlay,
							drawPage.left, drawPage.top, drawPage.right - drawPage.left, drawPage.bottom - drawPage.top, drDoc.TableOutlineDr, _lastBounds);
					}
				}
			}

			drDoc.contentControls && drDoc.contentControls.DrawContentControlsTrack(overlay);

			if (drDoc.placeholders.objects.length > 0)
			{
				for (var indP = drDoc.m_lDrawingFirst; indP <= drDoc.m_lDrawingEnd; indP++)
				{
					const oPage = drDoc.m_arrPages[indP];
					const oPixelRect = {};
					oPixelRect.left = AscCommon.AscBrowser.convertToRetinaValue(oPage.drawingPage.left, true);
					oPixelRect.right = AscCommon.AscBrowser.convertToRetinaValue(oPage.drawingPage.right, true);
					oPixelRect.top = AscCommon.AscBrowser.convertToRetinaValue(oPage.drawingPage.top, true);
					oPixelRect.bottom = AscCommon.AscBrowser.convertToRetinaValue(oPage.drawingPage.bottom, true);
					drDoc.placeholders.draw(overlay, indP, oPixelRect, oPage.width_mm, oPage.height_mm);
				}
			}

			if (drDoc.TableOutlineDr.bIsTracked)
			{
				drDoc.DrawTableTrack(overlay);
			}

			if (drDoc.FrameRect.IsActive)
			{
				drDoc.DrawFrameTrack(overlay);
			}

			if (drDoc.MathTrack.IsActive())
			{
				drDoc.DrawMathTrack(overlay);
			}

			if (drDoc.FieldTrack.IsActive)
			{
				drDoc.DrawFieldTrack(overlay);
			}

			if (drDoc.InlineTextTrackEnabled && null != drDoc.InlineTextTrack)
			{
				var _oldPage        = drDoc.AutoShapesTrack.PageIndex;
				var _oldCurPageInfo = drDoc.AutoShapesTrack.CurrentPageInfo;

				drDoc.AutoShapesTrack.PageIndex = drDoc.InlineTextTrackPage;
				drDoc.AutoShapesTrack.DrawInlineMoveCursor(drDoc.InlineTextTrack.X, drDoc.InlineTextTrack.Y, drDoc.InlineTextTrack.Height, drDoc.InlineTextTrack.transform);

				drDoc.AutoShapesTrack.PageIndex       = _oldPage;
				drDoc.AutoShapesTrack.CurrentPageInfo = _oldCurPageInfo;
			}

			if (this.m_oApi.isDrawTablePen || this.m_oApi.isDrawTableErase)
			{
				var logicObj = this.m_oLogicDocument.DrawTableMode;
				if (logicObj.Start)
				{
					var drObject = null;
					if (logicObj.Table)
						drObject = logicObj.Table.GetDrawLine(logicObj.StartX, logicObj.StartY,
							logicObj.EndX, logicObj.EndY,
							logicObj.TablePageStart, logicObj.TablePageEnd, this.m_oApi.isDrawTablePen);
					drDoc.DrawCustomTableMode(overlay, drObject, logicObj, this.m_oApi.isDrawTablePen);
				}
			}

			drDoc.DrawHorVerAnchor();
		}
		else
		{
			drDoc.m_oDocumentRenderer.onUpdateOverlay();
		}

		if (this.MobileTouchManager)
		{
			let targetElement = (this.m_oDrawingDocument && this.m_oDrawingDocument.isDrawTargetGlass()) ? this.m_oDrawingDocument.TargetHtmlElement : null;
			this.MobileTouchManager.CheckGlass(overlay, this.m_oEditor.HtmlElement, targetElement);
		}
	};

	this.OnScroll       = function(isFromLA)
	{
		if (isFromLA)
			this.IsRepaintOnCallbackLongAction = false;

		this.OnCalculatePagesPlace();
		this.m_bIsScroll = true;
	};

	///

	this.ToSearchResult = function()
	{
		var naviG = this.m_oDrawingDocument.CurrentSearchNavi;

		var navi = naviG[0];
		var x    = navi.X;
		var y    = navi.Y;

		var type    = (naviG.Type & 0xFF);
		var PageNum = navi.PageNum;

		if (navi.Transform)
		{
			var xx = navi.Transform.TransformPointX(x, y);
			var yy = navi.Transform.TransformPointY(x, y);

			x = xx;
			y = yy;
		}

		var rectH = (navi.H * this.m_nZoomValue * g_dKoef_mm_to_pix / 100);
		var pos = this.m_oDrawingDocument.ConvertCoordsToCursor2(x, y, PageNum);

		if (true === pos.Error)
			return;

		var boxX = 0;
		var boxY = 0;
		var boxR = ((this.m_oEditor.HtmlElement.width / AscCommon.AscBrowser.retinaPixelRatio) >> 0) - 2;
		var boxB = ((this.m_oEditor.HtmlElement.height / AscCommon.AscBrowser.retinaPixelRatio) >> 0) - rectH;

		var offsetBorder = 20;

		var nValueScrollHor = 0;
		if (pos.X < boxX)
		{
			nValueScrollHor = pos.X - boxX - offsetBorder;
		}
		if (pos.X > boxR)
		{
			nValueScrollHor = pos.X - boxR + offsetBorder;
		}

		var nValueScrollVer = 0;
		if (pos.Y < boxY)
		{
			nValueScrollVer = pos.Y - boxY - offsetBorder;
		}
		if (pos.Y > boxB)
		{
			nValueScrollVer = pos.Y - boxB + offsetBorder;
		}

		var isNeedScroll = false;
		if (0 != nValueScrollHor)
		{
			isNeedScroll                   = true;
			this.m_bIsUpdateTargetNoAttack = true;
			this.m_oScrollHorApi.scrollByX(nValueScrollHor);
		}
		if (0 != nValueScrollVer)
		{
			isNeedScroll                   = true;
			this.m_bIsUpdateTargetNoAttack = true;
			this.m_oScrollVerApi.scrollByY(nValueScrollVer);
		}

		if (true === isNeedScroll)
		{
			this.OnScroll();
			return;
		}

		// и, в самом конце, перерисовываем оверлей
		this.OnUpdateOverlay();
	};

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
		if (this.m_oApi.isDrawTablePen || this.m_oApi.isDrawTableErase || this.m_oApi.isInkDrawerOn())
			return;
		this.m_oDrawingDocument.UnlockCursorType();
	};

	this.TransformDivUseAnimation = function(_div, topPos)
	{
		_div.style[window.asc_sdk_transitionProperty] = "top";
		_div.style[window.asc_sdk_transitionDuration] = "1000ms";
		_div.style.top                                = topPos + "px";
	};

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
		oThis.m_oLogicDocument.ContinueTrackRevisions();
	};

	this.ChangeHintProps = function()
	{
		var bFlag = false;
		if (global_keyboardEvent.CtrlKey)
		{
			if (null != this.m_oLogicDocument)
			{
				if (49 == global_keyboardEvent.KeyCode)
				{
                    AscCommon.g_fontManager.SetHintsProps(false, false);
					bFlag = true;
				}
				else if (50 == global_keyboardEvent.KeyCode)
				{
                    AscCommon.g_fontManager.SetHintsProps(true, false);
					bFlag = true;
				}
				else if (51 == global_keyboardEvent.KeyCode)
				{
                    AscCommon.g_fontManager.SetHintsProps(true, true);
					bFlag = true;
				}
			}
		}

		if (bFlag)
		{
			this.m_oDrawingDocument.ClearCachePages();

			if (AscCommon.g_fontManager2)
				AscCommon.g_fontManager2.ClearFontsRasterCache();
		}

		return bFlag;
	};

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

		this.m_oHorRuler.CreateBackground(this.m_oDrawingDocument.m_arrPages[0]);
		this.m_oVerRuler.CreateBackground(this.m_oDrawingDocument.m_arrPages[0]);
		this.UpdateHorRuler();
		this.UpdateVerRuler();
	};

	this.InitControl = function()
	{
		if (this.IsInitControl)
			return;

		this.CalculateDocumentSize();

		if (!this.m_oApi.isOnlyReaderMode)
			this.StartMainTimer();

		this.m_oHorRuler.CreateBackground(this.m_oDrawingDocument.m_arrPages[0]);
		this.m_oVerRuler.CreateBackground(this.m_oDrawingDocument.m_arrPages[0]);
		this.UpdateHorRuler();
		this.UpdateVerRuler();

        if (!this.m_oApi.isPdfEditor())
        {
            AscCommon.InitBrowserInputContext(this.m_oApi, "id_target_cursor");
            if (AscCommon.g_inputContext)
                AscCommon.g_inputContext.onResize("id_main_view");

            this.initEventsMobile();
        }

		if (undefined !== this.m_oApi.startMobileOffset)
		{
			this.setOffsetTop(this.m_oApi.startMobileOffset.offset, this.m_oApi.startMobileOffset.offsetScrollTop);
			delete this.m_oApi.startMobileOffset;
		}

		this.IsInitControl = true;
	};

	// current page ---
	this.SetCurrentPage  = function(isNoUpdateRulers)
	{
		var drDoc = this.m_oDrawingDocument;

		if (isNoUpdateRulers === undefined)
		{
			if (0 <= drDoc.m_lCurrentPage && drDoc.m_lCurrentPage < drDoc.m_lPagesCount)
			{
				this.m_oHorRuler.CreateBackground(drDoc.m_arrPages[drDoc.m_lCurrentPage]);
				this.m_oVerRuler.CreateBackground(drDoc.m_arrPages[drDoc.m_lCurrentPage]);

				this.m_oHorRuler.IsCanMoveMargins = true;
				this.m_oVerRuler.IsCanMoveMargins = true;
			}
		}

		this.m_bIsUpdateHorRuler = true;
		this.m_bIsUpdateVerRuler = true;

		this.OnScroll();

		this.m_oApi.sync_currentPageCallback(drDoc.m_lCurrentPage);
	};

	this.SetCurrentPage2 = function()
	{
		var drDoc = this.m_oDrawingDocument;
		if (0 <= drDoc.m_lCurrentPage && drDoc.m_lCurrentPage < drDoc.m_lPagesCount)
		{
			this.m_oHorRuler.CreateBackground(drDoc.m_arrPages[drDoc.m_lCurrentPage]);
			this.m_oVerRuler.CreateBackground(drDoc.m_arrPages[drDoc.m_lCurrentPage]);
		}

		this.m_bIsUpdateHorRuler = true;
		this.m_bIsUpdateVerRuler = true;

		this.m_oApi.sync_currentPageCallback(drDoc.m_lCurrentPage);
	};

	this.GoToPage = function(lPageNum)
	{
		var drDoc = this.m_oDrawingDocument;
		if (lPageNum < 0 || lPageNum >= drDoc.m_lPagesCount)
			return;

		// сначала вычислим место для скролла
		var dKoef = g_dKoef_mm_to_pix * this.m_nZoomValue / 100;

		var lYPos = 0;
		for (var i = 0; i < lPageNum; i++)
		{
			lYPos += (20 + parseInt(this.m_oDrawingDocument.m_arrPages[i].height_mm * dKoef));
		}

		drDoc.m_lCurrentPage = lPageNum;
		this.m_oHorRuler.CreateBackground(drDoc.m_arrPages[drDoc.m_lCurrentPage]);
		this.m_oVerRuler.CreateBackground(drDoc.m_arrPages[drDoc.m_lCurrentPage]);

		this.m_bIsUpdateHorRuler = true;
		this.m_bIsUpdateVerRuler = true;

		if (this.m_dDocumentHeight > (this.m_oEditor.HtmlElement.height + 10))
		{
			var y = lYPos * this.m_dScrollY_max / (this.m_dDocumentHeight - this.m_oEditor.HtmlElement.height);
			this.m_oScrollVerApi.scrollTo(0, y + 1);
		}

		if (this.m_oApi.isViewMode === false && null != this.m_oLogicDocument)
		{
			if (false === drDoc.IsFreezePage(drDoc.m_lCurrentPage))
			{
				this.m_oLogicDocument.GoToPage(drDoc.m_lCurrentPage);
				this.m_oApi.sync_currentPageCallback(drDoc.m_lCurrentPage);
			}
		}
		else
		{
			this.m_oApi.sync_currentPageCallback(drDoc.m_lCurrentPage);
		}
	};

	this.GetMainContentBounds = function()
	{
		return this.m_oMainContent.AbsolutePosition;
	};
	this.GetVertRulerLeft = function()
	{
		return this.m_oLeftRuler.AbsolutePosition.L;
	};

	// mobile ---
	this.setOffsetTop = function(offset, offsetScrollTop)
	{
		if (this.m_oDrawingDocument.m_oDocumentRenderer &&
			offset !== undefined &&
			offset !== this.m_oDrawingDocument.m_oDocumentRenderer.offsetTop)
		{
			this.m_oDrawingDocument.m_oDocumentRenderer.setOffsetTop(offset);
		}
		else if (offset !== undefined && offset !== this.offsetTop)
		{
			this.offsetTop = offset;

			this.UpdateScrolls();

			if (this.MobileTouchManager)
				this.MobileTouchManager.UpdateScrolls();

			this.OnScroll();
		}

		if (undefined !== offsetScrollTop && this.MobileTouchManager)
			this.MobileTouchManager.iScroll.setOffsetTop(offsetScrollTop);
	};
}

//------------------------------------------------------------export----------------------------------------------------
window['AscCommon']                      = window['AscCommon'] || {};
window['AscCommonWord']                  = window['AscCommonWord'] || {};
window['AscCommonWord'].CEditorPage      = CEditorPage;

window['AscCommon'].Page_Width      = Page_Width;
window['AscCommon'].Page_Height     = Page_Height;
window['AscCommon'].X_Left_Margin   = X_Left_Margin;
window['AscCommon'].X_Right_Margin  = X_Right_Margin;
window['AscCommon'].Y_Bottom_Margin = Y_Bottom_Margin;
window['AscCommon'].Y_Top_Margin    = Y_Top_Margin;
