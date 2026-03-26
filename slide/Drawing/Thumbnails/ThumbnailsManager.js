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

(function(undefined) {

	function CThumbnailsManagerBase(editorPage) {
		this.m_oWordControl = editorPage;
	}
	CThumbnailsManagerBase.prototype.initEvents = function()
	{
		var control = this.m_oWordControl.m_oThumbnails.HtmlElement;
		const oThis = this;
		AscCommon.addMouseEvent(control, "down", function () {
			const thumbnails = oThis.m_oWordControl.Thumbnails;
			return thumbnails.onMouseDown.apply(thumbnails, arguments);
		});
		AscCommon.addMouseEvent(control, "move", function () {
			const thumbnails = oThis.m_oWordControl.Thumbnails;
			return thumbnails.onMouseMove.apply(thumbnails, arguments);
		});
		AscCommon.addMouseEvent(control, "up", function () {
			const thumbnails = oThis.m_oWordControl.Thumbnails;
			return thumbnails.onMouseUp.apply(thumbnails, arguments);
		});

		control.onmouseout = function () {
			const thumbnails = oThis.m_oWordControl.Thumbnails;
			return thumbnails.onMouseLeave.apply(thumbnails, arguments);
		};

		const onmousewheel = function () {
			const thumbnails = oThis.m_oWordControl.Thumbnails;
			return thumbnails.onMouseWhell.apply(thumbnails, arguments);
		};
		control.onmousewheel = onmousewheel;
		if (control.addEventListener)
		{
			control.addEventListener("DOMMouseScroll", onmousewheel, false);
		}
	};
	CThumbnailsManagerBase.prototype.isThumbnailsShown = function () {
		const thumbnailsContainer = this.m_oWordControl.m_oThumbnailsContainer;
		if (!thumbnailsContainer) return false;
		if (!this.m_oWordControl.IsThumbnailsSupported()) {
			return false;
		}

		const absolutePosition = thumbnailsContainer.AbsolutePosition;
		const width = absolutePosition.R - absolutePosition.L;
		const height = absolutePosition.B - absolutePosition.T;

		const isHorizontalThumbnails = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;
		return isHorizontalThumbnails
			? width >= 0 && height >= 1
			: width >= 1 && height >= 0;
	};


	function CThumbnailsManager(editorPage) {
		CThumbnailsManagerBase.call(this, editorPage);
		this.isInit = false;
		this.lastPixelRatio = 0;
		this.m_oFontManager = new AscFonts.CFontManager();

		this.m_bIsScrollVisible = true;

		// cached measure
		this.DigitWidths = [];

		// skin
		this.backgroundColor = "#B0B0B0";

		this.const_offset_x = 0;
		this.const_offset_y = 0;
		this.const_offset_r = 4;
		this.const_offset_b = 0;
		this.const_border_w = 4;

		// size & position
		this.SlideWidth = 297;
		this.SlideHeight = 210;

		this.m_dScrollY = 0;
		this.m_dScrollY_max = 0;

		this.m_bIsUpdate = false;

		//regular mode
		this.thumbnails = new AscCommon.CSlidesThumbnails();

		this.m_nCurrentPage = -1;
		this.m_arrPages = [];
		this.m_lDrawingFirst = -1;
		this.m_lDrawingEnd = -1;

		this.bIsEmptyDrawed = false;

		this.m_oCacheManager = new CCacheManager(true);

		this.FocusObjType = FOCUS_OBJECT_MAIN;
		this.LockMainObjType = false;

		this.SelectPageEnabled = true;

		this.MouseDownTrack = new AscCommon.CMouseDownTrack(this);

		this.MouseTrackCommonImage = null;

		this.MouseThumbnailsAnimateScrollTopTimer = -1;
		this.MouseThumbnailsAnimateScrollBottomTimer = -1;

		this.ScrollerHeight = 0;
		this.ScrollerWidth = 0;
	}
	AscFormat.InitClassWithoutType(CThumbnailsManager, CThumbnailsManagerBase);
	CThumbnailsManager.prototype.Init = function()
	{
		this.isInit = true;
		this.m_oFontManager.Initialize(true);
		this.m_oFontManager.SetHintsProps(true, true);

		this.InitCheckOnResize();

		this.MouseTrackCommonImage = document.createElement("canvas");

		var _im_w = 9;
		var _im_h = 9;

		this.MouseTrackCommonImage.width = _im_w;
		this.MouseTrackCommonImage.height = _im_h;

		var _ctx = this.MouseTrackCommonImage.getContext('2d');
		var _data = _ctx.createImageData(_im_w, _im_h);
		var _pixels = _data.data;

		var _ind = 0;
		for (var j = 0; j < _im_h; ++j)
		{
			var _off1 = (j > (_im_w >> 1)) ? (_im_w - j - 1) : j;
			var _off2 = _im_w - _off1 - 1;

			for (var r = 0; r < _im_w; ++r)
			{
				if (r <= _off1 || r >= _off2)
				{
					_pixels[_ind++] = 183;
					_pixels[_ind++] = 183;
					_pixels[_ind++] = 183;
					_pixels[_ind++] = 255;
				} else
				{
					_pixels[_ind + 3] = 0;
					_ind += 4;
				}
			}
		}

		_ctx.putImageData(_data, 0, 0);
	};

	CThumbnailsManager.prototype.InitCheckOnResize = function()
	{
		if (!this.isInit)
			return;

		if (Math.abs(this.lastPixelRatio - AscCommon.AscBrowser.retinaPixelRatio) < 0.01)
			return;

		this.lastPixelRatio = AscCommon.AscBrowser.retinaPixelRatio;
		this.SetFont({
			FontFamily: {Name: "Arial", Index: -1},
			Italic: false,
			Bold: false,
			FontSize: Math.round(10 * AscCommon.AscBrowser.retinaPixelRatio)
		});

		// измеряем все цифры
		for (var i = 0; i < 10; i++)
		{
			var _meas = this.m_oFontManager.MeasureChar(("" + i).charCodeAt(0));
			if (_meas)
				this.DigitWidths[i] = _meas.fAdvanceX * 25.4 / 96;
			else
				this.DigitWidths[i] = 10;
		}

		if (Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom) {
			this.const_offset_x = Math.round(this.lastPixelRatio * 17);
			this.const_offset_r = this.const_offset_y;
			this.const_offset_b = Math.round(this.lastPixelRatio * 10);
			this.const_border_w = Math.round(this.lastPixelRatio * 7);
		} else {
			this.const_offset_y = Math.round(this.lastPixelRatio * 17);
			this.const_offset_b = this.const_offset_y;
			this.const_offset_r = Math.round(this.lastPixelRatio * 10);
			this.const_border_w = Math.round(this.lastPixelRatio * 7);
		}

	};

	CThumbnailsManager.prototype.GetPageByPos = function(oPos)
	{
		return this.thumbnails.GetPage(oPos);
	};

	CThumbnailsManager.prototype.GetSlidesCount = function()
	{
		return Asc.editor.getCountSlides();
	};

	CThumbnailsManager.prototype.SetFont = function(font)
	{
		font.FontFamily.Name = g_fontApplication.GetFontFileWeb(font.FontFamily.Name, 0).m_wsFontName;

		if (-1 == font.FontFamily.Index)
			font.FontFamily.Index = AscFonts.g_map_font_index[font.FontFamily.Name];

		if (font.FontFamily.Index == undefined || font.FontFamily.Index == -1)
			return;

		var bItalic = true === font.Italic;
		var bBold   = true === font.Bold;

		var oFontStyle = FontStyle.FontStyleRegular;
		if (!bItalic && bBold)
			oFontStyle = FontStyle.FontStyleBold;
		else if (bItalic && !bBold)
			oFontStyle = FontStyle.FontStyleItalic;
		else if (bItalic && bBold)
			oFontStyle = FontStyle.FontStyleBoldItalic;

		g_fontApplication.LoadFont(font.FontFamily.Name, AscCommon.g_font_loader, this.m_oFontManager, font.FontSize, oFontStyle, 96, 96);
	};

	CThumbnailsManager.prototype.SetSlideRecalc = function (nIdx)
	{
		this.thumbnails.SetSlideRecalc(nIdx);
	};

	CThumbnailsManager.prototype.SelectAll = function()
	{
		this.thumbnails.SelectAll();
		this.OnUpdateOverlay();
	};

	CThumbnailsManager.prototype.GetFirstSelectedType = function()
	{
		return this.m_oWordControl.m_oLogicDocument.GetFirstSelectedType();
	};
	CThumbnailsManager.prototype.GetSlideType = function(nIdx)
	{
		return this.m_oWordControl.m_oLogicDocument.GetSlideType(nIdx);
	};
	CThumbnailsManager.prototype.IsMixedSelection = function()
	{
		return this.m_oWordControl.m_oLogicDocument.IsMixedSelection();
	};

	// events
	CThumbnailsManager.prototype.onMouseDown = function (e) {
		const mobileTouchManager = this.m_oWordControl ? this.m_oWordControl.MobileTouchManagerThumbnails : null;
		if (mobileTouchManager && mobileTouchManager.checkTouchEvent(e, true)) {
			mobileTouchManager.startTouchingInProcess();
			const res = mobileTouchManager.mainOnTouchStart(e);
			mobileTouchManager.stopTouchingInProcess();
			return res;
		}

		if (this.m_oWordControl) {
			this.m_oWordControl.m_oApi.checkInterfaceElementBlur();
			this.m_oWordControl.m_oApi.checkLastWork();

			// после fullscreen возможно изменение X, Y после вызова Resize.
			this.m_oWordControl.checkBodyOffset();
		}

		AscCommon.stopEvent(e);

		if (AscCommon.g_inputContext && AscCommon.g_inputContext.externalChangeFocus())
			return;

		var control = this.m_oWordControl.m_oThumbnails.HtmlElement;
		if (global_mouseEvent.IsLocked == true && global_mouseEvent.Sender != control) {
			// кто-то зажал мышку. кто-то другой
			return false;
		}

		AscCommon.check_MouseDownEvent(e);
		global_mouseEvent.LockMouse();
		const oPresentation = this.m_oWordControl.m_oLogicDocument;
		let nStartHistoryIndex = oPresentation.History.Index;
		function checkSelectionEnd() {
			if (nStartHistoryIndex === oPresentation.History.Index) {
				Asc.editor.sendEvent("asc_onSelectionEnd");
			}
		}

		this.m_oWordControl.m_oApi.sync_EndAddShape();
		if (global_mouseEvent.Sender != control) {
			// такого быть не должно
			return false;
		}

		if (global_mouseEvent.Button == undefined)
			global_mouseEvent.Button = 0;

		this.SetFocusElement(FOCUS_OBJECT_THUMBNAILS);

		var pos = this.ConvertCoords(global_mouseEvent.X, global_mouseEvent.Y);
		if (pos.Page == -1) {
			if (global_mouseEvent.Button == 2) {
				this.showContextMenu(false);
			}
			checkSelectionEnd();
			return false;
		}

		if (global_keyboardEvent.CtrlKey && !this.m_oWordControl.m_oApi.isReporterMode) {
			if (this.m_arrPages[pos.Page].IsSelected === true) {
				this.m_arrPages[pos.Page].IsSelected = false;
				var arr = this.GetSelectedArray();
				if (0 == arr.length) {
					this.m_arrPages[pos.Page].IsSelected = true;
					this.ShowPage(pos.Page);
				} else {
					this.OnUpdateOverlay();

					this.SelectPageEnabled = false;
					this.m_oWordControl.GoToPage(arr[0]);
					this.SelectPageEnabled = true;

					this.ShowPage(arr[0]);
				}
			} else {
				if (this.GetFirstSelectedType() === this.GetSlideType(pos.Page)) {
					this.m_arrPages[pos.Page].IsSelected = true;
					this.OnUpdateOverlay();

					this.SelectPageEnabled = false;
					this.m_oWordControl.GoToPage(pos.Page);
					this.SelectPageEnabled = true;
					this.ShowPage(pos.Page);
				}
			}
		} else if (global_keyboardEvent.ShiftKey && !this.m_oWordControl.m_oApi.isReporterMode) {

			var pages_count = this.m_arrPages.length;
			for (var i = 0; i < pages_count; i++) {
				this.m_arrPages[i].IsSelected = false;
			}

			var _max = pos.Page;
			var _min = this.m_oWordControl.m_oDrawingDocument.SlideCurrent;
			if (_min > _max) {
				var _temp = _max;
				_max = _min;
				_min = _temp;
			}

			let nSlideType = this.GetSlideType(_min);
			for (var i = _min; i <= _max; i++) {
				if (nSlideType === this.GetSlideType(i)) {
					this.m_arrPages[i].IsSelected = true;
				}
			}

			this.OnUpdateOverlay();
			this.ShowPage(pos.Page);
			oPresentation.Document_UpdateInterfaceState();
		} else if (0 == global_mouseEvent.Button || 2 == global_mouseEvent.Button) {

			let isMouseDownOnAnimPreview = false;
			if (0 == global_mouseEvent.Button) {
				if (this.m_arrPages[pos.Page].animateLabelRect) // click on the current star, preview animation button, slide transition
				{
					let animateLabelRect = this.m_arrPages[pos.Page].animateLabelRect;
					if (pos.X >= animateLabelRect.minX && pos.X <= animateLabelRect.maxX && pos.Y >= animateLabelRect.minY && pos.Y <= animateLabelRect.maxY)
						isMouseDownOnAnimPreview = true
				}

				if (!isMouseDownOnAnimPreview) // приготавливаемся к треку
				{
					this.MouseDownTrack.Start(pos.Page, global_mouseEvent.X, global_mouseEvent.Y);
				}
			}

			if (this.m_arrPages[pos.Page].IsSelected) {
				let isStartedAnimPreview = (this.m_oWordControl.m_oLogicDocument.IsStartedPreview() || (this.m_oWordControl.m_oDrawingDocument.TransitionSlide && this.m_oWordControl.m_oDrawingDocument.TransitionSlide.IsPlaying()));

				this.SelectPageEnabled = false;
				this.m_oWordControl.GoToPage(pos.Page);
				this.SelectPageEnabled = true;

				if (!isStartedAnimPreview) {
					if (isMouseDownOnAnimPreview) {
						this.m_oWordControl.m_oApi.SlideTransitionPlay(function () { this.m_oWordControl.m_oApi.asc_StartAnimationPreview() });
					}
				}

				if (this.m_oWordControl.m_oNotesApi.IsEmptyDraw) {
					this.m_oWordControl.m_oNotesApi.IsEmptyDraw = false;
					this.m_oWordControl.m_oNotesApi.IsRepaint = true;
				}

				if (global_mouseEvent.Button == 2 && !global_keyboardEvent.CtrlKey) {
					this.showContextMenu(false);
				}
				checkSelectionEnd();
				return false;
			}

			var pages_count = this.m_arrPages.length;
			for (var i = 0; i < pages_count; i++) {
				this.m_arrPages[i].IsSelected = false;
			}

			this.m_arrPages[pos.Page].IsSelected = true;

			this.OnUpdateOverlay();

			if (global_mouseEvent.Button == 0 && this.m_arrPages[pos.Page].animateLabelRect) // click on the current star, preview animation button, slide transition
			{
				let animateLabelRect = this.m_arrPages[pos.Page].animateLabelRect;
				let isMouseDownOnAnimPreview = false;
				if (pos.X >= animateLabelRect.minX && pos.X <= animateLabelRect.maxX && pos.Y >= animateLabelRect.minY && pos.Y <= animateLabelRect.maxY);
				isMouseDownOnAnimPreview = true;
			}

			this.SelectPageEnabled = false;
			this.m_oWordControl.GoToPage(pos.Page);
			this.SelectPageEnabled = true;

			if (isMouseDownOnAnimPreview) {
				this.m_oWordControl.m_oApi.SlideTransitionPlay(function () { this.m_oWordControl.m_oApi.asc_StartAnimationPreview() });
			}

			this.ShowPage(pos.Page);
		}

		if (global_mouseEvent.Button == 2 && !global_keyboardEvent.CtrlKey) {
			this.showContextMenu(false);
		}
		checkSelectionEnd();
		return false;
	};

	CThumbnailsManager.prototype.onMouseMove = function(e)
	{
		let mobileTouchManager = this.m_oWordControl ? this.m_oWordControl.MobileTouchManagerThumbnails : null;
		if (mobileTouchManager && mobileTouchManager.checkTouchEvent(e, true))
		{
			mobileTouchManager.startTouchingInProcess();
			let res = mobileTouchManager.mainOnTouchMove(e);
			mobileTouchManager.stopTouchingInProcess();
			return res;
		}

		const isHorizontalThumbnails = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;

		if (this.m_oWordControl)
			this.m_oWordControl.m_oApi.checkLastWork();

		var control = this.m_oWordControl.m_oThumbnails.HtmlElement;
		if (global_mouseEvent.IsLocked == true && global_mouseEvent.Sender != control)
		{
			// кто-то зажал мышку. кто-то другой
			return;
		}
		AscCommon.check_MouseMoveEvent(e);
		if (global_mouseEvent.Sender != control)
		{
			// такого быть не должно
			return;
		}

		if (this.MouseDownTrack.IsStarted())
		{
			// это трек для перекидывания слайдов
			if (this.MouseDownTrack.IsSimple() && !this.m_oWordControl.m_oApi.isViewMode)
			{
				if (Math.abs(this.MouseDownTrack.GetX() - global_mouseEvent.X) > 10 || Math.abs(this.MouseDownTrack.GetY() - global_mouseEvent.Y) > 10)
					this.MouseDownTrack.ResetSimple(this.ConvertCoords2(global_mouseEvent.X, global_mouseEvent.Y));
			}
			else
			{
				if (!this.MouseDownTrack.IsSimple())
				{
					// нужно определить активная позиция между слайдами
					this.MouseDownTrack.SetPosition(this.ConvertCoords2(global_mouseEvent.X, global_mouseEvent.Y));
				}
			}


			this.OnUpdateOverlay();

			// теперь нужно посмотреть, нужно ли проскроллить
			if (this.m_bIsScrollVisible) {
				const _X = global_mouseEvent.X - this.m_oWordControl.X;
				const _Y = global_mouseEvent.Y - this.m_oWordControl.Y;

				const _abs_pos = this.m_oWordControl.m_oThumbnails.AbsolutePosition;
				const _XMax = (_abs_pos.R - _abs_pos.L) * g_dKoef_mm_to_pix;
				const _YMax = (_abs_pos.B - _abs_pos.T) * g_dKoef_mm_to_pix;

				const mousePos = isHorizontalThumbnails ? _X : _Y;
				const maxPos = isHorizontalThumbnails ? _XMax : _YMax;

				let _check_type = -1;
				if (mousePos < 30) {
					_check_type = 0;
				} else if (mousePos >= (maxPos - 30)) {
					_check_type = 1;
				}

				this.CheckNeedAnimateScrolls(_check_type);
			}

			if (!this.MouseDownTrack.IsSimple())
			{
				var cursor_dragged = "default";
				if (AscCommon.AscBrowser.isWebkit)
					cursor_dragged = "-webkit-grabbing";
				else if (AscCommon.AscBrowser.isMozilla)
					cursor_dragged = "-moz-grabbing";

				this.m_oWordControl.m_oThumbnails.HtmlElement.style.cursor = cursor_dragged;
			}

			return;
		}

		var pos = this.ConvertCoords(global_mouseEvent.X, global_mouseEvent.Y);

		var _is_old_focused = false;

		var pages_count = this.m_arrPages.length;
		for (var i = 0; i < pages_count; i++)
		{
			if (this.m_arrPages[i].IsFocused)
			{
				_is_old_focused = true;
				this.m_arrPages[i].IsFocused = false;
			}
		}

		var cursor_moved = "default";

		if (pos.Page != -1)
		{
			this.m_arrPages[pos.Page].IsFocused = true;
			this.OnUpdateOverlay();

			cursor_moved = "pointer";
		} else if (_is_old_focused)
		{
			this.OnUpdateOverlay();
		}

		this.m_oWordControl.m_oThumbnails.HtmlElement.style.cursor = cursor_moved;
	};

	CThumbnailsManager.prototype.onMouseUp = function(e, bIsWindow)
	{
		let mobileTouchManager = this.m_oWordControl ? this.m_oWordControl.MobileTouchManagerThumbnails : null;
		if (mobileTouchManager && mobileTouchManager.checkTouchEvent(e, true))
		{
			mobileTouchManager.startTouchingInProcess();
			let res = mobileTouchManager.mainOnTouchEnd(e);
			mobileTouchManager.stopTouchingInProcess();
			return res;
		}

		if (this.m_oWordControl)
			this.m_oWordControl.m_oApi.checkLastWork();

		var _oldSender = global_mouseEvent.Sender;
		AscCommon.check_MouseUpEvent(e);
		global_mouseEvent.UnLockMouse();

		var control = this.m_oWordControl.m_oThumbnails.HtmlElement;
		if (global_mouseEvent.Sender != control)
		{
			if (_oldSender != control || true !== bIsWindow)
			{
				// такого быть не должно
				return;
			}
		}

		this.CheckNeedAnimateScrolls(-1);

		if (!this.MouseDownTrack.IsStarted())
			return;

		// теперь смотрим, просто ли это селект, или же это трек
		if (this.MouseDownTrack.IsSimple())
		{
			if (this.MouseDownTrack.IsMoved(global_mouseEvent.X, global_mouseEvent.Y))
				this.MouseDownTrack.ResetSimple(this.ConvertCoords2(global_mouseEvent.X, global_mouseEvent.Y));
		}

		if (this.MouseDownTrack.IsSimple())
		{
			// это просто селект
			var pages_count = this.m_arrPages.length;
			for (var i = 0; i < pages_count; i++)
			{
				this.m_arrPages[i].IsSelected = false;
			}

			this.m_arrPages[this.MouseDownTrack.GetPage()].IsSelected = true;

			this.OnUpdateOverlay();

			// послали уже на mouseDown
			//this.SelectPageEnabled = false;
			//this.m_oWordControl.GoToPage(this.MouseDownTrack.GetPage());
			//this.SelectPageEnabled = true;
		} else
		{
			// это трек
			this.MouseDownTrack.SetPosition(this.ConvertCoords2(global_mouseEvent.X, global_mouseEvent.Y));

			if (-1 !== this.MouseDownTrack.GetPosition() && (!this.MouseDownTrack.IsSamePos() || AscCommon.global_mouseEvent.CtrlKey))
			{
				// вызвать функцию апи для смены слайдов местами
				var _array = this.GetSelectedArray();
				this.m_oWordControl.m_oLogicDocument.shiftSlides(this.MouseDownTrack.GetPosition(), _array);
				this.ClearCacheAttack();
			}

			this.OnUpdateOverlay();
		}
		this.MouseDownTrack.Reset();

		this.onMouseMove(e);
	};

	CThumbnailsManager.prototype.onMouseLeave = function(e)
	{
		var pages_count = this.m_arrPages.length;
		for (var i = 0; i < pages_count; i++)
		{
			this.m_arrPages[i].IsFocused = false;
		}
		this.OnUpdateOverlay();
	};

	CThumbnailsManager.prototype.onMouseWhell = function (e) {
		if (false === this.m_bIsScrollVisible || !this.m_oWordControl.m_oScrollThumbApi)
			return;

		if (global_keyboardEvent.CtrlKey)
			return;

		if (undefined !== window["AscDesktopEditor"] && false === window["AscDesktopEditor"]["CheckNeedWheel"]())
			return;

		const delta = GetWheelDeltaY(e);
		const isHorizontalOrientation = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;
		const isRightToLeft = Asc.editor.isRtlInterface;
		isHorizontalOrientation
			? this.m_oWordControl.m_oScrollThumbApi.scrollBy(isRightToLeft ? -delta : delta, 0)
			: this.m_oWordControl.m_oScrollThumbApi.scrollBy(0, delta, false);

		if (e.preventDefault)
			e.preventDefault();
		else
			e.returnValue = false;

		AscCommon.stopEvent(e);
		return false;
	};

	CThumbnailsManager.prototype.onKeyDown = function(e)
	{
		var control = this.m_oWordControl.m_oThumbnails.HtmlElement;
		if (global_mouseEvent.IsLocked == true && global_mouseEvent.Sender == control)
		{
			e.preventDefault();
			return false;
		}
		AscCommon.check_KeyboardEvent(e);
		let oApi = this.m_oWordControl.m_oApi;
		let oPresentation = this.m_oWordControl.m_oLogicDocument;
		let oDrawingDocument = this.m_oWordControl.m_oDrawingDocument;
		let oEvent = global_keyboardEvent;
		let nShortCutAction = oApi.getShortcut(oEvent);
		let bReturnValue = false, bPreventDefault = true;
		let sSelectedIdx;
		let nStartHistoryIndex = oPresentation.History.Index;
		switch (nShortCutAction)
		{
			case Asc.c_oAscPresentationShortcutType.EditSelectAll:
			{
				this.SelectAll();
				bReturnValue = false;
				bPreventDefault = true;
				break;
			}
			case Asc.c_oAscPresentationShortcutType.EditUndo:
			case Asc.c_oAscPresentationShortcutType.EditRedo:
			{
				bReturnValue = true;
				bPreventDefault = false;
				break;
			}
			case Asc.c_oAscPresentationShortcutType.Duplicate:
			{
				if (oPresentation.CanEdit())
				{
					oApi.DublicateSlide();
				}
				bReturnValue = false;
				bPreventDefault = true;
				break;
			}
			case Asc.c_oAscPresentationShortcutType.PrintPreviewAndPrint:
			{
				oApi.onPrint();
				bReturnValue = false;
				bPreventDefault = true;
				break;
			}
			case Asc.c_oAscPresentationShortcutType.Save:
			{
				if (!oPresentation.IsViewMode())
				{
					oApi.asc_Save();
				}
				bReturnValue = false;
				bPreventDefault = true;
				break;
			}
			case Asc.c_oAscPresentationShortcutType.OpenContextMenu:
			{
				this.showContextMenu(true);
				bReturnValue = false;
				bPreventDefault = false;
				break;
			}
			default:
			{
				break;
			}
		}
		if (!nShortCutAction)
		{
			if (AscCommon.isCopyPasteEvent(oEvent)) {
				return;
			}
			switch (oEvent.KeyCode)
			{
				case 13: // enter
				{
					if (oPresentation.CanEdit())
					{
						sSelectedIdx = this.GetSelectedArray();
						bReturnValue = false;
						bPreventDefault = true;
						if (sSelectedIdx.length > 0)
						{
							this.m_oWordControl.GoToPage(sSelectedIdx[sSelectedIdx.length - 1]);
							oPresentation.addNextSlide();
							bPreventDefault = false;
						}
					}
					break;
				}
				case 46: // delete
				case 8: // backspace
				{
					if (oPresentation.CanEdit())
					{
						const arrSlides = oPresentation.GetSelectedSlideObjects();
						if (!oApi.IsSupportEmptyPresentation)
						{
							if (arrSlides.length === oDrawingDocument.GetSlidesCount())
							{
								arrSlides.splice(0, 1);
							}
						}
						if (arrSlides.length !== 0)
						{
							oPresentation.deleteSlides(arrSlides);
						}
						if (0 === oPresentation.GetSlidesCount())
						{
							this.m_bIsUpdate = true;
						}
						bReturnValue = false;
						bPreventDefault = true;
					}
					break;
				}
				case 34: // PgDown
				case 40: // bottom arrow
				{
					if (oEvent.CtrlKey && oEvent.ShiftKey)
					{
						oPresentation.moveSelectedSlidesToEnd();
					} else if (oEvent.CtrlKey)
					{
						oPresentation.moveSlidesNextPos();
					} else if (oEvent.ShiftKey)
					{
						this.CorrectShiftSelect(false, false);
					} else
					{
						if (oDrawingDocument.SlideCurrent < oDrawingDocument.GetSlidesCount() - 1)
						{
							this.m_oWordControl.GoToPage(oDrawingDocument.SlideCurrent + 1);
						}
					}
					break;
				}
				case 36: // home
				{
					if (!oEvent.ShiftKey)
					{
						if (oDrawingDocument.SlideCurrent > 0)
						{
							this.m_oWordControl.GoToPage(0);
						}
					} else
					{
						if (oDrawingDocument.SlideCurrent > 0)
						{
							this.CorrectShiftSelect(true, true);
						}
					}
					break;
				}
				case 35: // end
				{
					var slidesCount = oDrawingDocument.GetSlidesCount();
					if (!oEvent.ShiftKey)
					{
						if (oDrawingDocument.SlideCurrent !== (slidesCount - 1))
						{
							this.m_oWordControl.GoToPage(slidesCount - 1);
						}
					} else
					{
						if (oDrawingDocument.SlideCurrent !== (slidesCount - 1))
						{
							this.CorrectShiftSelect(false, true);
						}
					}
					break;
				}
				case 33: // PgUp
				case 38: // UpArrow
				{
					if (oEvent.CtrlKey && oEvent.ShiftKey)
					{
						oPresentation.moveSelectedSlidesToStart();
					} else if (oEvent.CtrlKey)
					{
						oPresentation.moveSlidesPrevPos();
					} else if (oEvent.ShiftKey)
					{
						this.CorrectShiftSelect(true, false);
					} else
					{
						if (oDrawingDocument.SlideCurrent > 0)
						{
							this.m_oWordControl.GoToPage(oDrawingDocument.SlideCurrent - 1);
						}
					}
					break;
				}
				case 77:
				{
					if (oEvent.CtrlKey)
					{
						if (oPresentation.CanEdit())
						{
							sSelectedIdx = this.GetSelectedArray();
							if (sSelectedIdx.length > 0)
							{
								this.m_oWordControl.GoToPage(sSelectedIdx[sSelectedIdx.length - 1]);
								oPresentation.addNextSlide();
								bReturnValue = false;
							} else if (this.GetSlidesCount() === 0)
							{
								oPresentation.addNextSlide();
								this.m_oWordControl.GoToPage(0);
							}
						}
					}
					bReturnValue = false;
					bPreventDefault = true;
					break;
				}
				case 122:
				case 123:
				{
					return;
				}
				default:
					break;
			}
		}
		if (bPreventDefault)
		{
			e.preventDefault();
		}
		if(nStartHistoryIndex === oPresentation.History.Index)
		{
			oApi.sendEvent("asc_onSelectionEnd");
		}
		return bReturnValue;
	};

	CThumbnailsManager.prototype.SetFocusElement = function(type)
	{
		if(this.FocusObjType === type)
			return;
		switch (type)
		{
			case FOCUS_OBJECT_MAIN:
			{
				this.FocusObjType = FOCUS_OBJECT_MAIN;
				break;
			}
			case FOCUS_OBJECT_THUMBNAILS:
			{
				if (this.LockMainObjType)
					return;

				this.FocusObjType = FOCUS_OBJECT_THUMBNAILS;
				if (this.m_oWordControl.m_oLogicDocument)
				{
					this.m_oWordControl.m_oLogicDocument.resetStateCurSlide(true);
				}
				break;
			}
			case FOCUS_OBJECT_ANIM_PANE:
			{
				this.FocusObjType = FOCUS_OBJECT_ANIM_PANE;
				break;
			}
			case FOCUS_OBJECT_NOTES:
			{
				break;
			}
			default:
				break;
		}
	};

	// draw
	CThumbnailsManager.prototype.ShowPage = function (pageNum) {
		const scrollApi = this.m_oWordControl.m_oScrollThumbApi;
		if (!scrollApi) return;

		let startCoord, endCoord, visibleAreaSize, scrollTo, scrollBy;
		if (Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom) {
			startCoord = this.m_arrPages[pageNum].left - this.const_border_w;
			endCoord = this.m_arrPages[pageNum].right + this.const_border_w;
			visibleAreaSize = this.m_oWordControl.m_oThumbnails.HtmlElement.width;

			scrollTo = scrollApi.scrollToX.bind(scrollApi);
			scrollBy = scrollApi.scrollByX.bind(scrollApi);
		} else {
			startCoord = this.m_arrPages[pageNum].top - this.const_border_w;
			endCoord = this.m_arrPages[pageNum].bottom + this.const_border_w;
			visibleAreaSize = this.m_oWordControl.m_oThumbnails.HtmlElement.height;

			scrollTo = scrollApi.scrollToY.bind(scrollApi);
			scrollBy = scrollApi.scrollByY.bind(scrollApi);
		}

		if (startCoord < 0) {
			const size = endCoord - startCoord;
			const shouldReversePageIndexes = Asc.editor.isRtlInterface &&
				Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;
			const pos = shouldReversePageIndexes
				? (size + this.const_border_w) * (this.m_arrPages.length - pageNum - 1)
				: (size + this.const_border_w) * pageNum;
			scrollTo(pos);
		} else if (endCoord > visibleAreaSize) {
			scrollBy(endCoord + this.const_border_w - visibleAreaSize);
		}
	};

	CThumbnailsManager.prototype.ClearCacheAttack = function()
	{
		var pages_count = this.m_arrPages.length;

		for (var i = 0; i < pages_count; i++)
		{
			this.m_arrPages[i].IsRecalc = true;
		}

		this.m_bIsUpdate = true;
	};

	CThumbnailsManager.prototype.FocusRectDraw = function(ctx, x, y, r, b)
	{
		ctx.rect(x - this.const_border_w, y, r - x + this.const_border_w, b - y);
	};
	CThumbnailsManager.prototype.FocusRectFlat = function(_color, ctx, x, y, r, b, isFocus)
	{
		ctx.beginPath();
		ctx.strokeStyle = _color;
		var lineW = Math.round((isFocus ? 2 : 3) * AscCommon.AscBrowser.retinaPixelRatio);
		var dist = Math.round(2 * AscCommon.AscBrowser.retinaPixelRatio);
		ctx.lineWidth = lineW;
		var extend = dist + (lineW / 2);

		ctx.rect(x - extend, y - extend, r - x + (2 * extend), b - y + (2 * extend));
		ctx.stroke();

		ctx.beginPath();
	};
	CThumbnailsManager.prototype.OutlineRect = function(_color, ctx, x, y, r, b)
	{
		ctx.beginPath();
		ctx.strokeStyle = "#C0C0C0";
		const lineW = Math.max(1, Math.round(1 * AscCommon.AscBrowser.retinaPixelRatio));
		ctx.lineWidth = lineW;
		const extend = lineW / 2;
		ctx.rect(x - extend, y - extend, r - x + (2 * extend), b - y + (2 * extend));
		ctx.stroke();

		ctx.beginPath();
	};

	CThumbnailsManager.prototype.DrawAnimLabel = function(oGraphics, nX, nY, oColor)
	{
		let fCX = function(nVal)
		{
			return AscCommon.AscBrowser.convertToRetinaValue(nVal, true) + nX;
		};
		let fCY = function(nVal)
		{
			return AscCommon.AscBrowser.convertToRetinaValue(nVal, true) + nY;
		};
		oGraphics.b_color1(oColor.R, oColor.G, oColor.B, 255);
		oGraphics.SaveGrState();
		oGraphics.SetIntegerGrid(true);
		let oCtx = oGraphics.m_oContext;
		oCtx.beginPath();
		oCtx.moveTo(fCX(10.5), fCY(4));
		oCtx.lineTo(fCX(12), fCY(8));
		oCtx.lineTo(fCX(16), fCY(8));
		oCtx.lineTo(fCX(12.5), fCY(10.5));
		oCtx.lineTo(fCX(14), fCY(15));
		oCtx.lineTo(fCX(10.5), fCY(12.5));
		oCtx.lineTo(fCX(7), fCY(15));
		oCtx.lineTo(fCX(8.5), fCY(10.5));
		oCtx.lineTo(fCX(5), fCY(8));
		oCtx.lineTo(fCX(9), fCY(8));
		oCtx.lineTo(fCX(10.5), fCY(4));
		oCtx.closePath();
		oCtx.fill();

		oCtx.beginPath();
		oCtx.moveTo(fCX(6), fCY(5))
		oCtx.lineTo(fCX(9), fCY(5));
		oCtx.lineTo(fCX(9), fCY(4))
		oCtx.lineTo(fCX(6), fCY(4));
		oCtx.lineTo(fCX(6), fCY(5))
		oCtx.closePath();
		oCtx.fill();

		oCtx.beginPath();
		oCtx.moveTo(fCX(4), fCY(7));
		oCtx.lineTo(fCX(8), fCY(7));
		oCtx.lineTo(fCX(8), fCY(6));
		oCtx.lineTo(fCX(4), fCY(6));
		oCtx.lineTo(fCX(4), fCY(7));
		oCtx.closePath();
		oCtx.fill();
		oCtx.beginPath();
		oGraphics.RestoreGrState();
		return {maxX: fCX(16), minX: fCX(4), maxY: fCY(15), minY: fCY(4)}
	};
	CThumbnailsManager.prototype.DrawPin = function(oGraphics, nX, nY, oColor, nAngle) {
		const fCX = function(nVal)
		{
			return AscCommon.AscBrowser.convertToRetinaValue(nVal, true) + nX;
		};
		const fCY = function(nVal)
		{
			return AscCommon.AscBrowser.convertToRetinaValue(nVal, true) + nY;
		};
		const rotateAt = function(angle, x, y) {
			oCtx.translate(fCX(x), fCY(y));
			oCtx.rotate(AscCommon.deg2rad(angle));
			oCtx.translate(-fCX(x), -fCY(y));
		}
		oGraphics.b_color1(oColor.R, oColor.G, oColor.B, 255);
		oGraphics.SaveGrState();
		oGraphics.SetIntegerGrid(true);
		let oCtx = oGraphics.m_oContext;
		oCtx.save();
		oCtx.lineCap = 'round';
		oCtx.lineWidth = AscCommon.AscBrowser.convertToRetinaValue(1, true);
		rotateAt(nAngle, 5, 9);
		oCtx.beginPath();
		oCtx.moveTo(fCX(2), fCY(4));
		oCtx.lineTo(fCX(8), fCY(4));
		oCtx.closePath();
		oCtx.stroke();

		oCtx.beginPath();
		oCtx.moveTo(fCX(3), fCY(4));
		oCtx.lineTo(fCX(7), fCY(4));
		oCtx.lineTo(fCX(7), fCY(7));
		oCtx.lineTo(fCX(9), fCY(9));
		oCtx.lineTo(fCX(9), fCY(10));
		oCtx.lineTo(fCX(1), fCY(10));
		oCtx.lineTo(fCX(1), fCY(9));
		oCtx.lineTo(fCX(3), fCY(7));
		oCtx.closePath();
		oCtx.fill();

		oCtx.beginPath();
		oCtx.moveTo(fCX(5), fCY(10));
		oCtx.lineTo(fCX(5), fCY(14));
		oCtx.closePath();
		oCtx.stroke();
		rotateAt(-nAngle, 5, 9);
		oCtx.beginPath();
		oCtx.restore();
		oGraphics.RestoreGrState();
	};

	CThumbnailsManager.prototype.OnPaint = function () {
		if (!this.isThumbnailsShown()) {
			return;
		}

		const canvas = this.m_oWordControl.m_oThumbnails.HtmlElement;
		if (!canvas)
			return;

		this.m_oWordControl.m_oApi.clearEyedropperImgData();

		const context = AscCommon.AscBrowser.getContext2D(canvas);
		context.clearRect(0, 0, canvas.width, canvas.height);

		const digitDistance = this.const_offset_x * g_dKoef_pix_to_mm;
		const logicDocument = this.m_oWordControl.m_oLogicDocument;
		const arrSlides = logicDocument.GetAllSlides();
		for (let slideIndex = 0; slideIndex < this.GetSlidesCount(); slideIndex++) {
			const page = this.m_arrPages[slideIndex];

			const oSlide = arrSlides[slideIndex];
			const slideType = oSlide.deleteLock.Lock.Get_Type();
			const bLocked = slideType !== AscCommon.c_oAscLockTypes.kLockTypeMine && slideType !== AscCommon.c_oAscLockTypes.kLockTypeNone;

			if (slideIndex < this.m_lDrawingFirst || slideIndex > this.m_lDrawingEnd) {
				this.m_oCacheManager.UnLock(page.cachedImage);
				page.cachedImage = null;
				continue;
			}

			const font = {
				FontFamily: { Name: "Arial", Index: -1 },
				Italic: false,
				Bold: false,
				FontSize: Math.round(10 * AscCommon.AscBrowser.retinaPixelRatio)
			};

			// Create graphics drawer
			const graphics = new AscCommon.CGraphics();
			graphics.init(context, canvas.width, canvas.height, canvas.width * g_dKoef_pix_to_mm, canvas.height * g_dKoef_pix_to_mm);
			graphics.m_oFontManager = this.m_oFontManager;
			graphics.transform(1, 0, 0, 1, 0, 0);
			graphics.SetFont(font);

			page.Draw(context, page.left, page.top, page.right - page.left, page.bottom - page.top);

			const slideNumber = logicDocument.GetSlideNumber(slideIndex);
			if (slideNumber !== null) {
				let currentSlideNumber = slideNumber;
				let slideNumberTextWidth = 0;
				while (currentSlideNumber !== 0) {
					const lastDigit = currentSlideNumber % 10;
					slideNumberTextWidth += this.DigitWidths[lastDigit];
					currentSlideNumber = (currentSlideNumber / 10) >> 0;
				}

				// Draw slide number
				const textColorHex = bLocked
					? AscCommon.GlobalSkin.ThumbnailsLockColor
					: AscCommon.GlobalSkin.ThumbnailsPageNumberText;

				const textColor = AscCommon.RgbaHexToRGBA(textColorHex);
				graphics.b_color1(textColor.R, textColor.G, textColor.B, 255);

				const canvasWidthMm = canvas.width * g_dKoef_pix_to_mm;
				const canvasHeightMm = canvas.height * g_dKoef_pix_to_mm;

				let textPosX, textPosY;
				if (Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom) {
					const textHeightMm = 10 * AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_pix_to_mm;
					textPosY = (page.top * g_dKoef_pix_to_mm + textHeightMm) / 2;
					textPosX = Asc.editor.isRtlInterface
						? page.right * g_dKoef_pix_to_mm - slideNumberTextWidth - 1 * AscCommon.AscBrowser.retinaPixelRatio
						: page.left * g_dKoef_pix_to_mm + 1 * AscCommon.AscBrowser.retinaPixelRatio;
				} else {
					textPosY = page.top * g_dKoef_pix_to_mm + 3 * AscCommon.AscBrowser.retinaPixelRatio;
					textPosX = Asc.editor.isRtlInterface
						? canvasWidthMm - (digitDistance + slideNumberTextWidth) / 2
						: (digitDistance - slideNumberTextWidth) / 2;
				}
				const textBounds = graphics.t("" + slideNumber, textPosX, textPosY, true);

				// Draw stroke and background if slide is hidden
				const oSlide = logicDocument.GetSlide(slideIndex);
				if (oSlide && !oSlide.isVisible()) {
					context.lineWidth = AscCommon.AscBrowser.convertToRetinaValue(1, true);
					context.strokeStyle = textColorHex;

					context.beginPath();
					context.moveTo(textBounds.x - 3, textBounds.y);
					context.lineTo(textBounds.r + 3, textBounds.b);
					context.stroke();

					context.beginPath();
					context.fillStyle = AscCommon.GlobalSkin.BackgroundColorThumbnails;
					context.globalAlpha = 0.5;
					context.fillRect(page.left, page.top, page.right - page.left, page.bottom - page.top);
					context.globalAlpha = 1;
				}

				// Draw animation label if slide has transition
				page.animateLabelRect = null;
				let nBottomBounds = textBounds.b;
				if (logicDocument.isSlideAnimated(slideIndex)) {
					let iconX, iconY;
					if (Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom) {
						iconX = Asc.editor.isRtlInterface
							? textBounds.x - 3 - AscCommon.AscBrowser.convertToRetinaValue(19, true)
							: textBounds.r + 3;
						iconY = (textBounds.y + textBounds.b) / 2 - AscCommon.AscBrowser.convertToRetinaValue(9.5, true);
					} else {
						iconX = (textBounds.x + textBounds.r) / 2 - AscCommon.AscBrowser.convertToRetinaValue(9.5, true);
						iconY = nBottomBounds + 3;
					}
					const iconHeight = AscCommon.AscBrowser.convertToRetinaValue(15, true);

					if (iconY + iconHeight < page.bottom) {
						const labelCoords = this.DrawAnimLabel(graphics, iconX, iconY, textColor);
						page.animateLabelRect = labelCoords;
						nBottomBounds = labelCoords.maxY;
					}
				}

				// Draw pin if masterSlide is preserved
				if (logicDocument.isSlidePreserved(slideIndex)) {
					const pinOriginalWidth = AscCommon.AscBrowser.convertToRetinaValue(8, true);
					const pinOriginalHeight = AscCommon.AscBrowser.convertToRetinaValue(10, true);
					const pinAngle = 45;
					const pinSizes = AscFormat.fGetOuterRectangle(pinOriginalWidth, pinOriginalHeight, pinAngle);

					let pinX, pinY;
					if (Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom) {
						pinX = Asc.editor.isRtlInterface
							? textBounds.x - 3 - pinSizes.width
							: textBounds.r + 3 + pinSizes.width / 2;
						pinY = textBounds.b - pinSizes.height;
					} else {
						pinX = (textBounds.x + textBounds.r) / 2 - pinSizes.width / 2;
						pinY = nBottomBounds + 3;
					}

					let nIconH = pinSizes.height;
					if (pinY + nIconH < page.bottom) {
						this.DrawPin(graphics, pinX, pinY, textColor, pinAngle);
					}
				}
			}
		}

		this.OnUpdateOverlay();
	};

	CThumbnailsManager.prototype.onCheckUpdate = function()
	{
		if (!this.isThumbnailsShown() || 0 == this.DigitWidths.length)
			return;

		if (this.m_oWordControl.m_oApi.isSaveFonts_Images)
			return;

		if (this.m_lDrawingFirst == -1 || this.m_lDrawingEnd == -1)
		{
			if (this.m_oWordControl.m_oDrawingDocument.IsEmptyPresentation)
			{
				if (!this.bIsEmptyDrawed)
				{
					this.bIsEmptyDrawed = true;
					this.OnPaint();
				}
				return;
			}

			this.bIsEmptyDrawed = false;
			return;
		}

		this.bIsEmptyDrawed = false;

		if (!this.m_bIsUpdate)
		{
			// определяем, нужно ли пересчитать и закэшировать табнейл (хотя бы один)
			for (var i = this.m_lDrawingFirst; i <= this.m_lDrawingEnd; i++)
			{
				var page = this.m_arrPages[i];
				if (null == page.cachedImage || page.IsRecalc)
				{
					this.m_bIsUpdate = true;
					break;
				}
				if ((page.cachedImage.image.width != (page.right - page.left)) || (page.cachedImage.image.height != (page.bottom - page.top)))
				{
					this.m_bIsUpdate = true;
					break;
				}
			}
		}

		if (!this.m_bIsUpdate)
			return;

		for (var i = this.m_lDrawingFirst; i <= this.m_lDrawingEnd; i++)
		{
			var page = this.m_arrPages[i];
			var w = page.right - page.left;
			var h = page.bottom - page.top;

			if (null != page.cachedImage)
			{
				if ((page.cachedImage.image.width != w) || (page.cachedImage.image.height != h) || page.IsRecalc)
				{
					this.m_oCacheManager.UnLock(page.cachedImage);
					page.cachedImage = null;
				}
			}

			if (null == page.cachedImage)
			{
				page.cachedImage = this.m_oCacheManager.Lock(w, h);

				var g = new AscCommon.CGraphics();
				g.IsNoDrawingEmptyPlaceholder = true;
				g.IsThumbnail = true;
				if (this.m_oWordControl.m_oLogicDocument.IsVisioEditor())
				{
					const sizes = this.m_oWordControl.m_oLogicDocument.GetSizesMM(i);
					//todo override CThumbnailsManager
					let SlideWidth = sizes.width;
					let SlideHeight = sizes.height;
					g.init(page.cachedImage.image.ctx, w, h, SlideWidth, SlideHeight);
				}
				else
				{
					g.init(page.cachedImage.image.ctx, w, h, this.SlideWidth, this.SlideHeight);
				}
				g.m_oFontManager = this.m_oFontManager;

				g.transform(1, 0, 0, 1, 0, 0);

				var bIsShowPars = this.m_oWordControl.m_oApi.ShowParaMarks;
				this.m_oWordControl.m_oApi.ShowParaMarks = false;
				this.m_oWordControl.m_oLogicDocument.DrawPage(i, g);
				this.m_oWordControl.m_oApi.ShowParaMarks = bIsShowPars;

				page.IsRecalc = false;

				this.m_bIsUpdate = true;
				break;
			}
		}

		this.OnPaint();
		this.m_bIsUpdate = false;
	};

	CThumbnailsManager.prototype.OnUpdateOverlay = function () {
		if (!this.isThumbnailsShown())
			return;

		const canvas = this.m_oWordControl.m_oThumbnailsBack.HtmlElement;
		if (canvas == null)
			return;

		if (this.m_oWordControl)
			this.m_oWordControl.m_oApi.checkLastWork();

		const context = canvas.getContext("2d");
		const canvasWidth = canvas.width;
		const canvasHeight = canvas.height;

		this.drawThumbnailsBorders(context, canvasWidth, canvasHeight);

		if (this.MouseDownTrack.IsDragged()) {
			this.drawThumbnailsInsertionLine(context, canvasWidth, canvasHeight);
		}
	};
	CThumbnailsManager.prototype.drawThumbnailsBorders = function (context, canvasWidth, canvasHeight) {
		context.fillStyle = GlobalSkin.BackgroundColorThumbnails;
		context.fillRect(0, 0, canvasWidth, canvasHeight);

		const _style_select = GlobalSkin.ThumbnailsPageOutlineActive;
		const _style_focus = GlobalSkin.ThumbnailsPageOutlineHover;
		const _style_select_focus = GlobalSkin.ThumbnailsPageOutlineActive;

		context.fillStyle = _style_select;

		const oLogicDocument = this.m_oWordControl && this.m_oWordControl.m_oLogicDocument;
		if (!oLogicDocument) {
			return;
		}

		const arrSlides = oLogicDocument.GetAllSlides();


		for (let i = 0; i < arrSlides.length; i++) {
			const page = this.m_arrPages[i];
			const oSlide = arrSlides[i];
			const slideType = oSlide.deleteLock.Lock.Get_Type();
			const bLocked = slideType !== AscCommon.c_oAscLockTypes.kLockTypeMine && slideType !== AscCommon.c_oAscLockTypes.kLockTypeNone;

			let color = null;
			let isFocus = false;

			if (bLocked) {
				color = AscCommon.GlobalSkin.ThumbnailsLockColor;
			} else if (page.IsSelected && page.IsFocused) {
				color = _style_select_focus;
			} else if (page.IsSelected) {
				color = _style_select;
			} else if (page.IsFocused) {
				color = _style_focus;
				isFocus = true;
			}

			if (color) {
				this.FocusRectFlat(color, context, page.left, page.top, page.right, page.bottom, isFocus);
			} else {
				this.OutlineRect(color, context, page.left, page.top, page.right, page.bottom);
			}
		}
	};
	CThumbnailsManager.prototype.drawThumbnailsInsertionLine = function (context, canvasWidth, canvasHeight) {
		const isHorizontalThumbnails = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;
		const isRightToLeft = Asc.editor.isRtlInterface;
		const nPosition = this.MouseDownTrack.GetPosition();
		const oPage = this.m_arrPages[nPosition - 1];

		context.strokeStyle = "#DEDEDE";
		if (isHorizontalThumbnails) {
			const containerRightBorder = editor.WordControl.m_oThumbnailsContainer.AbsolutePosition.R * AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_mm_to_pix;

			// x, topY & bottomY - coordinates of the insertion line itself
			const x = oPage
				? isRightToLeft
					? (oPage.left - 1.5 * this.const_border_w) >> 0
					: (oPage.right + 1.5 * this.const_border_w) >> 0
				: isRightToLeft
					? (containerRightBorder - this.const_offset_x / 2) >> 0
					: (this.const_offset_x / 2) >> 0;

			let topY, bottomY;
			if (this.m_arrPages.length > 0) {
				topY = this.m_arrPages[0].top + 4;
				bottomY = this.m_arrPages[0].bottom - 4;
			} else {
				topY = 0;
				bottomY = canvasHeight;
			}

			context.beginPath();
			context.moveTo(x + 0.5, topY);
			context.lineTo(x + 0.5, bottomY);
			context.closePath();
			context.lineWidth = 3;
			context.stroke();

			if (this.MouseTrackCommonImage) {
				const commonImageWidth = this.MouseTrackCommonImage.width;
				const commonImageHeight = this.MouseTrackCommonImage.height;

				context.save();
				context.translate(x, topY);
				context.rotate(Math.PI / 2);

				context.drawImage(
					this.MouseTrackCommonImage,
					0, 0,
					commonImageWidth / 2, commonImageHeight,
					-commonImageWidth, -commonImageHeight / 2,
					commonImageWidth / 2, commonImageHeight
				);

				context.translate(bottomY - topY, 0);
				context.drawImage(
					this.MouseTrackCommonImage,
					commonImageWidth / 2, 0,
					commonImageWidth / 2, commonImageHeight,
					commonImageWidth / 2, -commonImageHeight / 2,
					commonImageWidth / 2, commonImageHeight
				);

				context.restore();
			}
		} else {
			const y = oPage
				? (oPage.bottom + 1.5 * this.const_border_w) >> 0
				: this.const_offset_y / 2 >> 0;

			let _left_pos = 0;
			let _right_pos = canvasWidth;
			if (this.m_arrPages.length > 0) {
				_left_pos = this.m_arrPages[0].left + 4;
				_right_pos = this.m_arrPages[0].right - 4;
			}

			context.lineWidth = 3;
			context.beginPath();
			context.moveTo(_left_pos, y + 0.5);
			context.lineTo(_right_pos, y + 0.5);
			context.stroke();
			context.beginPath();

			if (null != this.MouseTrackCommonImage) {
				context.drawImage(
					this.MouseTrackCommonImage,
					0, 0,
					(this.MouseTrackCommonImage.width + 1) >> 1, this.MouseTrackCommonImage.height,
					_left_pos - this.MouseTrackCommonImage.width, y - (this.MouseTrackCommonImage.height >> 1),
					(this.MouseTrackCommonImage.width + 1) >> 1, this.MouseTrackCommonImage.height
				);

				context.drawImage(
					this.MouseTrackCommonImage,
					this.MouseTrackCommonImage.width >> 1, 0,
					(this.MouseTrackCommonImage.width + 1) >> 1, this.MouseTrackCommonImage.height,
					_right_pos + (this.MouseTrackCommonImage.width >> 1), y - (this.MouseTrackCommonImage.height >> 1),
					(this.MouseTrackCommonImage.width + 1) >> 1, this.MouseTrackCommonImage.height
				);
			}
		}
	};

	// select
	CThumbnailsManager.prototype.SelectSlides = function(aSelectedIdx, bReset)
	{
		if (aSelectedIdx.length > 0)
		{
			if(bReset === undefined || bReset)
			{
				for (let i = 0; i < this.m_arrPages.length; i++)
				{
					this.m_arrPages[i].IsSelected = false;
				}
			}
			let nCount = aSelectedIdx.length;
			let nSlideType = this.GetSlideType(aSelectedIdx[0]);
			for (let nIdx = 0; nIdx < nCount; nIdx++)
			{
				if(this.GetSlideType(aSelectedIdx[nIdx]) === nSlideType)
				{
					let oPage = this.m_arrPages[aSelectedIdx[nIdx]];
					if (oPage)
					{
						oPage.IsSelected = true;
					}
				}
			}
			this.OnUpdateOverlay();
		}
	};

	CThumbnailsManager.prototype.GetSelectedSlidesRange = function()
	{
		var _min = this.m_oWordControl.m_oDrawingDocument.SlideCurrent;
		var _max = _min;
		var pages_count = this.m_arrPages.length;
		for (var i = 0; i < pages_count; i++)
		{
			if (this.m_arrPages[i].IsSelected)
			{
				if (i < _min)
					_min = i;
				if (i > _max)
					_max = i;
			}
		}
		return {Min: _min, Max: _max};
	};

	CThumbnailsManager.prototype.GetSelectedArray = function()
	{
		var _array = [];
		var pages_count = this.m_arrPages.length;
		for (var i = 0; i < pages_count; i++)
		{
			if (this.m_arrPages[i].IsSelected)
			{
				_array[_array.length] = i;
			}
		}
		return _array;
	};

	CThumbnailsManager.prototype.CorrectShiftSelect = function(isTop, isEnd)
	{
		var drDoc = this.m_oWordControl.m_oDrawingDocument;
		var slidesCount = drDoc.GetSlidesCount();
		var min_max = this.GetSelectedSlidesRange();

		var _page = this.m_oWordControl.m_oDrawingDocument.SlideCurrent;
		if (isEnd)
		{
			_page = isTop ? 0 : slidesCount - 1;
		} else if (isTop)
		{
			if (min_max.Min != _page)
			{
				_page = min_max.Min - 1;
				if (_page < 0)
					_page = 0;
			} else
			{
				_page = min_max.Max - 1;
				if (_page < 0)
					_page = 0;
			}
		} else
		{
			if (min_max.Min != _page)
			{
				_page = min_max.Min + 1;
				if (_page >= slidesCount)
					_page = slidesCount - 1;
			} else
			{
				_page = min_max.Max + 1;
				if (_page >= slidesCount)
					_page = slidesCount - 1;
			}
		}

		var _max = _page;
		var _min = this.m_oWordControl.m_oDrawingDocument.SlideCurrent;
		if (_min > _max)
		{
			var _temp = _max;
			_max = _min;
			_min = _temp;
		}

		for (var i = 0; i < _min; i++)
		{
			this.m_arrPages[i].IsSelected = false;
		}
		let nSlideType = this.GetSlideType(_min);
		for (var i = _min; i <= _max; i++)
		{
			if(this.GetSlideType(i) === nSlideType)
			{
				this.m_arrPages[i].IsSelected = true;
			}
		}
		for (var i = _max + 1; i < slidesCount; i++)
		{
			this.m_arrPages[i].IsSelected = false;
		}


		this.m_oWordControl.m_oLogicDocument.Document_UpdateInterfaceState();
		this.OnUpdateOverlay();
		this.ShowPage(_page);
	};

	CThumbnailsManager.prototype.SelectAll = function()
	{
		for (var i = 0; i < this.m_arrPages.length; i++)
		{
			this.m_arrPages[i].IsSelected = true;
		}
		this.UpdateInterface();
		this.OnUpdateOverlay();
	};

	CThumbnailsManager.prototype.IsSlideHidden = function(aSelected)
	{
		var oPresentation = this.m_oWordControl.m_oLogicDocument;
		for (var i = 0; i < aSelected.length; ++i)
		{
			if (oPresentation.IsVisibleSlide(aSelected[i]))
				return false;
		}
		return true;
	};

	CThumbnailsManager.prototype.isSelectedPage = function(pageNum)
	{
		if (this.m_arrPages[pageNum] && this.m_arrPages[pageNum].IsSelected)
			return true;
		return false;
	};

	CThumbnailsManager.prototype.SelectPage = function (pageNum) {
		if (!this.SelectPageEnabled)
			return;

		const pagesCount = this.m_arrPages.length;
		if (pageNum < 0 || pageNum >= pagesCount)
			return;

		let bIsUpdate = false;

		for (let i = 0; i < pagesCount; i++) {
			if (this.m_arrPages[i].IsSelected === true && i != pageNum) {
				bIsUpdate = true;
			}
			this.m_arrPages[i].IsSelected = false;
		}

		if (this.m_arrPages[pageNum].IsSelected === false) {
			bIsUpdate = true;
		}
		this.m_arrPages[pageNum].IsSelected = true;
		this.m_bIsUpdate = bIsUpdate;

		if (bIsUpdate) {
			this.ShowPage(pageNum);
		}
	};

	// position
	CThumbnailsManager.prototype.ConvertCoords = function (x, y) {
		let posX, posY;
		switch (Asc.editor.getThumbnailsPosition()) {
			case thumbnailsPositionMap.left:
				posX = x - this.m_oWordControl.X;
				posY = y - this.m_oWordControl.Y;
				break;

			case thumbnailsPositionMap.right:
				posX = x - this.m_oWordControl.X - this.m_oWordControl.Width + this.m_oWordControl.splitters[0].position * g_dKoef_mm_to_pix;
				posY = y - this.m_oWordControl.Y;
				break;

			case thumbnailsPositionMap.bottom:
				posX = x - this.m_oWordControl.X;
				posY = y - this.m_oWordControl.Y - this.m_oWordControl.Height + this.m_oWordControl.splitters[0].position * g_dKoef_mm_to_pix;
				break;

			default:
				posX = posY = 0; // just in case
				break;
		}

		const convertedX = AscCommon.AscBrowser.convertToRetinaValue(posX, true);
		const convertedY = AscCommon.AscBrowser.convertToRetinaValue(posY, true);
		const pageIndex = this.m_arrPages.findIndex(function (page) {
			return page.Hit(convertedX, convertedY);
		});

		return {
			X: convertedX,
			Y: convertedY,
			Page: pageIndex
		};
	};
	CThumbnailsManager.prototype.ConvertCoords2 = function (x, y) {
		if (this.m_arrPages.length == 0)
			return -1;

		let posX, posY;
		switch (Asc.editor.getThumbnailsPosition()) {
			case thumbnailsPositionMap.left:
				posX = x - this.m_oWordControl.X;
				posY = y - this.m_oWordControl.Y;
				break;

			case thumbnailsPositionMap.right:
				posX = x - this.m_oWordControl.X - this.m_oWordControl.Width + this.m_oWordControl.splitters[0].position * g_dKoef_mm_to_pix;
				posY = y - this.m_oWordControl.Y;
				break;

			case thumbnailsPositionMap.bottom:
				posX = x - this.m_oWordControl.X;
				posY = y - this.m_oWordControl.Y - this.m_oWordControl.Height + this.m_oWordControl.splitters[0].position * g_dKoef_mm_to_pix;
				break;

			default:
				posX = posY = 0; // just in case
				break;
		}

		const convertedX = AscCommon.AscBrowser.convertToRetinaValue(posX, true);
		const convertedY = AscCommon.AscBrowser.convertToRetinaValue(posY, true);

		const absolutePosition = this.m_oWordControl.m_oThumbnails.AbsolutePosition;
		const thControlWidth = (absolutePosition.R - absolutePosition.L) * g_dKoef_mm_to_pix * AscCommon.AscBrowser.retinaPixelRatio;
		const thControlHeight = (absolutePosition.B - absolutePosition.T) * g_dKoef_mm_to_pix * AscCommon.AscBrowser.retinaPixelRatio;

		if (convertedX < 0 || convertedX > thControlWidth || convertedY < 0 || convertedY > thControlHeight)
			return -1;

		const isHorizontalThumbnails = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;
		const isRightToLeft = Asc.editor.isRtlInterface;

		let minDistance = Infinity;
		let minPositionPage = 0;

		for (let i = 0; i < this.m_arrPages.length; i++) {
			const page = this.m_arrPages[i];

			let distanceToStart, distanceToEnd;
			if (isHorizontalThumbnails) {
				if (isRightToLeft) {
					distanceToStart = Math.abs(convertedX - page.right);
					distanceToEnd = Math.abs(convertedX - page.left);
				} else {
					distanceToStart = Math.abs(convertedX - page.left);
					distanceToEnd = Math.abs(convertedX - page.right);
				}
			} else {
				distanceToStart = Math.abs(convertedY - page.top);
				distanceToEnd = Math.abs(convertedY - page.bottom);
			}

			if (distanceToStart < minDistance) {
				minDistance = distanceToStart;
				minPositionPage = i;
			}
			if (distanceToEnd < minDistance) {
				minDistance = distanceToEnd;
				minPositionPage = i + 1;
			}
		}

		return minPositionPage;
	};
	CThumbnailsManager.prototype.CalculatePlaces = function () {
		if (!this.isThumbnailsShown())
			return;

		if (this.m_oWordControl && this.m_oWordControl.MobileTouchManagerThumbnails)
			this.m_oWordControl.MobileTouchManagerThumbnails.ClearContextMenu();

		const thumbnails = this.m_oWordControl.m_oThumbnails;
		const thumbnailsCanvas = thumbnails.HtmlElement;
		if (!thumbnailsCanvas)
			return;

		const canvasWidth = thumbnailsCanvas.width;
		const canvasHeight = thumbnailsCanvas.height;

		const thumbnailsWidthMm = thumbnails.AbsolutePosition.R - thumbnails.AbsolutePosition.L;
		const thumbnailsHeightMm = thumbnails.AbsolutePosition.B - thumbnails.AbsolutePosition.T;

		const pixelRatio = AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_mm_to_pix;
		const isVerticalThumbnails = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.right
			|| Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.left;
		const isRightToLeft = Asc.editor.isRtlInterface;

		let thSlideWidthPx, thSlideHeightPx;
		let startOffset, supplement;
		if (isVerticalThumbnails) {
			thSlideWidthPx = (thumbnailsWidthMm * pixelRatio >> 0) - this.const_offset_x - this.const_offset_r;
			thSlideHeightPx = (thSlideWidthPx * this.SlideHeight / this.SlideWidth) >> 0;
			startOffset = this.const_offset_y;
		} else {
			thSlideHeightPx = (thumbnailsHeightMm * pixelRatio >> 0) - this.const_offset_y - this.const_offset_b;
			thSlideWidthPx = (thSlideHeightPx * this.SlideWidth / this.SlideHeight) >> 0;
			startOffset = this.const_offset_x;
		}

		if (this.m_bIsScrollVisible) {
			const scrollApi = editor.WordControl.m_oScrollThumbApi;
			if (scrollApi) {
				this.m_dScrollY_max = isVerticalThumbnails ? scrollApi.getMaxScrolledY() : scrollApi.getMaxScrolledX();
				this.m_dScrollY = isVerticalThumbnails ? scrollApi.getCurScrolledY() : scrollApi.getCurScrolledX();
			}
		}

		const currentScrollPx = isRightToLeft && !isVerticalThumbnails
			? this.m_dScrollY_max - this.m_dScrollY >> 0
			: this.m_dScrollY >> 0;

		let isFirstSlideFound = false;
		let isLastSlideFound = false;

		const totalSlidesCount = this.GetSlidesCount();
		for (let slideIndex = 0; slideIndex < totalSlidesCount; slideIndex++) {
			if (this.m_oWordControl.m_oLogicDocument.IsVisioEditor()) {
				const sizes = this.m_oWordControl.m_oLogicDocument.GetSizesMM(slideIndex);
				let visioSlideWidthMm = sizes.width;
				let visioSlideHeightMm = sizes.height;
				if (isVerticalThumbnails) {
					thSlideHeightPx = (thSlideWidthPx * visioSlideHeightMm / visioSlideWidthMm) >> 0;
				} else {
					thSlideWidthPx = (thSlideHeightPx * visioSlideWidthMm / visioSlideHeightMm) >> 0;
				}
			}

			if (slideIndex >= this.m_arrPages.length) {
				this.m_arrPages[slideIndex] = new CThPage();
				if (slideIndex === 0)
					this.m_arrPages[0].IsSelected = true;
			}

			const slideData = this.m_oWordControl.m_oLogicDocument.GetSlide(slideIndex);
			const slideRect = this.m_arrPages[slideIndex];
			slideRect.pageIndex = slideIndex;

			if (isVerticalThumbnails) {
				slideRect.left = isRightToLeft ? this.const_offset_r : this.const_offset_x;
				slideRect.top = startOffset - currentScrollPx;
				slideRect.right = slideRect.left + thSlideWidthPx;
				slideRect.bottom = slideRect.top + thSlideHeightPx;

				if (!isFirstSlideFound) {
					if ((startOffset + thSlideHeightPx) > currentScrollPx) {
						this.m_lDrawingFirst = slideIndex;
						isFirstSlideFound = true;
					}
				}

				if (!isLastSlideFound) {
					if (slideRect.top > canvasHeight) {
						this.m_lDrawingEnd = slideIndex - 1;
						isLastSlideFound = true;
					}
				}

				supplement = (thSlideHeightPx + 3 * this.const_border_w);
			} else {
				slideRect.top = this.const_offset_y;
				slideRect.bottom = slideRect.top + thSlideHeightPx;

				if (isRightToLeft) {
					slideRect.right = canvasWidth - startOffset + currentScrollPx;
					slideRect.left = slideRect.right - thSlideWidthPx;
				} else {
					slideRect.left = startOffset - currentScrollPx;
					slideRect.right = slideRect.left + thSlideWidthPx;
				}

				if (!isFirstSlideFound) {
					if ((startOffset + thSlideWidthPx) > currentScrollPx) {
						this.m_lDrawingFirst = slideIndex;
						isFirstSlideFound = true;
					}
				}

				if (!isLastSlideFound) {
					const isHidden = isRightToLeft
						? slideRect.right < 0
						: slideRect.left > canvasWidth;
					if (isHidden) {
						this.m_lDrawingEnd = slideIndex - 1;
						isLastSlideFound = true;
					}
				}

				supplement = (thSlideWidthPx + 3 * this.const_border_w);
			}

			if (slideData.getObjectType() === AscDFH.historyitem_type_SlideLayout) {
				const scaledWidth = (slideRect.right - slideRect.left) * AscCommonSlide.SlideLayout.LAYOUT_THUMBNAIL_SCALE;
				const scaledHeight = (slideRect.bottom - slideRect.top) * AscCommonSlide.SlideLayout.LAYOUT_THUMBNAIL_SCALE;

				if (isVerticalThumbnails) {
					slideRect.bottom = (slideRect.top + scaledHeight) + 0.5 >> 0;
					isRightToLeft
						? slideRect.right = (slideRect.left + scaledWidth) + 0.5 >> 0
						: slideRect.left = (slideRect.right - scaledWidth) + 0.5 >> 0;
					supplement = scaledHeight + 3 * this.const_border_w;
				} else {
					slideRect.top = (slideRect.bottom - scaledHeight) + 0.5 >> 0;
					isRightToLeft
						? slideRect.left = (slideRect.right - scaledWidth) + 0.5 >> 0
						: slideRect.right = (slideRect.left + scaledWidth) + 0.5 >> 0;
					supplement = scaledWidth + 3 * this.const_border_w;
				}
			}

			startOffset += supplement >> 0;
		}

		if (this.m_arrPages.length > totalSlidesCount)
			this.m_arrPages.splice(totalSlidesCount, this.m_arrPages.length - totalSlidesCount);

		if (!isLastSlideFound) {
			this.m_lDrawingEnd = totalSlidesCount - 1;
		}
	};

	CThumbnailsManager.prototype.GetThumbnailPagePosition = function (pageIndex) {
		if (pageIndex < 0 || pageIndex >= this.m_arrPages.length)
			return null;

		const drawRect = this.m_arrPages[pageIndex];
		return {
			X: AscCommon.AscBrowser.convertToRetinaValue(drawRect.left),
			Y: AscCommon.AscBrowser.convertToRetinaValue(drawRect.top),
			W: AscCommon.AscBrowser.convertToRetinaValue(drawRect.right - drawRect.left),
			H: AscCommon.AscBrowser.convertToRetinaValue(drawRect.bottom - drawRect.top)
		};
	};
	CThumbnailsManager.prototype.getSpecialPasteButtonCoords = function (sSlideId) {
		if (!sSlideId) return null;

		const oPresentation = this.m_oWordControl.m_oLogicDocument;
		const aSlides = oPresentation.GetAllSlides();

		const nSlideIdx = aSlides.findIndex(function (slide) {
			return slide.Get_Id() === sSlideId;
		});
		if (nSlideIdx === -1) return null;

		const oRect = this.GetThumbnailPagePosition(nSlideIdx);
		if (!oRect) return null;

		const oThContainer = editor.WordControl.m_oThumbnailsContainer;
		const isHidden = oThContainer.HtmlElement.style.display === "none";
		const isHorizontalThumbnails = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;

		const offsetX = oThContainer.AbsolutePosition.L * g_dKoef_mm_to_pix;
		const offsetY = oThContainer.AbsolutePosition.T * g_dKoef_mm_to_pix;

		const coordX = isHidden
			? 0
			: oRect.X + oRect.W - AscCommon.specialPasteElemWidth;
		const coordY = isHorizontalThumbnails
			? oRect.Y - AscCommon.specialPasteElemHeight
			: oRect.Y + oRect.H;

		return {
			X: offsetX + coordX,
			Y: offsetY + coordY
		};
	};

	// scroll
	CThumbnailsManager.prototype.CheckNeedAnimateScrolls = function (type) {
		switch (type) {
			case -1:
				if (this.MouseThumbnailsAnimateScrollTopTimer != -1) {
					clearInterval(this.MouseThumbnailsAnimateScrollTopTimer);
					this.MouseThumbnailsAnimateScrollTopTimer = -1;
				}
				if (this.MouseThumbnailsAnimateScrollBottomTimer != -1) {
					clearInterval(this.MouseThumbnailsAnimateScrollBottomTimer);
					this.MouseThumbnailsAnimateScrollBottomTimer = -1;
				}
				break;

			case 0:
				if (this.MouseThumbnailsAnimateScrollBottomTimer != -1) {
					clearInterval(this.MouseThumbnailsAnimateScrollBottomTimer);
					this.MouseThumbnailsAnimateScrollBottomTimer = -1;
				}
				if (this.MouseThumbnailsAnimateScrollTopTimer == -1) {
					this.MouseThumbnailsAnimateScrollTopTimer = setInterval(this.OnScrollTrackTop, 50);
				}
				break;

			case 1:
				if (this.MouseThumbnailsAnimateScrollTopTimer != -1) {
					clearInterval(this.MouseThumbnailsAnimateScrollTopTimer);
					this.MouseThumbnailsAnimateScrollTopTimer = -1;
				}
				if (-1 == this.MouseThumbnailsAnimateScrollBottomTimer) {
					this.MouseThumbnailsAnimateScrollBottomTimer = setInterval(this.OnScrollTrackBottom, 50);
				}
				break;
		}
	};
	CThumbnailsManager.prototype.OnScrollTrackTop = function () {
		Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom
			? this.m_oWordControl.m_oScrollThumbApi.scrollByX(-45)
			: this.m_oWordControl.m_oScrollThumbApi.scrollByY(-45);
	};
	CThumbnailsManager.prototype.OnScrollTrackBottom = function () {
		Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom
			? this.m_oWordControl.m_oScrollThumbApi.scrollByX(45)
			: this.m_oWordControl.m_oScrollThumbApi.scrollByY(45);
	};
	CThumbnailsManager.prototype.thumbnailsScroll = function (sender, scrollPosition, maxScrollPosition, isAtTop, isAtBottom) {
		if (false === this.m_oWordControl.m_oApi.bInit_word_control || false === this.m_bIsScrollVisible)
			return;

		this.m_dScrollY = Math.max(0, Math.min(scrollPosition, maxScrollPosition));
		this.m_dScrollY_max = maxScrollPosition;

		this.CalculatePlaces();
		AscCommon.g_specialPasteHelper.SpecialPasteButton_Update_Position();
		this.m_bIsUpdate = true;

		if (!this.m_oWordControl.m_oApi.isMobileVersion)
			this.SetFocusElement(FOCUS_OBJECT_THUMBNAILS);
	};

	// calculate
	CThumbnailsManager.prototype.RecalculateAll = function()
	{
		const sizes = this.m_oWordControl.m_oLogicDocument.GetSizesMM();
		this.SlideWidth = sizes.width;
		this.SlideHeight = sizes.height;
		this.CheckSizes();

		this.ClearCacheAttack();
	};

	// size
	CThumbnailsManager.prototype.CheckSizes = function () {
		this.InitCheckOnResize();

		if (!this.isThumbnailsShown())
			return;

		this.calculateThumbnailsOffsets();

		const wordControl = this.m_oWordControl;
		const dKoefToPix = AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_mm_to_pix;
		const isHorizontalOrientation = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;
		const dPosition = this.m_dScrollY_max != 0 ? this.m_dScrollY / this.m_dScrollY_max : 0;

		// Width and Height of 'id_panel_thumbnails' in millimeters
		const thumbnailsContainerWidth = wordControl.m_oThumbnailsContainer.AbsolutePosition.R - wordControl.m_oThumbnailsContainer.AbsolutePosition.L;
		const thumbnailsContainerHeight = wordControl.m_oThumbnailsContainer.AbsolutePosition.B - wordControl.m_oThumbnailsContainer.AbsolutePosition.T;

		let totalThumbnailsLength = this.calculateTotalThumbnailsLength(thumbnailsContainerWidth, thumbnailsContainerHeight);
		const thumbnailsContainerLength = isHorizontalOrientation
			? thumbnailsContainerWidth * dKoefToPix >> 0
			: thumbnailsContainerHeight * dKoefToPix >> 0;

		if (totalThumbnailsLength < thumbnailsContainerLength) {
			// All thumbnails fit inside - no scroller needed
			if (this.m_bIsScrollVisible && GlobalSkin.ThumbnailScrollWidthNullIfNoScrolling) {
				wordControl.m_oThumbnails.Bounds.R = 0;
				wordControl.m_oThumbnailsBack.Bounds.R = 0;
				wordControl.m_oThumbnails_scroll.Bounds.AbsW = 0;
				wordControl.m_oThumbnailsContainer.Resize(thumbnailsContainerWidth, thumbnailsContainerHeight, wordControl);
			} else {
				wordControl.m_oThumbnails_scroll.HtmlElement.style.display = 'none';
			}
			this.m_bIsScrollVisible = false;
			this.m_dScrollY = isHorizontalOrientation && Asc.editor.isRtlInterface
				? this.m_dScrollY_max
				: 0;

		} else {
			// Scrollbar is needed
			if (!this.m_bIsScrollVisible) {
				if (GlobalSkin.ThumbnailScrollWidthNullIfNoScrolling) {
					wordControl.m_oThumbnailsBack.Bounds.R = wordControl.ScrollWidthPx * dKoefToPix;
					wordControl.m_oThumbnails.Bounds.R = wordControl.ScrollWidthPx * dKoefToPix;
					const _width_mm_scroll = 10;
					wordControl.m_oThumbnails_scroll.Bounds.AbsW = _width_mm_scroll * dKoefToPix;
				} else if (!wordControl.m_oApi.isMobileVersion) {
					wordControl.m_oThumbnails_scroll.HtmlElement.style.display = 'block';
				}

				const thumbnailsContainerParent = wordControl.m_oThumbnailsContainer.Parent; // editor.WordControl.m_oBody
				wordControl.m_oThumbnailsContainer.Resize(
					thumbnailsContainerParent.AbsolutePosition.R - thumbnailsContainerParent.AbsolutePosition.L,
					thumbnailsContainerParent.AbsolutePosition.B - thumbnailsContainerParent.AbsolutePosition.T,
					wordControl
				);
			}
			this.m_bIsScrollVisible = true;

			const thumbnailsWidth = wordControl.m_oThumbnails.AbsolutePosition.R - wordControl.m_oThumbnails.AbsolutePosition.L;
			const thumbnailsHeight = wordControl.m_oThumbnails.AbsolutePosition.B - wordControl.m_oThumbnails.AbsolutePosition.T;
			totalThumbnailsLength = this.calculateTotalThumbnailsLength(thumbnailsWidth, thumbnailsHeight);

			const settings = this.getSettingsForScrollObject(totalThumbnailsLength);
			if (wordControl.m_oScrollThumb_ && wordControl.m_oScrollThumb_.isHorizontalScroll === isHorizontalOrientation) {
				wordControl.m_oScrollThumb_.Repos(settings);
			} else {
				const holder = wordControl.m_oThumbnails_scroll.HtmlElement;
				const canvas = holder.getElementsByTagName('canvas')[0];
				canvas && holder.removeChild(canvas);

				wordControl.m_oScrollThumb_ = new AscCommon.ScrollObject('id_vertical_scroll_thmbnl', settings);
				wordControl.m_oScrollThumbApi = wordControl.m_oScrollThumb_;

				const eventName = isHorizontalOrientation ? 'scrollhorizontal' : 'scrollvertical';
				wordControl.m_oScrollThumb_.bind(eventName, function (evt) {
					const maxScrollPosition = isHorizontalOrientation ? evt.maxScrollX : evt.maxScrollY;
					this.thumbnailsScroll(this, evt.scrollD, maxScrollPosition);
				});
			}
			wordControl.m_oScrollThumb_.isHorizontalScroll = isHorizontalOrientation;

			if (Asc.editor.isRtlInterface && isHorizontalOrientation && this.m_dScrollY_max === 0) {
				wordControl.m_oScrollThumbApi.scrollToX(wordControl.m_oScrollThumbApi.maxScrollX);
			}
		}

		if (this.m_bIsScrollVisible) {
			const lPosition = (dPosition * wordControl.m_oScrollThumbApi.getMaxScrolledY()) >> 0;
			wordControl.m_oScrollThumbApi.scrollToY(lPosition);
		}

		isHorizontalOrientation
			? this.ScrollerWidth = totalThumbnailsLength
			: this.ScrollerHeight = totalThumbnailsLength;
		if (wordControl.MobileTouchManagerThumbnails) {
			wordControl.MobileTouchManagerThumbnails.Resize();
		}

		this.CalculatePlaces();
		AscCommon.g_specialPasteHelper.SpecialPasteButton_Update_Position();
		this.m_bIsUpdate = true;
	};
	CThumbnailsManager.prototype.calculateThumbnailsOffsets = function () {
		const isHorizontalOrientation = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;
		if (isHorizontalOrientation) {
			this.const_offset_y = 25 + Math.round(9 * AscCommon.AscBrowser.retinaPixelRatio)
		} else {
			const _tmpDig = this.DigitWidths.length > 5 ? this.DigitWidths[5] : 0;
			const dKoefToPix = AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_mm_to_pix;
			const slidesCount = this.GetSlidesCount();
			const firstSlideNumber = this.m_oWordControl.m_oLogicDocument.getFirstSlideNumber();
			const totalSlidesLength = String(slidesCount + firstSlideNumber).length;
			this.const_offset_x = Math.max((_tmpDig * dKoefToPix * totalSlidesLength >> 0), 25) + Math.round(9 * AscCommon.AscBrowser.retinaPixelRatio);
		}
	};
	CThumbnailsManager.prototype.calculateTotalThumbnailsLength = function (containerWidth, containerHeight) {
		const dKoefToPix = AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_mm_to_pix;
		const oPresentation = this.m_oWordControl.m_oLogicDocument;
		const slidesCount = this.GetSlidesCount();

		const isHorizontalOrientation = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;

		let thumbnailHeight, thumbnailWidth;
		if (isHorizontalOrientation) {
			thumbnailHeight = (containerHeight * dKoefToPix >> 0) - (this.const_offset_y + this.const_offset_b);
			thumbnailWidth = thumbnailHeight * this.SlideWidth / this.SlideHeight >> 0;
		} else {
			thumbnailWidth = (containerWidth * dKoefToPix >> 0) - (this.const_offset_x + this.const_offset_r);
			thumbnailHeight = thumbnailWidth * this.SlideHeight / this.SlideWidth >> 0;
		}

		const cumulativeThumbnailLength = oPresentation.getCumulativeThumbnailsLength(isHorizontalOrientation, thumbnailWidth, thumbnailHeight);

		const totalThumbnailsLength = isHorizontalOrientation
			? 2 * this.const_offset_x + cumulativeThumbnailLength + (slidesCount > 0 ? (slidesCount - 1) * 3 * this.const_border_w : 0)
			: 2 * this.const_offset_y + cumulativeThumbnailLength + (slidesCount > 0 ? (slidesCount - 1) * 3 * this.const_border_w : 0);

		return totalThumbnailsLength;
	};
	CThumbnailsManager.prototype.getSettingsForScrollObject = function (totalContentLength) {
		const settings = new AscCommon.ScrollSettings();
		settings.screenW = this.m_oWordControl.m_oThumbnails.HtmlElement.width;
		settings.screenH = this.m_oWordControl.m_oThumbnails.HtmlElement.height;
		settings.showArrows = false;
		settings.slimScroll = true;
		settings.cornerRadius = 1;

		settings.scrollBackgroundColor = GlobalSkin.BackgroundColorThumbnails;
		settings.scrollBackgroundColorHover = GlobalSkin.BackgroundColorThumbnails;
		settings.scrollBackgroundColorActive = GlobalSkin.BackgroundColorThumbnails;

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

		if (Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom) {
			settings.isHorizontalScroll = true;
			settings.isVerticalScroll = false;
			settings.contentW = totalContentLength;
		} else {
			settings.isHorizontalScroll = false;
			settings.isVerticalScroll = true;
			settings.contentH = totalContentLength;
		}

		return settings;
	};

	// interface
	CThumbnailsManager.prototype.UpdateInterface = function()
	{
		this.m_oWordControl.m_oLogicDocument.Document_UpdateInterfaceState();
	};

	CThumbnailsManager.prototype.showContextMenu = function(bPosBySelect)
	{
		let sSelectedIdx = this.GetSelectedArray();
		let oMenuPos;
		let bIsSlideSelect = false;
		if(bPosBySelect)
		{
			oMenuPos = this.GetThumbnailPagePosition(Math.min.apply(Math, sSelectedIdx));
			bIsSlideSelect = sSelectedIdx.length > 0;
		}
		else
		{
			let oEditorCtrl = this.m_oWordControl;
			let oThCtrlPos = oEditorCtrl.m_oThumbnails.AbsolutePosition;
			oMenuPos = {
				X: global_mouseEvent.X - ((oThCtrlPos.L * g_dKoef_mm_to_pix) >> 0) - oEditorCtrl.X,
				Y: global_mouseEvent.Y - ((oThCtrlPos.T * g_dKoef_mm_to_pix) >> 0) - oEditorCtrl.Y
			};

			let pos = this.ConvertCoords(global_mouseEvent.X, global_mouseEvent.Y);
			bIsSlideSelect = pos.Page !== -1;
		}
		if (oMenuPos)
		{
			const oLogicDocument = this.m_oWordControl.m_oLogicDocument;
			let oFirstSlide = oLogicDocument.GetSlide(sSelectedIdx[0]);
			let nType = Asc.c_oAscContextMenuTypes.Thumbnails;
			if(oFirstSlide)
			{
				if(oFirstSlide.getObjectType() === AscDFH.historyitem_type_SlideLayout)
				{
					nType = Asc.c_oAscContextMenuTypes.Layout;
				}
				else if(oFirstSlide.getObjectType() === AscDFH.historyitem_type_SlideMaster)
				{
					nType = Asc.c_oAscContextMenuTypes.Master;
				}
			}
			let oData =
				{
					Type: nType,
					X_abs: oMenuPos.X,
					Y_abs: oMenuPos.Y,
					IsSlideSelect: bIsSlideSelect,
					IsSlideHidden: this.IsSlideHidden(sSelectedIdx),
					IsSlidePreserve: oLogicDocument.isPreserveSelectionSlides()
				};
			editor.sync_ContextMenuCallback(new AscCommonSlide.CContextMenuData(oData));
		}
	};

	CThumbnailsManager.prototype.GetCurSld = function()
	{
		return this.thumbnails.GetCurSld();
	};


	function COutlineThumbnailsManager(editorPage) {
		CThumbnailsManagerBase.call(this, editorPage);
		this.isInit = false;
		this.lastPixelRatio = 0;
		this.m_oFontManager = new AscFonts.CFontManager();

		this.m_bIsScrollVisible = true;

		// cached measure
		this.DigitWidths = [];

		// skin
		this.backgroundColor = "#B0B0B0";

		this.const_offset_x = 0;
		this.const_offset_y = 0;
		this.const_offset_r = 4;
		this.const_offset_b = 0;
		this.const_border_w = 4;

		// size & position
		this.SlideWidth = 297;
		this.SlideHeight = 210;

		this.m_dScrollY = 0;
		this.m_dScrollY_max = 0;

		this.m_bIsUpdate = false;

		//regular mode
		this.thumbnails = new AscCommon.CSlidesThumbnails();

		this.m_nCurrentPage = -1;
		this.m_arrPages = [];
		this.m_lDrawingFirst = -1;
		this.m_lDrawingEnd = -1;

		this.bIsEmptyDrawed = false;

		this.m_oCacheManager = new CCacheManager(true);

		this.FocusObjType = FOCUS_OBJECT_MAIN;
		this.LockMainObjType = false;

		this.SelectPageEnabled = true;

		this.MouseDownTrack = new AscCommon.CMouseDownTrack(this);

		this.MouseTrackCommonImage = null;

		this.MouseThumbnailsAnimateScrollTopTimer = -1;
		this.MouseThumbnailsAnimateScrollBottomTimer = -1;

		this.ScrollerHeight = 0;
		this.ScrollerWidth = 0;
	}
	AscFormat.InitClassWithoutType(COutlineThumbnailsManager, CThumbnailsManagerBase);

	COutlineThumbnailsManager.prototype.Init = function()
	{
		this.isInit = true;
		this.m_oFontManager.Initialize(true);
		this.m_oFontManager.SetHintsProps(true, true);

		this.InitCheckOnResize();

		this.MouseTrackCommonImage = document.createElement("canvas");

		var _im_w = 9;
		var _im_h = 9;

		this.MouseTrackCommonImage.width = _im_w;
		this.MouseTrackCommonImage.height = _im_h;

		var _ctx = this.MouseTrackCommonImage.getContext('2d');
		var _data = _ctx.createImageData(_im_w, _im_h);
		var _pixels = _data.data;

		var _ind = 0;
		for (var j = 0; j < _im_h; ++j)
		{
			var _off1 = (j > (_im_w >> 1)) ? (_im_w - j - 1) : j;
			var _off2 = _im_w - _off1 - 1;

			for (var r = 0; r < _im_w; ++r)
			{
				if (r <= _off1 || r >= _off2)
				{
					_pixels[_ind++] = 183;
					_pixels[_ind++] = 183;
					_pixels[_ind++] = 183;
					_pixels[_ind++] = 255;
				} else
				{
					_pixels[_ind + 3] = 0;
					_ind += 4;
				}
			}
		}

		_ctx.putImageData(_data, 0, 0);
	};

	COutlineThumbnailsManager.prototype.InitCheckOnResize = function()
	{
		if (!this.isInit)
			return;

		if (Math.abs(this.lastPixelRatio - AscCommon.AscBrowser.retinaPixelRatio) < 0.01)
			return;

		this.lastPixelRatio = AscCommon.AscBrowser.retinaPixelRatio;
		this.SetFont({
			FontFamily: {Name: "Arial", Index: -1},
			Italic: false,
			Bold: false,
			FontSize: Math.round(10 * AscCommon.AscBrowser.retinaPixelRatio)
		});

		// измеряем все цифры
		for (var i = 0; i < 10; i++)
		{
			var _meas = this.m_oFontManager.MeasureChar(("" + i).charCodeAt(0));
			if (_meas)
				this.DigitWidths[i] = _meas.fAdvanceX * 25.4 / 96;
			else
				this.DigitWidths[i] = 10;
		}

		if (Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom) {
			this.const_offset_x = Math.round(this.lastPixelRatio * 17);
			this.const_offset_r = this.const_offset_y;
			this.const_offset_b = Math.round(this.lastPixelRatio * 10);
			this.const_border_w = Math.round(this.lastPixelRatio * 7);
		} else {
			this.const_offset_y = Math.round(this.lastPixelRatio * 17);
			this.const_offset_b = this.const_offset_y;
			this.const_offset_r = Math.round(this.lastPixelRatio * 10);
			this.const_border_w = Math.round(this.lastPixelRatio * 7);
		}

	};

	COutlineThumbnailsManager.prototype.GetPageByPos = function(oPos)
	{
		return this.thumbnails.GetPage(oPos);
	};

	COutlineThumbnailsManager.prototype.GetSlidesCount = function()
	{
		return Asc.editor.getCountSlides();
	};

	COutlineThumbnailsManager.prototype.SetFont = function(font)
	{
		font.FontFamily.Name = g_fontApplication.GetFontFileWeb(font.FontFamily.Name, 0).m_wsFontName;

		if (-1 == font.FontFamily.Index)
			font.FontFamily.Index = AscFonts.g_map_font_index[font.FontFamily.Name];

		if (font.FontFamily.Index == undefined || font.FontFamily.Index == -1)
			return;

		var bItalic = true === font.Italic;
		var bBold   = true === font.Bold;

		var oFontStyle = FontStyle.FontStyleRegular;
		if (!bItalic && bBold)
			oFontStyle = FontStyle.FontStyleBold;
		else if (bItalic && !bBold)
			oFontStyle = FontStyle.FontStyleItalic;
		else if (bItalic && bBold)
			oFontStyle = FontStyle.FontStyleBoldItalic;

		g_fontApplication.LoadFont(font.FontFamily.Name, AscCommon.g_font_loader, this.m_oFontManager, font.FontSize, oFontStyle, 96, 96);
	};

	COutlineThumbnailsManager.prototype.SetSlideRecalc = function (nIdx)
	{
		this.thumbnails.SetSlideRecalc(nIdx);
	};

	COutlineThumbnailsManager.prototype.SelectAll = function()
	{
		this.thumbnails.SelectAll();
		this.OnUpdateOverlay();
	};

	COutlineThumbnailsManager.prototype.GetFirstSelectedType = function()
	{
		return this.m_oWordControl.m_oLogicDocument.GetFirstSelectedType();
	};
	COutlineThumbnailsManager.prototype.GetSlideType = function(nIdx)
	{
		return this.m_oWordControl.m_oLogicDocument.GetSlideType(nIdx);
	};
	COutlineThumbnailsManager.prototype.IsMixedSelection = function()
	{
		return this.m_oWordControl.m_oLogicDocument.IsMixedSelection();
	};

	// events
	COutlineThumbnailsManager.prototype.onMouseDown = function (e) {
		const mobileTouchManager = this.m_oWordControl ? this.m_oWordControl.MobileTouchManagerThumbnails : null;
		if (mobileTouchManager && mobileTouchManager.checkTouchEvent(e, true)) {
			mobileTouchManager.startTouchingInProcess();
			const res = mobileTouchManager.mainOnTouchStart(e);
			mobileTouchManager.stopTouchingInProcess();
			return res;
		}

		if (this.m_oWordControl) {
			this.m_oWordControl.m_oApi.checkInterfaceElementBlur();
			this.m_oWordControl.m_oApi.checkLastWork();

			// после fullscreen возможно изменение X, Y после вызова Resize.
			this.m_oWordControl.checkBodyOffset();
		}

		AscCommon.stopEvent(e);

		if (AscCommon.g_inputContext && AscCommon.g_inputContext.externalChangeFocus())
			return;

		var control = this.m_oWordControl.m_oThumbnails.HtmlElement;
		if (global_mouseEvent.IsLocked == true && global_mouseEvent.Sender != control) {
			// кто-то зажал мышку. кто-то другой
			return false;
		}

		AscCommon.check_MouseDownEvent(e);
		global_mouseEvent.LockMouse();
		const oPresentation = this.m_oWordControl.m_oLogicDocument;
		let nStartHistoryIndex = oPresentation.History.Index;
		function checkSelectionEnd() {
			if (nStartHistoryIndex === oPresentation.History.Index) {
				Asc.editor.sendEvent("asc_onSelectionEnd");
			}
		}

		this.m_oWordControl.m_oApi.sync_EndAddShape();
		if (global_mouseEvent.Sender != control) {
			// такого быть не должно
			return false;
		}

		if (global_mouseEvent.Button == undefined)
			global_mouseEvent.Button = 0;

		this.SetFocusElement(FOCUS_OBJECT_THUMBNAILS);

		var pos = this.ConvertCoords(global_mouseEvent.X, global_mouseEvent.Y);
		if (pos.Page == -1) {
			if (global_mouseEvent.Button == 2) {
				this.showContextMenu(false);
			}
			checkSelectionEnd();
			return false;
		}

		if (global_keyboardEvent.CtrlKey && !this.m_oWordControl.m_oApi.isReporterMode) {
			if (this.m_arrPages[pos.Page].IsSelected === true) {
				this.m_arrPages[pos.Page].IsSelected = false;
				var arr = this.GetSelectedArray();
				if (0 == arr.length) {
					this.m_arrPages[pos.Page].IsSelected = true;
					this.ShowPage(pos.Page);
				} else {
					this.OnUpdateOverlay();

					this.SelectPageEnabled = false;
					this.m_oWordControl.GoToPage(arr[0]);
					this.SelectPageEnabled = true;

					this.ShowPage(arr[0]);
				}
			} else {
				if (this.GetFirstSelectedType() === this.GetSlideType(pos.Page)) {
					this.m_arrPages[pos.Page].IsSelected = true;
					this.OnUpdateOverlay();

					this.SelectPageEnabled = false;
					this.m_oWordControl.GoToPage(pos.Page);
					this.SelectPageEnabled = true;
					this.ShowPage(pos.Page);
				}
			}
		} else if (global_keyboardEvent.ShiftKey && !this.m_oWordControl.m_oApi.isReporterMode) {

			var pages_count = this.m_arrPages.length;
			for (var i = 0; i < pages_count; i++) {
				this.m_arrPages[i].IsSelected = false;
			}

			var _max = pos.Page;
			var _min = this.m_oWordControl.m_oDrawingDocument.SlideCurrent;
			if (_min > _max) {
				var _temp = _max;
				_max = _min;
				_min = _temp;
			}

			let nSlideType = this.GetSlideType(_min);
			for (var i = _min; i <= _max; i++) {
				if (nSlideType === this.GetSlideType(i)) {
					this.m_arrPages[i].IsSelected = true;
				}
			}

			this.OnUpdateOverlay();
			this.ShowPage(pos.Page);
			oPresentation.Document_UpdateInterfaceState();
		} else if (0 == global_mouseEvent.Button || 2 == global_mouseEvent.Button) {

			let isMouseDownOnAnimPreview = false;
			if (0 == global_mouseEvent.Button) {
				if (this.m_arrPages[pos.Page].animateLabelRect) // click on the current star, preview animation button, slide transition
				{
					let animateLabelRect = this.m_arrPages[pos.Page].animateLabelRect;
					if (pos.X >= animateLabelRect.minX && pos.X <= animateLabelRect.maxX && pos.Y >= animateLabelRect.minY && pos.Y <= animateLabelRect.maxY)
						isMouseDownOnAnimPreview = true
				}

				if (!isMouseDownOnAnimPreview) // приготавливаемся к треку
				{
					this.MouseDownTrack.Start(pos.Page, global_mouseEvent.X, global_mouseEvent.Y);
				}
			}

			if (this.m_arrPages[pos.Page].IsSelected) {
				let isStartedAnimPreview = (this.m_oWordControl.m_oLogicDocument.IsStartedPreview() || (this.m_oWordControl.m_oDrawingDocument.TransitionSlide && this.m_oWordControl.m_oDrawingDocument.TransitionSlide.IsPlaying()));

				this.SelectPageEnabled = false;
				this.m_oWordControl.GoToPage(pos.Page);
				this.SelectPageEnabled = true;

				if (!isStartedAnimPreview) {
					if (isMouseDownOnAnimPreview) {
						this.m_oWordControl.m_oApi.SlideTransitionPlay(function () { this.m_oWordControl.m_oApi.asc_StartAnimationPreview() });
					}
				}

				if (this.m_oWordControl.m_oNotesApi.IsEmptyDraw) {
					this.m_oWordControl.m_oNotesApi.IsEmptyDraw = false;
					this.m_oWordControl.m_oNotesApi.IsRepaint = true;
				}

				if (global_mouseEvent.Button == 2 && !global_keyboardEvent.CtrlKey) {
					this.showContextMenu(false);
				}
				checkSelectionEnd();
				return false;
			}

			var pages_count = this.m_arrPages.length;
			for (var i = 0; i < pages_count; i++) {
				this.m_arrPages[i].IsSelected = false;
			}

			this.m_arrPages[pos.Page].IsSelected = true;

			this.OnUpdateOverlay();

			if (global_mouseEvent.Button == 0 && this.m_arrPages[pos.Page].animateLabelRect) // click on the current star, preview animation button, slide transition
			{
				let animateLabelRect = this.m_arrPages[pos.Page].animateLabelRect;
				let isMouseDownOnAnimPreview = false;
				if (pos.X >= animateLabelRect.minX && pos.X <= animateLabelRect.maxX && pos.Y >= animateLabelRect.minY && pos.Y <= animateLabelRect.maxY);
				isMouseDownOnAnimPreview = true;
			}

			this.SelectPageEnabled = false;
			this.m_oWordControl.GoToPage(pos.Page);
			this.SelectPageEnabled = true;

			if (isMouseDownOnAnimPreview) {
				this.m_oWordControl.m_oApi.SlideTransitionPlay(function () { this.m_oWordControl.m_oApi.asc_StartAnimationPreview() });
			}

			this.ShowPage(pos.Page);
		}

		if (global_mouseEvent.Button == 2 && !global_keyboardEvent.CtrlKey) {
			this.showContextMenu(false);
		}
		checkSelectionEnd();
		return false;
	};

	COutlineThumbnailsManager.prototype.onMouseMove = function(e)
	{
		let mobileTouchManager = this.m_oWordControl ? this.m_oWordControl.MobileTouchManagerThumbnails : null;
		if (mobileTouchManager && mobileTouchManager.checkTouchEvent(e, true))
		{
			mobileTouchManager.startTouchingInProcess();
			let res = mobileTouchManager.mainOnTouchMove(e);
			mobileTouchManager.stopTouchingInProcess();
			return res;
		}

		const isHorizontalThumbnails = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;

		if (this.m_oWordControl)
			this.m_oWordControl.m_oApi.checkLastWork();

		var control = this.m_oWordControl.m_oThumbnails.HtmlElement;
		if (global_mouseEvent.IsLocked == true && global_mouseEvent.Sender != control)
		{
			// кто-то зажал мышку. кто-то другой
			return;
		}
		AscCommon.check_MouseMoveEvent(e);
		if (global_mouseEvent.Sender != control)
		{
			// такого быть не должно
			return;
		}

		if (this.MouseDownTrack.IsStarted())
		{
			// это трек для перекидывания слайдов
			if (this.MouseDownTrack.IsSimple() && !this.m_oWordControl.m_oApi.isViewMode)
			{
				if (Math.abs(this.MouseDownTrack.GetX() - global_mouseEvent.X) > 10 || Math.abs(this.MouseDownTrack.GetY() - global_mouseEvent.Y) > 10)
					this.MouseDownTrack.ResetSimple(this.ConvertCoords2(global_mouseEvent.X, global_mouseEvent.Y));
			}
			else
			{
				if (!this.MouseDownTrack.IsSimple())
				{
					// нужно определить активная позиция между слайдами
					this.MouseDownTrack.SetPosition(this.ConvertCoords2(global_mouseEvent.X, global_mouseEvent.Y));
				}
			}


			this.OnUpdateOverlay();

			// теперь нужно посмотреть, нужно ли проскроллить
			if (this.m_bIsScrollVisible) {
				const _X = global_mouseEvent.X - this.m_oWordControl.X;
				const _Y = global_mouseEvent.Y - this.m_oWordControl.Y;

				const _abs_pos = this.m_oWordControl.m_oThumbnails.AbsolutePosition;
				const _XMax = (_abs_pos.R - _abs_pos.L) * g_dKoef_mm_to_pix;
				const _YMax = (_abs_pos.B - _abs_pos.T) * g_dKoef_mm_to_pix;

				const mousePos = isHorizontalThumbnails ? _X : _Y;
				const maxPos = isHorizontalThumbnails ? _XMax : _YMax;

				let _check_type = -1;
				if (mousePos < 30) {
					_check_type = 0;
				} else if (mousePos >= (maxPos - 30)) {
					_check_type = 1;
				}

				this.CheckNeedAnimateScrolls(_check_type);
			}

			if (!this.MouseDownTrack.IsSimple())
			{
				var cursor_dragged = "default";
				if (AscCommon.AscBrowser.isWebkit)
					cursor_dragged = "-webkit-grabbing";
				else if (AscCommon.AscBrowser.isMozilla)
					cursor_dragged = "-moz-grabbing";

				this.m_oWordControl.m_oThumbnails.HtmlElement.style.cursor = cursor_dragged;
			}

			return;
		}

		var pos = this.ConvertCoords(global_mouseEvent.X, global_mouseEvent.Y);

		var _is_old_focused = false;

		var pages_count = this.m_arrPages.length;
		for (var i = 0; i < pages_count; i++)
		{
			if (this.m_arrPages[i].IsFocused)
			{
				_is_old_focused = true;
				this.m_arrPages[i].IsFocused = false;
			}
		}

		var cursor_moved = "default";

		if (pos.Page != -1)
		{
			this.m_arrPages[pos.Page].IsFocused = true;
			this.OnUpdateOverlay();

			cursor_moved = "pointer";
		} else if (_is_old_focused)
		{
			this.OnUpdateOverlay();
		}

		this.m_oWordControl.m_oThumbnails.HtmlElement.style.cursor = cursor_moved;
	};

	COutlineThumbnailsManager.prototype.onMouseUp = function(e, bIsWindow)
	{
		let mobileTouchManager = this.m_oWordControl ? this.m_oWordControl.MobileTouchManagerThumbnails : null;
		if (mobileTouchManager && mobileTouchManager.checkTouchEvent(e, true))
		{
			mobileTouchManager.startTouchingInProcess();
			let res = mobileTouchManager.mainOnTouchEnd(e);
			mobileTouchManager.stopTouchingInProcess();
			return res;
		}

		if (this.m_oWordControl)
			this.m_oWordControl.m_oApi.checkLastWork();

		var _oldSender = global_mouseEvent.Sender;
		AscCommon.check_MouseUpEvent(e);
		global_mouseEvent.UnLockMouse();

		var control = this.m_oWordControl.m_oThumbnails.HtmlElement;
		if (global_mouseEvent.Sender != control)
		{
			if (_oldSender != control || true !== bIsWindow)
			{
				// такого быть не должно
				return;
			}
		}

		this.CheckNeedAnimateScrolls(-1);

		if (!this.MouseDownTrack.IsStarted())
			return;

		// теперь смотрим, просто ли это селект, или же это трек
		if (this.MouseDownTrack.IsSimple())
		{
			if (this.MouseDownTrack.IsMoved(global_mouseEvent.X, global_mouseEvent.Y))
				this.MouseDownTrack.ResetSimple(this.ConvertCoords2(global_mouseEvent.X, global_mouseEvent.Y));
		}

		if (this.MouseDownTrack.IsSimple())
		{
			// это просто селект
			var pages_count = this.m_arrPages.length;
			for (var i = 0; i < pages_count; i++)
			{
				this.m_arrPages[i].IsSelected = false;
			}

			this.m_arrPages[this.MouseDownTrack.GetPage()].IsSelected = true;

			this.OnUpdateOverlay();

			// послали уже на mouseDown
			//this.SelectPageEnabled = false;
			//this.m_oWordControl.GoToPage(this.MouseDownTrack.GetPage());
			//this.SelectPageEnabled = true;
		} else
		{
			// это трек
			this.MouseDownTrack.SetPosition(this.ConvertCoords2(global_mouseEvent.X, global_mouseEvent.Y));

			if (-1 !== this.MouseDownTrack.GetPosition() && (!this.MouseDownTrack.IsSamePos() || AscCommon.global_mouseEvent.CtrlKey))
			{
				// вызвать функцию апи для смены слайдов местами
				var _array = this.GetSelectedArray();
				this.m_oWordControl.m_oLogicDocument.shiftSlides(this.MouseDownTrack.GetPosition(), _array);
				this.ClearCacheAttack();
			}

			this.OnUpdateOverlay();
		}
		this.MouseDownTrack.Reset();

		this.onMouseMove(e);
	};

	COutlineThumbnailsManager.prototype.onMouseLeave = function(e)
	{
		var pages_count = this.m_arrPages.length;
		for (var i = 0; i < pages_count; i++)
		{
			this.m_arrPages[i].IsFocused = false;
		}
		this.OnUpdateOverlay();
	};

	COutlineThumbnailsManager.prototype.onMouseWhell = function (e) {
		if (false === this.m_bIsScrollVisible || !this.m_oWordControl.m_oScrollThumbApi)
			return;

		if (global_keyboardEvent.CtrlKey)
			return;

		if (undefined !== window["AscDesktopEditor"] && false === window["AscDesktopEditor"]["CheckNeedWheel"]())
			return;

		const delta = GetWheelDeltaY(e);
		const isHorizontalOrientation = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;
		const isRightToLeft = Asc.editor.isRtlInterface;
		isHorizontalOrientation
			? this.m_oWordControl.m_oScrollThumbApi.scrollBy(isRightToLeft ? -delta : delta, 0)
			: this.m_oWordControl.m_oScrollThumbApi.scrollBy(0, delta, false);

		if (e.preventDefault)
			e.preventDefault();
		else
			e.returnValue = false;

		AscCommon.stopEvent(e);
		return false;
	};

	COutlineThumbnailsManager.prototype.onKeyDown = function(e)
	{
		var control = this.m_oWordControl.m_oThumbnails.HtmlElement;
		if (global_mouseEvent.IsLocked == true && global_mouseEvent.Sender == control)
		{
			e.preventDefault();
			return false;
		}
		AscCommon.check_KeyboardEvent(e);
		let oApi = this.m_oWordControl.m_oApi;
		let oPresentation = this.m_oWordControl.m_oLogicDocument;
		let oDrawingDocument = this.m_oWordControl.m_oDrawingDocument;
		let oEvent = global_keyboardEvent;
		let nShortCutAction = oApi.getShortcut(oEvent);
		let bReturnValue = false, bPreventDefault = true;
		let sSelectedIdx;
		let nStartHistoryIndex = oPresentation.History.Index;
		switch (nShortCutAction)
		{
			case Asc.c_oAscPresentationShortcutType.EditSelectAll:
			{
				this.SelectAll();
				bReturnValue = false;
				bPreventDefault = true;
				break;
			}
			case Asc.c_oAscPresentationShortcutType.EditUndo:
			case Asc.c_oAscPresentationShortcutType.EditRedo:
			{
				bReturnValue = true;
				bPreventDefault = false;
				break;
			}
			case Asc.c_oAscPresentationShortcutType.Duplicate:
			{
				if (oPresentation.CanEdit())
				{
					oApi.DublicateSlide();
				}
				bReturnValue = false;
				bPreventDefault = true;
				break;
			}
			case Asc.c_oAscPresentationShortcutType.PrintPreviewAndPrint:
			{
				oApi.onPrint();
				bReturnValue = false;
				bPreventDefault = true;
				break;
			}
			case Asc.c_oAscPresentationShortcutType.Save:
			{
				if (!oPresentation.IsViewMode())
				{
					oApi.asc_Save();
				}
				bReturnValue = false;
				bPreventDefault = true;
				break;
			}
			case Asc.c_oAscPresentationShortcutType.OpenContextMenu:
			{
				this.showContextMenu(true);
				bReturnValue = false;
				bPreventDefault = false;
				break;
			}
			default:
			{
				break;
			}
		}
		if (!nShortCutAction)
		{
			if (AscCommon.isCopyPasteEvent(oEvent)) {
				return;
			}
			switch (oEvent.KeyCode)
			{
				case 13: // enter
				{
					if (oPresentation.CanEdit())
					{
						sSelectedIdx = this.GetSelectedArray();
						bReturnValue = false;
						bPreventDefault = true;
						if (sSelectedIdx.length > 0)
						{
							this.m_oWordControl.GoToPage(sSelectedIdx[sSelectedIdx.length - 1]);
							oPresentation.addNextSlide();
							bPreventDefault = false;
						}
					}
					break;
				}
				case 46: // delete
				case 8: // backspace
				{
					if (oPresentation.CanEdit())
					{
						const arrSlides = oPresentation.GetSelectedSlideObjects();
						if (!oApi.IsSupportEmptyPresentation)
						{
							if (arrSlides.length === oDrawingDocument.GetSlidesCount())
							{
								arrSlides.splice(0, 1);
							}
						}
						if (arrSlides.length !== 0)
						{
							oPresentation.deleteSlides(arrSlides);
						}
						if (0 === oPresentation.GetSlidesCount())
						{
							this.m_bIsUpdate = true;
						}
						bReturnValue = false;
						bPreventDefault = true;
					}
					break;
				}
				case 34: // PgDown
				case 40: // bottom arrow
				{
					if (oEvent.CtrlKey && oEvent.ShiftKey)
					{
						oPresentation.moveSelectedSlidesToEnd();
					} else if (oEvent.CtrlKey)
					{
						oPresentation.moveSlidesNextPos();
					} else if (oEvent.ShiftKey)
					{
						this.CorrectShiftSelect(false, false);
					} else
					{
						if (oDrawingDocument.SlideCurrent < oDrawingDocument.GetSlidesCount() - 1)
						{
							this.m_oWordControl.GoToPage(oDrawingDocument.SlideCurrent + 1);
						}
					}
					break;
				}
				case 36: // home
				{
					if (!oEvent.ShiftKey)
					{
						if (oDrawingDocument.SlideCurrent > 0)
						{
							this.m_oWordControl.GoToPage(0);
						}
					} else
					{
						if (oDrawingDocument.SlideCurrent > 0)
						{
							this.CorrectShiftSelect(true, true);
						}
					}
					break;
				}
				case 35: // end
				{
					var slidesCount = oDrawingDocument.GetSlidesCount();
					if (!oEvent.ShiftKey)
					{
						if (oDrawingDocument.SlideCurrent !== (slidesCount - 1))
						{
							this.m_oWordControl.GoToPage(slidesCount - 1);
						}
					} else
					{
						if (oDrawingDocument.SlideCurrent !== (slidesCount - 1))
						{
							this.CorrectShiftSelect(false, true);
						}
					}
					break;
				}
				case 33: // PgUp
				case 38: // UpArrow
				{
					if (oEvent.CtrlKey && oEvent.ShiftKey)
					{
						oPresentation.moveSelectedSlidesToStart();
					} else if (oEvent.CtrlKey)
					{
						oPresentation.moveSlidesPrevPos();
					} else if (oEvent.ShiftKey)
					{
						this.CorrectShiftSelect(true, false);
					} else
					{
						if (oDrawingDocument.SlideCurrent > 0)
						{
							this.m_oWordControl.GoToPage(oDrawingDocument.SlideCurrent - 1);
						}
					}
					break;
				}
				case 77:
				{
					if (oEvent.CtrlKey)
					{
						if (oPresentation.CanEdit())
						{
							sSelectedIdx = this.GetSelectedArray();
							if (sSelectedIdx.length > 0)
							{
								this.m_oWordControl.GoToPage(sSelectedIdx[sSelectedIdx.length - 1]);
								oPresentation.addNextSlide();
								bReturnValue = false;
							} else if (this.GetSlidesCount() === 0)
							{
								oPresentation.addNextSlide();
								this.m_oWordControl.GoToPage(0);
							}
						}
					}
					bReturnValue = false;
					bPreventDefault = true;
					break;
				}
				case 122:
				case 123:
				{
					return;
				}
				default:
					break;
			}
		}
		if (bPreventDefault)
		{
			e.preventDefault();
		}
		if(nStartHistoryIndex === oPresentation.History.Index)
		{
			oApi.sendEvent("asc_onSelectionEnd");
		}
		return bReturnValue;
	};

	COutlineThumbnailsManager.prototype.SetFocusElement = function(type)
	{
		if(this.FocusObjType === type)
			return;
		switch (type)
		{
			case FOCUS_OBJECT_MAIN:
			{
				this.FocusObjType = FOCUS_OBJECT_MAIN;
				break;
			}
			case FOCUS_OBJECT_THUMBNAILS:
			{
				if (this.LockMainObjType)
					return;

				this.FocusObjType = FOCUS_OBJECT_THUMBNAILS;
				if (this.m_oWordControl.m_oLogicDocument)
				{
					this.m_oWordControl.m_oLogicDocument.resetStateCurSlide(true);
				}
				break;
			}
			case FOCUS_OBJECT_ANIM_PANE:
			{
				this.FocusObjType = FOCUS_OBJECT_ANIM_PANE;
				break;
			}
			case FOCUS_OBJECT_NOTES:
			{
				break;
			}
			default:
				break;
		}
	};

	// draw
	COutlineThumbnailsManager.prototype.ShowPage = function (pageNum) {
		const scrollApi = this.m_oWordControl.m_oScrollThumbApi;
		if (!scrollApi) return;

		let startCoord, endCoord, visibleAreaSize, scrollTo, scrollBy;
		if (Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom) {
			startCoord = this.m_arrPages[pageNum].left - this.const_border_w;
			endCoord = this.m_arrPages[pageNum].right + this.const_border_w;
			visibleAreaSize = this.m_oWordControl.m_oThumbnails.HtmlElement.width;

			scrollTo = scrollApi.scrollToX.bind(scrollApi);
			scrollBy = scrollApi.scrollByX.bind(scrollApi);
		} else {
			startCoord = this.m_arrPages[pageNum].top - this.const_border_w;
			endCoord = this.m_arrPages[pageNum].bottom + this.const_border_w;
			visibleAreaSize = this.m_oWordControl.m_oThumbnails.HtmlElement.height;

			scrollTo = scrollApi.scrollToY.bind(scrollApi);
			scrollBy = scrollApi.scrollByY.bind(scrollApi);
		}

		if (startCoord < 0) {
			const size = endCoord - startCoord;
			const shouldReversePageIndexes = Asc.editor.isRtlInterface &&
				Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;
			const pos = shouldReversePageIndexes
				? (size + this.const_border_w) * (this.m_arrPages.length - pageNum - 1)
				: (size + this.const_border_w) * pageNum;
			scrollTo(pos);
		} else if (endCoord > visibleAreaSize) {
			scrollBy(endCoord + this.const_border_w - visibleAreaSize);
		}
	};

	COutlineThumbnailsManager.prototype.ClearCacheAttack = function()
	{
		var pages_count = this.m_arrPages.length;

		for (var i = 0; i < pages_count; i++)
		{
			this.m_arrPages[i].IsRecalc = true;
		}

		this.m_bIsUpdate = true;
	};

	COutlineThumbnailsManager.prototype.FocusRectDraw = function(ctx, x, y, r, b)
	{
		ctx.rect(x - this.const_border_w, y, r - x + this.const_border_w, b - y);
	};
	COutlineThumbnailsManager.prototype.FocusRectFlat = function(_color, ctx, x, y, r, b, isFocus)
	{
		ctx.beginPath();
		ctx.strokeStyle = _color;
		var lineW = Math.round((isFocus ? 2 : 3) * AscCommon.AscBrowser.retinaPixelRatio);
		var dist = Math.round(2 * AscCommon.AscBrowser.retinaPixelRatio);
		ctx.lineWidth = lineW;
		var extend = dist + (lineW / 2);

		ctx.rect(x - extend, y - extend, r - x + (2 * extend), b - y + (2 * extend));
		ctx.stroke();

		ctx.beginPath();
	};
	COutlineThumbnailsManager.prototype.OutlineRect = function(_color, ctx, x, y, r, b)
	{
		ctx.beginPath();
		ctx.strokeStyle = "#C0C0C0";
		const lineW = Math.max(1, Math.round(1 * AscCommon.AscBrowser.retinaPixelRatio));
		ctx.lineWidth = lineW;
		const extend = lineW / 2;
		ctx.rect(x - extend, y - extend, r - x + (2 * extend), b - y + (2 * extend));
		ctx.stroke();

		ctx.beginPath();
	};

	COutlineThumbnailsManager.prototype.OnPaint = function () {
		if (!this.isThumbnailsShown()) {
			return;
		}

		const canvas = this.m_oWordControl.m_oThumbnails.HtmlElement;
		if (!canvas)
			return;

		this.m_oWordControl.m_oApi.clearEyedropperImgData();

		const context = AscCommon.AscBrowser.getContext2D(canvas);
		context.clearRect(0, 0, canvas.width, canvas.height);
		const graphics = new AscCommon.CGraphics();
		const widthMM = canvas.width * g_dKoef_pix_to_mm;
		const heightMM = canvas.height * g_dKoef_pix_to_mm;
		graphics.init(context, canvas.width, canvas.height, widthMM, heightMM);
		graphics.m_oFontManager = this.m_oFontManager;
		graphics.transform(1, 0, 0, 1, 0, 0);

		const presentation = this.m_oWordControl.m_oLogicDocument;
		const outlineView = new AscCommonSlide.OutlineView(presentation);
		outlineView.updateAll(widthMM, heightMM);
		outlineView.draw(graphics);
		// this.OnUpdateOverlay();
	};

	COutlineThumbnailsManager.prototype.onCheckUpdate = function()
	{
		if (!this.isThumbnailsShown() || 0 == this.DigitWidths.length)
			return;

		// if (!this.isThumbnailsShown() || 0 == this.DigitWidths.length)
		// 	return;
		//
		//
		// if (this.m_oWordControl.m_oApi.isSaveFonts_Images)
		// 	return;
		//
		// if (this.m_lDrawingFirst == -1 || this.m_lDrawingEnd == -1)
		// {
		// 	if (this.m_oWordControl.m_oDrawingDocument.IsEmptyPresentation)
		// 	{
		// 		if (!this.bIsEmptyDrawed)
		// 		{
		// 			this.bIsEmptyDrawed = true;
		// 			this.OnPaint();
		// 		}
		// 		return;
		// 	}
		//
		// 	this.bIsEmptyDrawed = false;
		// 	return;
		// }
		//
		// this.bIsEmptyDrawed = false;
		//
		// if (!this.m_bIsUpdate)
		// {
		// 	// определяем, нужно ли пересчитать и закэшировать табнейл (хотя бы один)
		// 	for (var i = this.m_lDrawingFirst; i <= this.m_lDrawingEnd; i++)
		// 	{
		// 		var page = this.m_arrPages[i];
		// 		if (null == page.cachedImage || page.IsRecalc)
		// 		{
		// 			this.m_bIsUpdate = true;
		// 			break;
		// 		}
		// 		if ((page.cachedImage.image.width != (page.right - page.left)) || (page.cachedImage.image.height != (page.bottom - page.top)))
		// 		{
		// 			this.m_bIsUpdate = true;
		// 			break;
		// 		}
		// 	}
		// }
		//
		// if (!this.m_bIsUpdate)
		// 	return;
		//
		// for (var i = this.m_lDrawingFirst; i <= this.m_lDrawingEnd; i++)
		// {
		// 	var page = this.m_arrPages[i];
		// 	var w = page.right - page.left;
		// 	var h = page.bottom - page.top;
		//
		// 	if (null != page.cachedImage)
		// 	{
		// 		if ((page.cachedImage.image.width != w) || (page.cachedImage.image.height != h) || page.IsRecalc)
		// 		{
		// 			this.m_oCacheManager.UnLock(page.cachedImage);
		// 			page.cachedImage = null;
		// 		}
		// 	}
		//
		// 	if (null == page.cachedImage)
		// 	{
		// 		page.cachedImage = this.m_oCacheManager.Lock(w, h);
		//
		// 		var g = new AscCommon.CGraphics();
		// 		g.IsNoDrawingEmptyPlaceholder = true;
		// 		g.IsThumbnail = true;
		// 		if (this.m_oWordControl.m_oLogicDocument.IsVisioEditor())
		// 		{
		// 			const sizes = this.m_oWordControl.m_oLogicDocument.GetSizesMM(i);
		// 			//todo override COutlineThumbnailsManager
		// 			let SlideWidth = sizes.width;
		// 			let SlideHeight = sizes.height;
		// 			g.init(page.cachedImage.image.ctx, w, h, SlideWidth, SlideHeight);
		// 		}
		// 		else
		// 		{
		// 			g.init(page.cachedImage.image.ctx, w, h, this.SlideWidth, this.SlideHeight);
		// 		}
		// 		g.m_oFontManager = this.m_oFontManager;
		//
		// 		g.transform(1, 0, 0, 1, 0, 0);
		//
		// 		var bIsShowPars = this.m_oWordControl.m_oApi.ShowParaMarks;
		// 		this.m_oWordControl.m_oApi.ShowParaMarks = false;
		// 		this.m_oWordControl.m_oLogicDocument.DrawPage(i, g);
		// 		this.m_oWordControl.m_oApi.ShowParaMarks = bIsShowPars;
		//
		// 		page.IsRecalc = false;
		//
		// 		this.m_bIsUpdate = true;
		// 		break;
		// 	}
		// }

		this.OnPaint();
		this.m_bIsUpdate = false;
	};

	COutlineThumbnailsManager.prototype.OnUpdateOverlay = function () {
		if (!this.isThumbnailsShown())
			return;

		const canvas = this.m_oWordControl.m_oThumbnailsBack.HtmlElement;
		if (canvas == null)
			return;

		if (this.m_oWordControl)
			this.m_oWordControl.m_oApi.checkLastWork();

		const context = canvas.getContext("2d");
		const canvasWidth = canvas.width;
		const canvasHeight = canvas.height;

		this.drawThumbnailsBorders(context, canvasWidth, canvasHeight);

		if (this.MouseDownTrack.IsDragged()) {
			this.drawThumbnailsInsertionLine(context, canvasWidth, canvasHeight);
		}
	};
	COutlineThumbnailsManager.prototype.drawThumbnailsBorders = function (context, canvasWidth, canvasHeight) {
		// context.fillStyle = GlobalSkin.BackgroundColorThumbnails;
		// context.fillRect(0, 0, canvasWidth, canvasHeight);
		//
		// const _style_select = GlobalSkin.ThumbnailsPageOutlineActive;
		// const _style_focus = GlobalSkin.ThumbnailsPageOutlineHover;
		// const _style_select_focus = GlobalSkin.ThumbnailsPageOutlineActive;
		//
		// context.fillStyle = _style_select;
		//
		// const oLogicDocument = this.m_oWordControl && this.m_oWordControl.m_oLogicDocument;
		// if (!oLogicDocument) {
		// 	return;
		// }
		//
		// const arrSlides = oLogicDocument.GetAllSlides();
		//
		//
		// for (let i = 0; i < arrSlides.length; i++) {
		// 	const page = this.m_arrPages[i];
		// 	const oSlide = arrSlides[i];
		// 	const slideType = oSlide.deleteLock.Lock.Get_Type();
		// 	const bLocked = slideType !== AscCommon.c_oAscLockTypes.kLockTypeMine && slideType !== AscCommon.c_oAscLockTypes.kLockTypeNone;
		//
		// 	let color = null;
		// 	let isFocus = false;
		//
		// 	if (bLocked) {
		// 		color = AscCommon.GlobalSkin.ThumbnailsLockColor;
		// 	} else if (page.IsSelected && page.IsFocused) {
		// 		color = _style_select_focus;
		// 	} else if (page.IsSelected) {
		// 		color = _style_select;
		// 	} else if (page.IsFocused) {
		// 		color = _style_focus;
		// 		isFocus = true;
		// 	}
		//
		// 	if (color) {
		// 		this.FocusRectFlat(color, context, page.left, page.top, page.right, page.bottom, isFocus);
		// 	} else {
		// 		this.OutlineRect(color, context, page.left, page.top, page.right, page.bottom);
		// 	}
		// }
	};
	COutlineThumbnailsManager.prototype.drawThumbnailsInsertionLine = function (context, canvasWidth, canvasHeight) {
		const isHorizontalThumbnails = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;
		const isRightToLeft = Asc.editor.isRtlInterface;
		const nPosition = this.MouseDownTrack.GetPosition();
		const oPage = this.m_arrPages[nPosition - 1];

		context.strokeStyle = "#DEDEDE";
		if (isHorizontalThumbnails) {
			const containerRightBorder = editor.WordControl.m_oThumbnailsContainer.AbsolutePosition.R * AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_mm_to_pix;

			// x, topY & bottomY - coordinates of the insertion line itself
			const x = oPage
				? isRightToLeft
					? (oPage.left - 1.5 * this.const_border_w) >> 0
					: (oPage.right + 1.5 * this.const_border_w) >> 0
				: isRightToLeft
					? (containerRightBorder - this.const_offset_x / 2) >> 0
					: (this.const_offset_x / 2) >> 0;

			let topY, bottomY;
			if (this.m_arrPages.length > 0) {
				topY = this.m_arrPages[0].top + 4;
				bottomY = this.m_arrPages[0].bottom - 4;
			} else {
				topY = 0;
				bottomY = canvasHeight;
			}

			context.beginPath();
			context.moveTo(x + 0.5, topY);
			context.lineTo(x + 0.5, bottomY);
			context.closePath();
			context.lineWidth = 3;
			context.stroke();

			if (this.MouseTrackCommonImage) {
				const commonImageWidth = this.MouseTrackCommonImage.width;
				const commonImageHeight = this.MouseTrackCommonImage.height;

				context.save();
				context.translate(x, topY);
				context.rotate(Math.PI / 2);

				context.drawImage(
					this.MouseTrackCommonImage,
					0, 0,
					commonImageWidth / 2, commonImageHeight,
					-commonImageWidth, -commonImageHeight / 2,
					commonImageWidth / 2, commonImageHeight
				);

				context.translate(bottomY - topY, 0);
				context.drawImage(
					this.MouseTrackCommonImage,
					commonImageWidth / 2, 0,
					commonImageWidth / 2, commonImageHeight,
					commonImageWidth / 2, -commonImageHeight / 2,
					commonImageWidth / 2, commonImageHeight
				);

				context.restore();
			}
		} else {
			const y = oPage
				? (oPage.bottom + 1.5 * this.const_border_w) >> 0
				: this.const_offset_y / 2 >> 0;

			let _left_pos = 0;
			let _right_pos = canvasWidth;
			if (this.m_arrPages.length > 0) {
				_left_pos = this.m_arrPages[0].left + 4;
				_right_pos = this.m_arrPages[0].right - 4;
			}

			context.lineWidth = 3;
			context.beginPath();
			context.moveTo(_left_pos, y + 0.5);
			context.lineTo(_right_pos, y + 0.5);
			context.stroke();
			context.beginPath();

			if (null != this.MouseTrackCommonImage) {
				context.drawImage(
					this.MouseTrackCommonImage,
					0, 0,
					(this.MouseTrackCommonImage.width + 1) >> 1, this.MouseTrackCommonImage.height,
					_left_pos - this.MouseTrackCommonImage.width, y - (this.MouseTrackCommonImage.height >> 1),
					(this.MouseTrackCommonImage.width + 1) >> 1, this.MouseTrackCommonImage.height
				);

				context.drawImage(
					this.MouseTrackCommonImage,
					this.MouseTrackCommonImage.width >> 1, 0,
					(this.MouseTrackCommonImage.width + 1) >> 1, this.MouseTrackCommonImage.height,
					_right_pos + (this.MouseTrackCommonImage.width >> 1), y - (this.MouseTrackCommonImage.height >> 1),
					(this.MouseTrackCommonImage.width + 1) >> 1, this.MouseTrackCommonImage.height
				);
			}
		}
	};

	// select
	COutlineThumbnailsManager.prototype.SelectSlides = function(aSelectedIdx, bReset)
	{
		if (aSelectedIdx.length > 0)
		{
			if(bReset === undefined || bReset)
			{
				for (let i = 0; i < this.m_arrPages.length; i++)
				{
					this.m_arrPages[i].IsSelected = false;
				}
			}
			let nCount = aSelectedIdx.length;
			let nSlideType = this.GetSlideType(aSelectedIdx[0]);
			for (let nIdx = 0; nIdx < nCount; nIdx++)
			{
				if(this.GetSlideType(aSelectedIdx[nIdx]) === nSlideType)
				{
					let oPage = this.m_arrPages[aSelectedIdx[nIdx]];
					if (oPage)
					{
						oPage.IsSelected = true;
					}
				}
			}
			this.OnUpdateOverlay();
		}
	};

	COutlineThumbnailsManager.prototype.GetSelectedSlidesRange = function()
	{
		var _min = this.m_oWordControl.m_oDrawingDocument.SlideCurrent;
		var _max = _min;
		var pages_count = this.m_arrPages.length;
		for (var i = 0; i < pages_count; i++)
		{
			if (this.m_arrPages[i].IsSelected)
			{
				if (i < _min)
					_min = i;
				if (i > _max)
					_max = i;
			}
		}
		return {Min: _min, Max: _max};
	};

	COutlineThumbnailsManager.prototype.GetSelectedArray = function()
	{
		var _array = [];
		var pages_count = this.m_arrPages.length;
		for (var i = 0; i < pages_count; i++)
		{
			if (this.m_arrPages[i].IsSelected)
			{
				_array[_array.length] = i;
			}
		}
		return _array;
	};

	COutlineThumbnailsManager.prototype.CorrectShiftSelect = function(isTop, isEnd)
	{
		var drDoc = this.m_oWordControl.m_oDrawingDocument;
		var slidesCount = drDoc.GetSlidesCount();
		var min_max = this.GetSelectedSlidesRange();

		var _page = this.m_oWordControl.m_oDrawingDocument.SlideCurrent;
		if (isEnd)
		{
			_page = isTop ? 0 : slidesCount - 1;
		} else if (isTop)
		{
			if (min_max.Min != _page)
			{
				_page = min_max.Min - 1;
				if (_page < 0)
					_page = 0;
			} else
			{
				_page = min_max.Max - 1;
				if (_page < 0)
					_page = 0;
			}
		} else
		{
			if (min_max.Min != _page)
			{
				_page = min_max.Min + 1;
				if (_page >= slidesCount)
					_page = slidesCount - 1;
			} else
			{
				_page = min_max.Max + 1;
				if (_page >= slidesCount)
					_page = slidesCount - 1;
			}
		}

		var _max = _page;
		var _min = this.m_oWordControl.m_oDrawingDocument.SlideCurrent;
		if (_min > _max)
		{
			var _temp = _max;
			_max = _min;
			_min = _temp;
		}

		for (var i = 0; i < _min; i++)
		{
			this.m_arrPages[i].IsSelected = false;
		}
		let nSlideType = this.GetSlideType(_min);
		for (var i = _min; i <= _max; i++)
		{
			if(this.GetSlideType(i) === nSlideType)
			{
				this.m_arrPages[i].IsSelected = true;
			}
		}
		for (var i = _max + 1; i < slidesCount; i++)
		{
			this.m_arrPages[i].IsSelected = false;
		}


		this.m_oWordControl.m_oLogicDocument.Document_UpdateInterfaceState();
		this.OnUpdateOverlay();
		this.ShowPage(_page);
	};

	COutlineThumbnailsManager.prototype.SelectAll = function()
	{
		for (var i = 0; i < this.m_arrPages.length; i++)
		{
			this.m_arrPages[i].IsSelected = true;
		}
		this.UpdateInterface();
		this.OnUpdateOverlay();
	};

	COutlineThumbnailsManager.prototype.IsSlideHidden = function(aSelected)
	{
		var oPresentation = this.m_oWordControl.m_oLogicDocument;
		for (var i = 0; i < aSelected.length; ++i)
		{
			if (oPresentation.IsVisibleSlide(aSelected[i]))
				return false;
		}
		return true;
	};

	COutlineThumbnailsManager.prototype.isSelectedPage = function(pageNum)
	{
		if (this.m_arrPages[pageNum] && this.m_arrPages[pageNum].IsSelected)
			return true;
		return false;
	};

	COutlineThumbnailsManager.prototype.SelectPage = function (pageNum) {
		if (!this.SelectPageEnabled)
			return;

		const pagesCount = this.m_arrPages.length;
		if (pageNum < 0 || pageNum >= pagesCount)
			return;

		let bIsUpdate = false;

		for (let i = 0; i < pagesCount; i++) {
			if (this.m_arrPages[i].IsSelected === true && i != pageNum) {
				bIsUpdate = true;
			}
			this.m_arrPages[i].IsSelected = false;
		}

		if (this.m_arrPages[pageNum].IsSelected === false) {
			bIsUpdate = true;
		}
		this.m_arrPages[pageNum].IsSelected = true;
		this.m_bIsUpdate = bIsUpdate;

		if (bIsUpdate) {
			this.ShowPage(pageNum);
		}
	};

	// position
	COutlineThumbnailsManager.prototype.ConvertCoords = function (x, y) {
		let posX, posY;
		switch (Asc.editor.getThumbnailsPosition()) {
			case thumbnailsPositionMap.left:
				posX = x - this.m_oWordControl.X;
				posY = y - this.m_oWordControl.Y;
				break;

			case thumbnailsPositionMap.right:
				posX = x - this.m_oWordControl.X - this.m_oWordControl.Width + this.m_oWordControl.splitters[0].position * g_dKoef_mm_to_pix;
				posY = y - this.m_oWordControl.Y;
				break;

			case thumbnailsPositionMap.bottom:
				posX = x - this.m_oWordControl.X;
				posY = y - this.m_oWordControl.Y - this.m_oWordControl.Height + this.m_oWordControl.splitters[0].position * g_dKoef_mm_to_pix;
				break;

			default:
				posX = posY = 0; // just in case
				break;
		}

		const convertedX = AscCommon.AscBrowser.convertToRetinaValue(posX, true);
		const convertedY = AscCommon.AscBrowser.convertToRetinaValue(posY, true);
		const pageIndex = this.m_arrPages.findIndex(function (page) {
			return page.Hit(convertedX, convertedY);
		});

		return {
			X: convertedX,
			Y: convertedY,
			Page: pageIndex
		};
	};
	COutlineThumbnailsManager.prototype.ConvertCoords2 = function (x, y) {
		if (this.m_arrPages.length == 0)
			return -1;

		let posX, posY;
		switch (Asc.editor.getThumbnailsPosition()) {
			case thumbnailsPositionMap.left:
				posX = x - this.m_oWordControl.X;
				posY = y - this.m_oWordControl.Y;
				break;

			case thumbnailsPositionMap.right:
				posX = x - this.m_oWordControl.X - this.m_oWordControl.Width + this.m_oWordControl.splitters[0].position * g_dKoef_mm_to_pix;
				posY = y - this.m_oWordControl.Y;
				break;

			case thumbnailsPositionMap.bottom:
				posX = x - this.m_oWordControl.X;
				posY = y - this.m_oWordControl.Y - this.m_oWordControl.Height + this.m_oWordControl.splitters[0].position * g_dKoef_mm_to_pix;
				break;

			default:
				posX = posY = 0; // just in case
				break;
		}

		const convertedX = AscCommon.AscBrowser.convertToRetinaValue(posX, true);
		const convertedY = AscCommon.AscBrowser.convertToRetinaValue(posY, true);

		const absolutePosition = this.m_oWordControl.m_oThumbnails.AbsolutePosition;
		const thControlWidth = (absolutePosition.R - absolutePosition.L) * g_dKoef_mm_to_pix * AscCommon.AscBrowser.retinaPixelRatio;
		const thControlHeight = (absolutePosition.B - absolutePosition.T) * g_dKoef_mm_to_pix * AscCommon.AscBrowser.retinaPixelRatio;

		if (convertedX < 0 || convertedX > thControlWidth || convertedY < 0 || convertedY > thControlHeight)
			return -1;

		const isHorizontalThumbnails = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;
		const isRightToLeft = Asc.editor.isRtlInterface;

		let minDistance = Infinity;
		let minPositionPage = 0;

		for (let i = 0; i < this.m_arrPages.length; i++) {
			const page = this.m_arrPages[i];

			let distanceToStart, distanceToEnd;
			if (isHorizontalThumbnails) {
				if (isRightToLeft) {
					distanceToStart = Math.abs(convertedX - page.right);
					distanceToEnd = Math.abs(convertedX - page.left);
				} else {
					distanceToStart = Math.abs(convertedX - page.left);
					distanceToEnd = Math.abs(convertedX - page.right);
				}
			} else {
				distanceToStart = Math.abs(convertedY - page.top);
				distanceToEnd = Math.abs(convertedY - page.bottom);
			}

			if (distanceToStart < minDistance) {
				minDistance = distanceToStart;
				minPositionPage = i;
			}
			if (distanceToEnd < minDistance) {
				minDistance = distanceToEnd;
				minPositionPage = i + 1;
			}
		}

		return minPositionPage;
	};
	COutlineThumbnailsManager.prototype.CalculatePlaces = function () {
		if (!this.isThumbnailsShown())
			return;

		if (this.m_oWordControl && this.m_oWordControl.MobileTouchManagerThumbnails)
			this.m_oWordControl.MobileTouchManagerThumbnails.ClearContextMenu();

		const thumbnails = this.m_oWordControl.m_oThumbnails;
		const thumbnailsCanvas = thumbnails.HtmlElement;
		if (!thumbnailsCanvas)
			return;

		const canvasWidth = thumbnailsCanvas.width;
		const canvasHeight = thumbnailsCanvas.height;

		const thumbnailsWidthMm = thumbnails.AbsolutePosition.R - thumbnails.AbsolutePosition.L;
		const thumbnailsHeightMm = thumbnails.AbsolutePosition.B - thumbnails.AbsolutePosition.T;

		const pixelRatio = AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_mm_to_pix;
		const isVerticalThumbnails = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.right
			|| Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.left;
		const isRightToLeft = Asc.editor.isRtlInterface;

		let thSlideWidthPx, thSlideHeightPx;
		let startOffset, supplement;
		if (isVerticalThumbnails) {
			thSlideWidthPx = (thumbnailsWidthMm * pixelRatio >> 0) - this.const_offset_x - this.const_offset_r;
			thSlideHeightPx = (thSlideWidthPx * this.SlideHeight / this.SlideWidth) >> 0;
			startOffset = this.const_offset_y;
		} else {
			thSlideHeightPx = (thumbnailsHeightMm * pixelRatio >> 0) - this.const_offset_y - this.const_offset_b;
			thSlideWidthPx = (thSlideHeightPx * this.SlideWidth / this.SlideHeight) >> 0;
			startOffset = this.const_offset_x;
		}

		if (this.m_bIsScrollVisible) {
			const scrollApi = editor.WordControl.m_oScrollThumbApi;
			if (scrollApi) {
				this.m_dScrollY_max = isVerticalThumbnails ? scrollApi.getMaxScrolledY() : scrollApi.getMaxScrolledX();
				this.m_dScrollY = isVerticalThumbnails ? scrollApi.getCurScrolledY() : scrollApi.getCurScrolledX();
			}
		}

		const currentScrollPx = isRightToLeft && !isVerticalThumbnails
			? this.m_dScrollY_max - this.m_dScrollY >> 0
			: this.m_dScrollY >> 0;

		let isFirstSlideFound = false;
		let isLastSlideFound = false;

		const totalSlidesCount = this.GetSlidesCount();
		for (let slideIndex = 0; slideIndex < totalSlidesCount; slideIndex++) {
			if (this.m_oWordControl.m_oLogicDocument.IsVisioEditor()) {
				const sizes = this.m_oWordControl.m_oLogicDocument.GetSizesMM(slideIndex);
				let visioSlideWidthMm = sizes.width;
				let visioSlideHeightMm = sizes.height;
				if (isVerticalThumbnails) {
					thSlideHeightPx = (thSlideWidthPx * visioSlideHeightMm / visioSlideWidthMm) >> 0;
				} else {
					thSlideWidthPx = (thSlideHeightPx * visioSlideWidthMm / visioSlideHeightMm) >> 0;
				}
			}

			if (slideIndex >= this.m_arrPages.length) {
				this.m_arrPages[slideIndex] = new CThPage();
				if (slideIndex === 0)
					this.m_arrPages[0].IsSelected = true;
			}

			const slideData = this.m_oWordControl.m_oLogicDocument.GetSlide(slideIndex);
			const slideRect = this.m_arrPages[slideIndex];
			slideRect.pageIndex = slideIndex;

			if (isVerticalThumbnails) {
				slideRect.left = isRightToLeft ? this.const_offset_r : this.const_offset_x;
				slideRect.top = startOffset - currentScrollPx;
				slideRect.right = slideRect.left + thSlideWidthPx;
				slideRect.bottom = slideRect.top + thSlideHeightPx;

				if (!isFirstSlideFound) {
					if ((startOffset + thSlideHeightPx) > currentScrollPx) {
						this.m_lDrawingFirst = slideIndex;
						isFirstSlideFound = true;
					}
				}

				if (!isLastSlideFound) {
					if (slideRect.top > canvasHeight) {
						this.m_lDrawingEnd = slideIndex - 1;
						isLastSlideFound = true;
					}
				}

				supplement = (thSlideHeightPx + 3 * this.const_border_w);
			} else {
				slideRect.top = this.const_offset_y;
				slideRect.bottom = slideRect.top + thSlideHeightPx;

				if (isRightToLeft) {
					slideRect.right = canvasWidth - startOffset + currentScrollPx;
					slideRect.left = slideRect.right - thSlideWidthPx;
				} else {
					slideRect.left = startOffset - currentScrollPx;
					slideRect.right = slideRect.left + thSlideWidthPx;
				}

				if (!isFirstSlideFound) {
					if ((startOffset + thSlideWidthPx) > currentScrollPx) {
						this.m_lDrawingFirst = slideIndex;
						isFirstSlideFound = true;
					}
				}

				if (!isLastSlideFound) {
					const isHidden = isRightToLeft
						? slideRect.right < 0
						: slideRect.left > canvasWidth;
					if (isHidden) {
						this.m_lDrawingEnd = slideIndex - 1;
						isLastSlideFound = true;
					}
				}

				supplement = (thSlideWidthPx + 3 * this.const_border_w);
			}

			if (slideData.getObjectType() === AscDFH.historyitem_type_SlideLayout) {
				const scaledWidth = (slideRect.right - slideRect.left) * AscCommonSlide.SlideLayout.LAYOUT_THUMBNAIL_SCALE;
				const scaledHeight = (slideRect.bottom - slideRect.top) * AscCommonSlide.SlideLayout.LAYOUT_THUMBNAIL_SCALE;

				if (isVerticalThumbnails) {
					slideRect.bottom = (slideRect.top + scaledHeight) + 0.5 >> 0;
					isRightToLeft
						? slideRect.right = (slideRect.left + scaledWidth) + 0.5 >> 0
						: slideRect.left = (slideRect.right - scaledWidth) + 0.5 >> 0;
					supplement = scaledHeight + 3 * this.const_border_w;
				} else {
					slideRect.top = (slideRect.bottom - scaledHeight) + 0.5 >> 0;
					isRightToLeft
						? slideRect.left = (slideRect.right - scaledWidth) + 0.5 >> 0
						: slideRect.right = (slideRect.left + scaledWidth) + 0.5 >> 0;
					supplement = scaledWidth + 3 * this.const_border_w;
				}
			}

			startOffset += supplement >> 0;
		}

		if (this.m_arrPages.length > totalSlidesCount)
			this.m_arrPages.splice(totalSlidesCount, this.m_arrPages.length - totalSlidesCount);

		if (!isLastSlideFound) {
			this.m_lDrawingEnd = totalSlidesCount - 1;
		}
	};

	COutlineThumbnailsManager.prototype.GetThumbnailPagePosition = function (pageIndex) {
		if (pageIndex < 0 || pageIndex >= this.m_arrPages.length)
			return null;

		const drawRect = this.m_arrPages[pageIndex];
		return {
			X: AscCommon.AscBrowser.convertToRetinaValue(drawRect.left),
			Y: AscCommon.AscBrowser.convertToRetinaValue(drawRect.top),
			W: AscCommon.AscBrowser.convertToRetinaValue(drawRect.right - drawRect.left),
			H: AscCommon.AscBrowser.convertToRetinaValue(drawRect.bottom - drawRect.top)
		};
	};
	COutlineThumbnailsManager.prototype.getSpecialPasteButtonCoords = function (sSlideId) {
		if (!sSlideId) return null;

		const oPresentation = this.m_oWordControl.m_oLogicDocument;
		const aSlides = oPresentation.GetAllSlides();

		const nSlideIdx = aSlides.findIndex(function (slide) {
			return slide.Get_Id() === sSlideId;
		});
		if (nSlideIdx === -1) return null;

		const oRect = this.GetThumbnailPagePosition(nSlideIdx);
		if (!oRect) return null;

		const oThContainer = editor.WordControl.m_oThumbnailsContainer;
		const isHidden = oThContainer.HtmlElement.style.display === "none";
		const isHorizontalThumbnails = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;

		const offsetX = oThContainer.AbsolutePosition.L * g_dKoef_mm_to_pix;
		const offsetY = oThContainer.AbsolutePosition.T * g_dKoef_mm_to_pix;

		const coordX = isHidden
			? 0
			: oRect.X + oRect.W - AscCommon.specialPasteElemWidth;
		const coordY = isHorizontalThumbnails
			? oRect.Y - AscCommon.specialPasteElemHeight
			: oRect.Y + oRect.H;

		return {
			X: offsetX + coordX,
			Y: offsetY + coordY
		};
	};

	// scroll
	COutlineThumbnailsManager.prototype.CheckNeedAnimateScrolls = function (type) {
		switch (type) {
			case -1:
				if (this.MouseThumbnailsAnimateScrollTopTimer != -1) {
					clearInterval(this.MouseThumbnailsAnimateScrollTopTimer);
					this.MouseThumbnailsAnimateScrollTopTimer = -1;
				}
				if (this.MouseThumbnailsAnimateScrollBottomTimer != -1) {
					clearInterval(this.MouseThumbnailsAnimateScrollBottomTimer);
					this.MouseThumbnailsAnimateScrollBottomTimer = -1;
				}
				break;

			case 0:
				if (this.MouseThumbnailsAnimateScrollBottomTimer != -1) {
					clearInterval(this.MouseThumbnailsAnimateScrollBottomTimer);
					this.MouseThumbnailsAnimateScrollBottomTimer = -1;
				}
				if (this.MouseThumbnailsAnimateScrollTopTimer == -1) {
					this.MouseThumbnailsAnimateScrollTopTimer = setInterval(this.OnScrollTrackTop, 50);
				}
				break;

			case 1:
				if (this.MouseThumbnailsAnimateScrollTopTimer != -1) {
					clearInterval(this.MouseThumbnailsAnimateScrollTopTimer);
					this.MouseThumbnailsAnimateScrollTopTimer = -1;
				}
				if (-1 == this.MouseThumbnailsAnimateScrollBottomTimer) {
					this.MouseThumbnailsAnimateScrollBottomTimer = setInterval(this.OnScrollTrackBottom, 50);
				}
				break;
		}
	};
	COutlineThumbnailsManager.prototype.OnScrollTrackTop = function () {
		Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom
			? this.m_oWordControl.m_oScrollThumbApi.scrollByX(-45)
			: this.m_oWordControl.m_oScrollThumbApi.scrollByY(-45);
	};
	COutlineThumbnailsManager.prototype.OnScrollTrackBottom = function () {
		Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom
			? this.m_oWordControl.m_oScrollThumbApi.scrollByX(45)
			: this.m_oWordControl.m_oScrollThumbApi.scrollByY(45);
	};
	COutlineThumbnailsManager.prototype.thumbnailsScroll = function (sender, scrollPosition, maxScrollPosition, isAtTop, isAtBottom) {
		if (false === this.m_oWordControl.m_oApi.bInit_word_control || false === this.m_bIsScrollVisible)
			return;

		this.m_dScrollY = Math.max(0, Math.min(scrollPosition, maxScrollPosition));
		this.m_dScrollY_max = maxScrollPosition;

		this.CalculatePlaces();
		AscCommon.g_specialPasteHelper.SpecialPasteButton_Update_Position();
		this.m_bIsUpdate = true;

		if (!this.m_oWordControl.m_oApi.isMobileVersion)
			this.SetFocusElement(FOCUS_OBJECT_THUMBNAILS);
	};

	// calculate
	COutlineThumbnailsManager.prototype.RecalculateAll = function()
	{
		const sizes = this.m_oWordControl.m_oLogicDocument.GetSizesMM();
		this.SlideWidth = sizes.width;
		this.SlideHeight = sizes.height;
		this.CheckSizes();

		this.ClearCacheAttack();
	};

	// size
	COutlineThumbnailsManager.prototype.CheckSizes = function () {
		this.InitCheckOnResize();

		if (!this.isThumbnailsShown())
			return;

		this.calculateThumbnailsOffsets();

		const wordControl = this.m_oWordControl;
		const dKoefToPix = AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_mm_to_pix;
		const isHorizontalOrientation = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;
		const dPosition = this.m_dScrollY_max != 0 ? this.m_dScrollY / this.m_dScrollY_max : 0;

		// Width and Height of 'id_panel_thumbnails' in millimeters
		const thumbnailsContainerWidth = wordControl.m_oThumbnailsContainer.AbsolutePosition.R - wordControl.m_oThumbnailsContainer.AbsolutePosition.L;
		const thumbnailsContainerHeight = wordControl.m_oThumbnailsContainer.AbsolutePosition.B - wordControl.m_oThumbnailsContainer.AbsolutePosition.T;

		let totalThumbnailsLength = this.calculateTotalThumbnailsLength(thumbnailsContainerWidth, thumbnailsContainerHeight);
		const thumbnailsContainerLength = isHorizontalOrientation
			? thumbnailsContainerWidth * dKoefToPix >> 0
			: thumbnailsContainerHeight * dKoefToPix >> 0;

		if (totalThumbnailsLength < thumbnailsContainerLength) {
			// All thumbnails fit inside - no scroller needed
			if (this.m_bIsScrollVisible && GlobalSkin.ThumbnailScrollWidthNullIfNoScrolling) {
				wordControl.m_oThumbnails.Bounds.R = 0;
				wordControl.m_oThumbnailsBack.Bounds.R = 0;
				wordControl.m_oThumbnails_scroll.Bounds.AbsW = 0;
				wordControl.m_oThumbnailsContainer.Resize(thumbnailsContainerWidth, thumbnailsContainerHeight, wordControl);
			} else {
				wordControl.m_oThumbnails_scroll.HtmlElement.style.display = 'none';
			}
			this.m_bIsScrollVisible = false;
			this.m_dScrollY = isHorizontalOrientation && Asc.editor.isRtlInterface
				? this.m_dScrollY_max
				: 0;

		} else {
			// Scrollbar is needed
			if (!this.m_bIsScrollVisible) {
				if (GlobalSkin.ThumbnailScrollWidthNullIfNoScrolling) {
					wordControl.m_oThumbnailsBack.Bounds.R = wordControl.ScrollWidthPx * dKoefToPix;
					wordControl.m_oThumbnails.Bounds.R = wordControl.ScrollWidthPx * dKoefToPix;
					const _width_mm_scroll = 10;
					wordControl.m_oThumbnails_scroll.Bounds.AbsW = _width_mm_scroll * dKoefToPix;
				} else if (!wordControl.m_oApi.isMobileVersion) {
					wordControl.m_oThumbnails_scroll.HtmlElement.style.display = 'block';
				}

				const thumbnailsContainerParent = wordControl.m_oThumbnailsContainer.Parent; // editor.WordControl.m_oBody
				wordControl.m_oThumbnailsContainer.Resize(
					thumbnailsContainerParent.AbsolutePosition.R - thumbnailsContainerParent.AbsolutePosition.L,
					thumbnailsContainerParent.AbsolutePosition.B - thumbnailsContainerParent.AbsolutePosition.T,
					wordControl
				);
			}
			this.m_bIsScrollVisible = true;

			const thumbnailsWidth = wordControl.m_oThumbnails.AbsolutePosition.R - wordControl.m_oThumbnails.AbsolutePosition.L;
			const thumbnailsHeight = wordControl.m_oThumbnails.AbsolutePosition.B - wordControl.m_oThumbnails.AbsolutePosition.T;
			totalThumbnailsLength = this.calculateTotalThumbnailsLength(thumbnailsWidth, thumbnailsHeight);

			const settings = this.getSettingsForScrollObject(totalThumbnailsLength);
			if (wordControl.m_oScrollThumb_ && wordControl.m_oScrollThumb_.isHorizontalScroll === isHorizontalOrientation) {
				wordControl.m_oScrollThumb_.Repos(settings);
			} else {
				const holder = wordControl.m_oThumbnails_scroll.HtmlElement;
				const canvas = holder.getElementsByTagName('canvas')[0];
				canvas && holder.removeChild(canvas);

				wordControl.m_oScrollThumb_ = new AscCommon.ScrollObject('id_vertical_scroll_thmbnl', settings);
				wordControl.m_oScrollThumbApi = wordControl.m_oScrollThumb_;

				const oThis = this;
				const eventName = isHorizontalOrientation ? 'scrollhorizontal' : 'scrollvertical';
				wordControl.m_oScrollThumb_.bind(eventName, function (evt) {
					const maxScrollPosition = isHorizontalOrientation ? evt.maxScrollX : evt.maxScrollY;
					oThis.thumbnailsScroll(oThis, evt.scrollD, maxScrollPosition);
				});
			}
			wordControl.m_oScrollThumb_.isHorizontalScroll = isHorizontalOrientation;

			if (Asc.editor.isRtlInterface && isHorizontalOrientation && this.m_dScrollY_max === 0) {
				wordControl.m_oScrollThumbApi.scrollToX(wordControl.m_oScrollThumbApi.maxScrollX);
			}
		}

		if (this.m_bIsScrollVisible) {
			const lPosition = (dPosition * wordControl.m_oScrollThumbApi.getMaxScrolledY()) >> 0;
			wordControl.m_oScrollThumbApi.scrollToY(lPosition);
		}

		isHorizontalOrientation
			? this.ScrollerWidth = totalThumbnailsLength
			: this.ScrollerHeight = totalThumbnailsLength;
		if (wordControl.MobileTouchManagerThumbnails) {
			wordControl.MobileTouchManagerThumbnails.Resize();
		}

		this.CalculatePlaces();
		AscCommon.g_specialPasteHelper.SpecialPasteButton_Update_Position();
		this.m_bIsUpdate = true;
	};
	COutlineThumbnailsManager.prototype.calculateThumbnailsOffsets = function () {
		const isHorizontalOrientation = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;
		if (isHorizontalOrientation) {
			this.const_offset_y = 25 + Math.round(9 * AscCommon.AscBrowser.retinaPixelRatio)
		} else {
			const _tmpDig = this.DigitWidths.length > 5 ? this.DigitWidths[5] : 0;
			const dKoefToPix = AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_mm_to_pix;
			const slidesCount = this.GetSlidesCount();
			const firstSlideNumber = this.m_oWordControl.m_oLogicDocument.getFirstSlideNumber();
			const totalSlidesLength = String(slidesCount + firstSlideNumber).length;
			this.const_offset_x = Math.max((_tmpDig * dKoefToPix * totalSlidesLength >> 0), 25) + Math.round(9 * AscCommon.AscBrowser.retinaPixelRatio);
		}
	};
	COutlineThumbnailsManager.prototype.calculateTotalThumbnailsLength = function (containerWidth, containerHeight) {
		const dKoefToPix = AscCommon.AscBrowser.retinaPixelRatio * g_dKoef_mm_to_pix;
		const oPresentation = this.m_oWordControl.m_oLogicDocument;
		const slidesCount = this.GetSlidesCount();

		const isHorizontalOrientation = Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom;

		let thumbnailHeight, thumbnailWidth;
		if (isHorizontalOrientation) {
			thumbnailHeight = (containerHeight * dKoefToPix >> 0) - (this.const_offset_y + this.const_offset_b);
			thumbnailWidth = thumbnailHeight * this.SlideWidth / this.SlideHeight >> 0;
		} else {
			thumbnailWidth = (containerWidth * dKoefToPix >> 0) - (this.const_offset_x + this.const_offset_r);
			thumbnailHeight = thumbnailWidth * this.SlideHeight / this.SlideWidth >> 0;
		}

		const cumulativeThumbnailLength = oPresentation.getCumulativeThumbnailsLength(isHorizontalOrientation, thumbnailWidth, thumbnailHeight);

		const totalThumbnailsLength = isHorizontalOrientation
			? 2 * this.const_offset_x + cumulativeThumbnailLength + (slidesCount > 0 ? (slidesCount - 1) * 3 * this.const_border_w : 0)
			: 2 * this.const_offset_y + cumulativeThumbnailLength + (slidesCount > 0 ? (slidesCount - 1) * 3 * this.const_border_w : 0);

		return totalThumbnailsLength;
	};
	COutlineThumbnailsManager.prototype.getSettingsForScrollObject = function (totalContentLength) {
		const settings = new AscCommon.ScrollSettings();
		settings.screenW = this.m_oWordControl.m_oThumbnails.HtmlElement.width;
		settings.screenH = this.m_oWordControl.m_oThumbnails.HtmlElement.height;
		settings.showArrows = false;
		settings.slimScroll = true;
		settings.cornerRadius = 1;

		settings.scrollBackgroundColor = GlobalSkin.BackgroundColorThumbnails;
		settings.scrollBackgroundColorHover = GlobalSkin.BackgroundColorThumbnails;
		settings.scrollBackgroundColorActive = GlobalSkin.BackgroundColorThumbnails;

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

		if (Asc.editor.getThumbnailsPosition() === thumbnailsPositionMap.bottom) {
			settings.isHorizontalScroll = true;
			settings.isVerticalScroll = false;
			settings.contentW = totalContentLength;
		} else {
			settings.isHorizontalScroll = false;
			settings.isVerticalScroll = true;
			settings.contentH = totalContentLength;
		}

		return settings;
	};

	// interface
	COutlineThumbnailsManager.prototype.UpdateInterface = function()
	{
		this.m_oWordControl.m_oLogicDocument.Document_UpdateInterfaceState();
	};

	COutlineThumbnailsManager.prototype.showContextMenu = function(bPosBySelect)
	{
		let sSelectedIdx = this.GetSelectedArray();
		let oMenuPos;
		let bIsSlideSelect = false;
		if(bPosBySelect)
		{
			oMenuPos = this.GetThumbnailPagePosition(Math.min.apply(Math, sSelectedIdx));
			bIsSlideSelect = sSelectedIdx.length > 0;
		}
		else
		{
			let oEditorCtrl = this.m_oWordControl;
			let oThCtrlPos = oEditorCtrl.m_oThumbnails.AbsolutePosition;
			oMenuPos = {
				X: global_mouseEvent.X - ((oThCtrlPos.L * g_dKoef_mm_to_pix) >> 0) - oEditorCtrl.X,
				Y: global_mouseEvent.Y - ((oThCtrlPos.T * g_dKoef_mm_to_pix) >> 0) - oEditorCtrl.Y
			};

			let pos = this.ConvertCoords(global_mouseEvent.X, global_mouseEvent.Y);
			bIsSlideSelect = pos.Page !== -1;
		}
		if (oMenuPos)
		{
			const oLogicDocument = this.m_oWordControl.m_oLogicDocument;
			let oFirstSlide = oLogicDocument.GetSlide(sSelectedIdx[0]);
			let nType = Asc.c_oAscContextMenuTypes.Thumbnails;
			if(oFirstSlide)
			{
				if(oFirstSlide.getObjectType() === AscDFH.historyitem_type_SlideLayout)
				{
					nType = Asc.c_oAscContextMenuTypes.Layout;
				}
				else if(oFirstSlide.getObjectType() === AscDFH.historyitem_type_SlideMaster)
				{
					nType = Asc.c_oAscContextMenuTypes.Master;
				}
			}
			let oData =
				{
					Type: nType,
					X_abs: oMenuPos.X,
					Y_abs: oMenuPos.Y,
					IsSlideSelect: bIsSlideSelect,
					IsSlideHidden: this.IsSlideHidden(sSelectedIdx),
					IsSlidePreserve: oLogicDocument.isPreserveSelectionSlides()
				};
			editor.sync_ContextMenuCallback(new AscCommonSlide.CContextMenuData(oData));
		}
	};

	COutlineThumbnailsManager.prototype.GetCurSld = function()
	{
		return this.thumbnails.GetCurSld();
	};




	window["AscCommon"] = window["AscCommon"] || {};
	window["AscCommon"].CThumbnailsManager = CThumbnailsManager;
	window["AscCommon"].COutlineThumbnailsManager = COutlineThumbnailsManager;
})();
