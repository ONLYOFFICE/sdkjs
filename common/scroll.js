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

"use strict";

(	/**
* @param {Window} window
* @param {undefined} undefined
*/
function (window, undefined) {

    var AscBrowser = window['AscCommon'].AscBrowser;

var debug = false;

var ArrowType = {
    NONE: 0,
    ARROW_TOP: 1,
    ARROW_RIGHT: 2,
    ARROW_BOTTOM: 3,
    ARROW_LEFT: 4
};

var AnimationType = {
    NONE: 0,
    SCROLL_HOVER: 1,
    ARROW_HOVER: 2,
    SCROLL_ACTIVE: 3,
    ARROW_ACTIVE: 4
};

function GetClientWidth( elem ) {
	if (!elem) return 0;
    var _w = elem.clientWidth;
    if ( 0 != _w )
        return _w;

    var _string_w = "" + elem.style.width;
    if ( -1 < _string_w.indexOf( "%" ) )
        return 0;

    var _intVal = parseInt( _string_w );
    if ( !isNaN( _intVal ) && 0 < _intVal )
        return _intVal;

    return 0;
}
function GetClientHeight( elem ) {
	if (!elem) return 0;
    var _w = elem.clientHeight;
    if ( 0 != _w )
        return _w;

    var _string_w = "" + elem.style.height;
    if ( -1 < _string_w.indexOf( "%" ) )
        return 0;

    var _intVal = parseInt( _string_w );
    if ( !isNaN( _intVal ) && 0 < _intVal )
        return _intVal;

    return 0;
}

function CArrowDrawer( settings ) {
    // размер квадратика в пикселах
    this.Size = 16;
    this.SizeW = 16;
    this.SizeH = 16;

    this.SizeNaturalW = this.SizeW;
    this.SizeNaturalH = this.SizeH;

    this.IsRetina = false;

    this.ColorGradStart  = {R: _HEXTORGB_(settings.arrowColor).R, G: _HEXTORGB_(settings.arrowColor).G, B: _HEXTORGB_(settings.arrowColor).B};

    // вот такие мега настройки для кастомизации)
    this.IsDrawBorderInNoneMode = false;
    this.IsDrawBorders = true;

    //arrow pixel size
    this.pxCount = settings.slimScroll ? 4 : 6;

    // имя - направление стрелки
    this.ImageLeft = null;
    this.ImageTop = null;
    this.ImageRight = null;
    this.ImageBottom = null;

    this.fadeInFadeOutDelay = settings.fadeInFadeOutDelay || 30;
}
CArrowDrawer.prototype.checkSettings = function ( settings )
{
	this.ColorGradStart  = {R: _HEXTORGB_(settings.arrowColor).R, G: _HEXTORGB_(settings.arrowColor).G, B: _HEXTORGB_(settings.arrowColor).B};
}

CArrowDrawer.prototype.InitSize = function ( sizeW, sizeH, is_retina ) {
  /*  if ( ( sizeH == this.SizeH || sizeW == this.SizeW ) && is_retina == this.IsRetina && null != this.ImageLeft )
        return;*/
	var dPR = AscBrowser.retinaPixelRatio;
    this.SizeW = Math.max( sizeW, 1 );
    this.SizeH = Math.max( sizeH, 1 );
    this.IsRetina = is_retina;

    this.SizeNaturalW = this.SizeW;
    this.SizeNaturalH = this.SizeH;

    if (null == this.ImageLeft || null == this.ImageTop || null == this.ImageRight || null == this.ImageBottom)
	{
		this.ImageLeft = document.createElement('canvas');
		this.ImageTop = document.createElement('canvas');
		this.ImageRight = document.createElement('canvas');
		this.ImageBottom = document.createElement('canvas');
	}

	var len = Math.floor(this.pxCount * dPR);
	if ( this.SizeH < this.pxCount )
		return;

	// теперь делаем нечетную длину
	if ( 0 == (len & 1) )
		len += 1;

	var countPart = (len + 1) >> 1,
		_data, px,
		_x = ((this.SizeW - len) >> 1),
		_y = this.SizeH - ((this.SizeH - countPart) >> 1),
		r, g, b;
	var __x = _x, __y = _y, _len = len;
	r = this.ColorGradStart.R;
	g = this.ColorGradStart.G;
	b = this.ColorGradStart.B;
	this.ImageTop.width = this.SizeW;
	this.ImageTop.height = this.SizeH;
    this.ImageBottom.width = this.SizeW;
    this.ImageBottom.height = this.SizeH;
    this.ImageLeft.width = this.SizeW;
    this.ImageLeft.height = this.SizeH;
    this.ImageRight.width = this.SizeW;
    this.ImageRight.height = this.SizeH;
	var ctx = this.ImageTop.getContext('2d');
	var ctxBottom = this.ImageBottom.getContext('2d');
	var ctxLeft = this.ImageLeft.getContext('2d');
    var ctxRight = this.ImageRight.getContext('2d');

	_data = ctx.createImageData(this.SizeW, this.SizeH);
	px = _data.data;

	while (_len > 0) {
		var ind = 4 * (this.SizeW * __y + __x);
		for (var i = 0; i < _len; i++) {
			px[ind++] = r;
			px[ind++] = g;
			px[ind++] = b;
			px[ind++] = 255;
		}

		r = r >> 0;
		g = g >> 0;
		b = b >> 0;

		__x += 1;
		__y -= 1;
		_len -= 2;
	}
    var dy = dPR <= 1 ? -1 : 0;

	ctx.putImageData(_data, 0, dy);

	_data = ctxLeft.createImageData(this.SizeW, this.SizeH);
	px = _data.data;
	_len = len;
	__x = _x, __y = _y;
	while (_len > 0) {
		var ind = 4 * (this.SizeH * __x + __y);
		var xx = __x;
		for (var i = 0; i < _len; i++) {
			px[ind++] = r;
			px[ind++] = g;
			px[ind++] = b;
			px[ind++] = 255;
			++xx;
			ind = 4 * (this.SizeH * xx  + __y);
		}
		r = r >> 0;
		g = g >> 0;
		b = b >> 0;
		__x += 1;
		__y -= 1;
		_len -= 2;
	}
	var dx = dPR <= 1 ? -1 : 0;
	var dy = this.SizeW % 2 === 0 ? 1 : 0;
	ctxLeft.putImageData(_data, dx, dy);

	dx = this.SizeW % 2 === 0 ? 1 : 0;
	ctxBottom.translate( this.SizeW / 2, this.SizeH / 2);
	ctxBottom.rotate(Math.PI);
	ctxBottom.translate(-this.SizeW / 2, -this.SizeH / 2);
    ctxBottom.drawImage( this.ImageTop, dx, 0, this.ImageBottom.width, this.ImageBottom.height);

	dy = this.SizeH % 2 === 0 ? -1 : 0;
	ctxRight.translate( this.SizeW / 2, this.SizeH / 2);
	ctxRight.rotate(Math.PI);
	ctxRight.translate(-this.SizeW / 2, -this.SizeH / 2);
	ctxRight.drawImage( this.ImageLeft, 0, dy, this.ImageRight.width, this.ImageRight.height);
};

	/**
	 * @constructor
	 */
	function ScrollSettings() {
		this.showArrows = true;
		this.screenW = -1;
		this.screenH = -1;
		this.screenAddH = 0;

		this.contentH = 0;
		this.contentW = 0;

		//timeout when scroll move
		this.initialDelay = 300;
		this.arrowRepeatFreq = 50;
		this.trackClickRepeatFreq = 70;

		this.scrollPagePercent = 1. / 8;

		//max and min scroll size
		this.scrollerMinHeight = 34;
		this.scrollerMaxHeight = 99999;
		this.scrollerMinWidth = 34;
		this.scrollerMaxWidth = 99999;

		//arrow dimension
		this.arrowDim = Math.round(13 * AscBrowser.retinaPixelRatio);

		//scroll elements color
		this.scrollerColor = "#f1f1f1";
		this.scrollerHoverColor = "#cfcfcf";
		this.scrollerActiveColor = "#adadad";
		this.scrollBackgroundColor = "#f4f4f4";
		this.scrollBackgroundColorHover = "#f4f4f4";
		this.scrollBackgroundColorActive = "#f4f4f4";
		this.strokeStyleNone = "#cfcfcf";
		this.strokeStyleOver = "#cfcfcf";
		this.strokeStyleActive = "#ADADAD";

		//scroll speed
		this.vscrollStep = 10;
		this.hscrollStep = 10;
		this.wheelScrollLines = 1;

		//arrow elements color
		this.arrowColor = "#ADADAD";
		this.arrowHoverColor = "#f1f1f1";
		this.arrowActiveColor = "#f1f1f1";

		this.arrowBorderColor = "#cfcfcf";
		this.arrowOverBorderColor = "#cfcfcf";
		this.arrowOverBackgroundColor = "#cfcfcf";
		this.arrowActiveBorderColor = "#ADADAD";
		this.arrowActiveBackgroundColor = "#ADADAD";

		//scroll animation time delay
		this.fadeInFadeOutDelay = 20;

		//stripes color
		this.targetColor = "#cfcfcf";
		this.targetHoverColor = "#f1f1f1";
		this.targetActiveColor = "#f1f1f1";

		this.defaultColor = 241;
		this.hoverColor = 207;
		this.activeColor = 173;

        this.arrowSizeW = Math.round(13 * AscBrowser.retinaPixelRatio);
        this.arrowSizeH = Math.round(13 * AscBrowser.retinaPixelRatio);
		this.cornerRadius = 0;
		this.slimScroll = false;
		this.alwaysVisible = false;
        this.isVerticalScroll = true;
        this.isHorizontalScroll = false;
	}

	/**
	 * @constructor
	 */
	function ScrollObject( elemID, settings, dbg ) {
		if ( dbg )
			debug = dbg;
		this.that = this;

		this.settings = settings;
		this.ArrowDrawer = new CArrowDrawer( this.settings );

		this.mouseUp = false;
		this.mouseDown = false;

		this.that.mouseover = false;

		this.scrollerMouseDown = false;
        this.animState = AnimationType.NONE;
        this.lastAnimState = this.animState;
        this.arrowState = ArrowType.NONE;
        this.lastArrowState = this.arrowState;

		this.moveble = false;
		this.lock = false;
		this.scrollTimeout = null;

		this.StartMousePosition = {x:0, y:0};
		this.EndMousePosition = {x:0, y:0};

		this.dragMinY = 0;
		this.dragMaxY = 0;

		this.scrollVCurrentY = 0;
		this.scrollHCurrentX = 0;
		this.arrowPosition = 0;

		this.verticalTrackHeight = 0;
		this.horizontalTrackWidth = 0;

		this.paneHeight = 0;
		this.paneWidth = 0;

		this.maxScrollY = 0;
		this.maxScrollX = 0;
		this.maxScrollY2 = 0;
		this.maxScrollX2 = 0;

		this.scrollCoeff = 0;
		this.isResizeArrows = false;

		this.scroller = {x:0, y:1, h:0, w:0};

		this.canvas = null;
		this.context = null;

		this.eventListeners = [];

		this.IsRetina = false;
		this.canvasW = 1;
		this.canvasH = 1;
		this.canvasOriginalW = 1;
		this.canvasOriginalH = 1;

		this.scrollColor = _HEXTORGB_(this.settings.scrollerColor).R;
		this.arrowColor = _HEXTORGB_(this.settings.arrowColor).R;
		this.firstArrow = {arrowColor: _HEXTORGB_(this.settings.arrowColor).R, arrowBackColor: _HEXTORGB_(this.settings.scrollerColor).R, arrowStrokeColor: _HEXTORGB_(this.settings.strokeStyleNone).R};
		this.secondArrow = {arrowColor: _HEXTORGB_(this.settings.arrowColor).R, arrowBackColor: _HEXTORGB_(this.settings.scrollerColor).R, arrowStrokeColor: _HEXTORGB_(this.settings.strokeStyleNone).R};

		this.targetColor = _HEXTORGB_(this.settings.targetColor).R;
		this.strokeColor = _HEXTORGB_(this.settings.strokeStyleNone).R;

        this.fadeTimeoutScroll = null;
        this.fadeTimeoutArrows = null;

		this.IsRetina = AscBrowser.isRetina;

		this.disableCurrentScroll = false;
        this._initPiperImg();

		this._init( elemID );
	}

	ScrollObject.prototype._initPiperImg = function() {
		var dPR = AscBrowser.retinaPixelRatio;
		this.piperImgVert = document.createElement( 'canvas' );
		this.piperImgHor =  document.createElement( 'canvas' );

		this.piperImgVert.height = Math.round( 13 * dPR);
		this.piperImgVert.width = Math.round(5 * dPR);

		this.piperImgHor.width = Math.round( 13 * dPR);
		this.piperImgHor.height = Math.round(5 * dPR);

		if(this.settings.slimScroll){
			this.piperImgVert.width =
				this.piperImgHor.height = Math.round(3 * dPR);
		}

		var ctx_piperImg;
		ctx_piperImg = this.piperImgVert.getContext('2d');
		ctx_piperImg.beginPath();
		ctx_piperImg.lineWidth = Math.floor(dPR);
		var count = 0;

		var _dPR = dPR < 0.5 ? Math.ceil(dPR) :  Math.round(dPR);

		for (var i = 0; i < this.piperImgVert.height; i += _dPR + Math.floor(dPR)) {
			ctx_piperImg.moveTo(0, i + 0.5 * ctx_piperImg.lineWidth);
			ctx_piperImg.lineTo(this.piperImgVert.width, i + 0.5 * ctx_piperImg.lineWidth);
			ctx_piperImg.stroke();
			++count;
			if (count > 6 && dPR >= 1) break;
		}

		ctx_piperImg = this.piperImgHor.getContext('2d');
		ctx_piperImg.beginPath();
		ctx_piperImg.lineWidth = Math.floor(dPR);
		var count = 0;
		for (var i = 0; i < this.piperImgHor.width; i += _dPR + Math.floor(dPR)) {
			ctx_piperImg.moveTo(i + 0.5 * ctx_piperImg.lineWidth, 0);
			ctx_piperImg.lineTo(i + 0.5 * ctx_piperImg.lineWidth, this.piperImgHor.height);
			ctx_piperImg.stroke();
			++count;
			if (count > 6 && dPR >= 1) break;
		}
	}

	ScrollObject.prototype._init = function ( elemID ) {
		if ( !elemID ) return false;

		var dPR = AscBrowser.retinaPixelRatio;
		var holder = document.getElementById( elemID );

		if ( holder.getElementsByTagName( 'canvas' ).length == 0 ){
			this.canvas = holder.appendChild( document.createElement( "CANVAS" ) );
		}
		else {
			this.canvas = holder.children[1];
		}

		this.canvas.style.width = "100%";
		this.canvas.style.height = "100%";

		this.canvas.that = this;
		this.canvas.style.zIndex = 100;
		this.canvas.style.position = "absolute";
		this.canvas.style.top = "0px";
		this.canvas.style["msTouchAction"] = "none";
		if ( navigator.userAgent.toLowerCase().indexOf( "webkit" ) != -1 ){
			this.canvas.style.webkitUserSelect = "none";
		}

		this.context = this.canvas.getContext( '2d' );

		this._setDimension( holder.clientHeight, holder.clientWidth );

		this.maxScrollY = this.maxScrollY2 = this.settings.contentH - this.settings.screenH > 0 ? this.settings.contentH - this.settings.screenH : 0;
		this.maxScrollX = this.maxScrollX2 = this.settings.contentW - this.settings.screenW > 0 ? this.settings.contentW - this.settings.screenW : 0;

		if(this.settings.isVerticalScroll && !this.settings.alwaysVisible) {
			this.canvas.style.display = this.maxScrollY == 0 ? "none" : "";
		}

		if(this.settings.isHorizontalScroll && !this.settings.alwaysVisible) {
			this.canvas.style.display = this.maxScrollX == 0 ? "none" : "";
		}

		this._setScrollerHW();
		this.settings.arrowDim = this.settings.slimScroll && this.settings.isVerticalScroll ? this.scroller.w : this.settings.arrowDim;
		this.settings.arrowDim = this.settings.slimScroll && this.settings.isHorizontalScroll ? this.scroller.h : this.settings.arrowDim;
		this.arrowPosition = this.settings.showArrows ? Math.round(this.settings.arrowDim + this._roundForScale(dPR) + this._roundForScale(dPR)) : Math.round(dPR);

		this.paneHeight = this.canvasH - this.arrowPosition * 2;
		this.paneWidth = this.canvasW - this.arrowPosition * 2;

		this.RecalcScroller();

		AscCommon.addMouseEvent(this.canvas, "down", this.evt_mousedown);
        AscCommon.addMouseEvent(this.canvas, "move", this.evt_mousemove);
        AscCommon.addMouseEvent(this.canvas, "up", this.evt_mouseup);
        AscCommon.addMouseEvent(this.canvas, "over", this.evt_mouseover);
        AscCommon.addMouseEvent(this.canvas, "out", this.evt_mouseout);
        this.canvas.onmousewheel = this.evt_mousewheel;

		var _that = this;
		this.canvas.ontouchstart = function ( e ) {
			_that.evt_mousedown( e.touches[0] );
			return false;
		};
		this.canvas.ontouchmove = function ( e ) {
			_that.evt_mousemove( e.touches[0] );
			return false;
		};
		this.canvas.ontouchend = function ( e ) {
			_that.evt_mouseup( e.changedTouches[0] );
			return false;
		};

		if ( this.canvas.addEventListener ){
			this.canvas.addEventListener( 'DOMMouseScroll', this.evt_mousewheel, false );
		}

		this.context.fillStyle = this.settings.scrollBackgroundColor;
		this.context.fillRect(0,0,this.canvasW,this.canvasH);
        this._drawArrows();
		this._draw();

		return true;
	};
    ScrollObject.prototype.disableCurrentScroll = function() {
        this.disableCurrentScroll = true;
    };
	ScrollObject.prototype.checkDisableCurrentScroll = function() {
        var ret = this.disableCurrentScroll;
        this.disableCurrentScroll = false;
        return ret;
    };
	ScrollObject.prototype.getMousePosition = function ( evt ) {
		// get canvas position
		var obj = this.canvas;
		var top = 0;
		var left = 0;
		while ( obj && obj.tagName != 'BODY' ) {
			top += obj.offsetTop;
			left += obj.offsetLeft;
			obj = obj.offsetParent;
		}
		var dPR = AscBrowser.retinaPixelRatio;
		// return relative mouse position
        var mouseX = (((evt.clientX * AscBrowser.zoom) >> 0) - left + window.pageXOffset) * dPR;
        var mouseY = (((evt.clientY * AscBrowser.zoom) >> 0) - top + window.pageYOffset) * dPR;

		return {
			x:mouseX,
			y:mouseY
		};
	};
	ScrollObject.prototype.RecalcScroller = function ( startpos ) {
		var dPR = AscBrowser.retinaPixelRatio;
		if ( this.settings.isVerticalScroll ) {
			if ( this.settings.showArrows ) {
				this.verticalTrackHeight = this.canvasH - this.arrowPosition * 2;
				this.scroller.y = this.arrowPosition;
			}
			else {
				this.verticalTrackHeight = this.canvasH;
				this.scroller.y = Math.round(dPR);
			}
			var percentInViewV;

			percentInViewV = (this.maxScrollY * dPR + this.paneHeight) / this.paneHeight;
			this.scroller.h = Math.ceil(Math.ceil( 1 / percentInViewV * this.verticalTrackHeight / dPR) * dPR);

			var scrollMinH = Math.round(this.settings.scrollerMinHeight * dPR);
			
			if ( this.scroller.h < scrollMinH)
				this.scroller.h = scrollMinH;
			else if ( this.scroller.h > this.settings.scrollerMaxHeight )
				this.scroller.h = this.settings.scrollerMaxHeight;
			this.scrollCoeff = this.maxScrollY / Math.max( 1, this.paneHeight - this.scroller.h );
			if ( startpos ) {
				this.scroller.y = startpos / this.scrollCoeff + this.arrowPosition;
			}
			this.dragMaxY = this.canvasH - this.arrowPosition - this.scroller.h + 1;
			this.dragMinY = this.arrowPosition;
		}

		if ( this.settings.isHorizontalScroll ) {
			if ( this.settings.showArrows ) {
				this.horizontalTrackWidth = this.canvasW - this.arrowPosition * 2;
				this.scroller.x = this.arrowPosition;
			}
			else {
				this.horizontalTrackWidth = this.canvasW;
				this.scroller.x = Math.round(dPR);
			}
			var percentInViewH;
			percentInViewH = ( this.maxScrollX * dPR + this.paneWidth ) / this.paneWidth;
			this.scroller.w = Math.ceil(Math.ceil( 1 / percentInViewH * this.horizontalTrackWidth / dPR) * dPR);

			var scrollMinW = Math.round(this.settings.scrollerMinWidth * dPR);

			if ( this.scroller.w < scrollMinW)
				this.scroller.w = scrollMinW;
			else if ( this.scroller.w > this.settings.scrollerMaxWidth )
				this.scroller.w = this.settings.scrollerMaxWidth;
			this.scrollCoeff = this.maxScrollX / Math.max( 1, this.paneWidth - this.scroller.w );
			if ( typeof startpos !== "undefined" ) {
				this.scroller.x = startpos / this.scrollCoeff + this.arrowPosition;
			}
			this.dragMaxX = this.canvasW - this.arrowPosition - this.scroller.w;
			this.dragMinX = this.arrowPosition;
		}
	};
	ScrollObject.prototype.Repos = function ( settings, bIsHorAttack, bIsVerAttack, pos ) {
		var dPR = AscBrowser.retinaPixelRatio;

		var isChangeTheme = settings && this.settings.scrollBackgroundColor !== settings.scrollBackgroundColor;

		if (isChangeTheme)
		{
			for ( var i in settings )
				this.settings[i] = settings[i];
		}

		if (this.settings.showArrows) {
			this.settings.arrowSizeW = Math.round(13 * dPR);
			this.settings.arrowSizeH = Math.round(13 * dPR);
			this.ArrowDrawer.InitSize(this.settings.arrowSizeH, this.settings.arrowSizeW, this.IsRetina);
		}

		if (bIsVerAttack)
		{
			var _canvasH = settings.screenH;
			if (undefined !== _canvasH && settings.screenAddH)
				_canvasH += settings.screenAddH;

			if (_canvasH == this.canvasH && undefined !== settings.contentH)
			{
				var _maxScrollY = settings.contentH - settings.screenH > 0 ? settings.contentH - settings.screenH : 0;
				if (_maxScrollY == this.maxScrollY && !isChangeTheme)
					return;
			}
		}
		if (bIsHorAttack)
		{
			if (settings.screenW == this.canvasW && undefined !== settings.contentW)
			{
				var _maxScrollX = settings.contentW - settings.screenW > 0 ? settings.contentW - settings.screenW : 0;
				if (_maxScrollX == this.maxScrollX && !isChangeTheme)
					return;
			}
		}
		var dPR = AscBrowser.retinaPixelRatio;
		var _parentClientW = GetClientWidth( this.canvas.parentNode );
		var _parentClientH = GetClientHeight( this.canvas.parentNode );

		var _firstChildW = settings.contentW;
        var _firstChildH = settings.contentH;

		this.maxScrollY = this.maxScrollY2 = _firstChildH - settings.screenH > 0 ? _firstChildH - settings.screenH : 0;
		this.maxScrollX = this.maxScrollX2 = _firstChildW - settings.screenW > 0 ? _firstChildW - settings.screenW : 0;

		this._setDimension( _parentClientH, _parentClientW );
		this._setScrollerHW();

		this.settings.arrowDim = Math.round(13 * dPR);
		this.settings.arrowDim = this.settings.slimScroll && this.settings.isVerticalScroll ? this.scroller.w : this.settings.arrowDim;
		this.settings.arrowDim = this.settings.slimScroll && this.settings.isHorizontalScroll ? this.scroller.h : this.settings.arrowDim;
		this.arrowPosition = this.settings.showArrows ? Math.round(this.settings.arrowDim + this._roundForScale(dPR) + this._roundForScale(dPR)) : Math.round(dPR);
		this.paneHeight = this.canvasH - this.arrowPosition * 2;
		this.paneWidth = this.canvasW - this.arrowPosition * 2;
		this.RecalcScroller();
		if ( this.settings.isVerticalScroll && !this.settings.alwaysVisible) {

			if (this.scrollVCurrentY > this.maxScrollY)
				this.scrollVCurrentY = this.maxScrollY;

			this.scrollToY( this.scrollVCurrentY );
			if(this.maxScrollY == 0){
				this.canvas.style.display = "none";
			}
			else{
				this.canvas.style.display = "";
			}
		}
		else if ( this.settings.isHorizontalScroll ) {

			if (this.scrollHCurrentX > this.maxScrollX)
				this.scrollHCurrentX = this.maxScrollX;

			this.scrollToX( this.scrollHCurrentX );
			if(this.maxScrollX == 0 && !this.settings.alwaysVisible){
				this.canvas.style.display = "none";
			}
			else{
				this.canvas.style.display = "";
			}
		}

		this.reinit = true;
		if ( this.settings.isVerticalScroll && pos) {
			pos !== undefined ? this.scrollByY( pos - this.scrollVCurrentY ) : this.scrollToY( this.scrollVCurrentY );
		}

		if ( this.settings.isHorizontalScroll && pos) {
			pos !== undefined ? this.scrollByX( pos - this.scrollHCurrentX ) : this.scrollToX( this.scrollHCurrentX );
		}
		this.reinit = false;

		if (this.isResizeArrows || isChangeTheme) {
			this.context.fillStyle = this.settings.scrollBackgroundColor;
			this.context.fillRect(0,0,this.canvasW,this.canvasH);

		    this.firstArrow = {arrowColor: _HEXTORGB_(this.settings.arrowColor).R, arrowBackColor: _HEXTORGB_(this.settings.scrollerColor).R, arrowStrokeColor: _HEXTORGB_(this.settings.strokeStyleNone).R};
		    this.secondArrow = {arrowColor: _HEXTORGB_(this.settings.arrowColor).R, arrowBackColor: _HEXTORGB_(this.settings.scrollerColor).R, arrowStrokeColor: _HEXTORGB_(this.settings.strokeStyleNone).R};
			this._drawArrows();
		}
		this._initPiperImg();
		this._draw();
	};
	ScrollObject.prototype._scrollV = function ( that, evt, pos, isTop, isBottom, bIsAttack ) {
		if ( !this.settings.isVerticalScroll ) {
			return;
		}

		if ( that.scrollVCurrentY !== pos || bIsAttack === true ) {
            var oldPos = that.scrollVCurrentY;
		    that.scrollVCurrentY = pos;
		    evt.scrollD = evt.scrollPositionY = that.scrollVCurrentY;
			evt.maxScrollY = that.maxScrollY;
			that.handleEvents( "onscrollvertical", evt );
			if (that.checkDisableCurrentScroll()) {
			    // prevented...
                that.scrollVCurrentY = oldPos;
                return;
            }
            that._draw();
		}
		else if ( that.scrollVCurrentY === pos && pos > 0 && !this.reinit && !this.moveble && !this.lock ) {
			evt.pos = pos;
			that.handleEvents( "onscrollVEnd", evt );
		}
	};
	ScrollObject.prototype._correctScrollV = function ( that, yPos ) {
		if ( !this.settings.isVerticalScroll )
			return null;

		var events = that.eventListeners["oncorrectVerticalScroll"];
		if ( events ) {
			if ( events.length != 1 )
				return null;

			return events[0].handler.apply( that, [yPos] );
		}
		return null;
	};
	ScrollObject.prototype._correctScrollByYDelta = function ( that, delta ) {
		if ( !this.settings.isVerticalScroll )
			return null;

		var events = that.eventListeners["oncorrectVerticalScrollDelta"];
		if ( events ) {
			if ( events.length != 1 )
				return null;

			return events[0].handler.apply( that, [delta] );
		}
		return null;
	};
	ScrollObject.prototype._scrollH = function ( that, evt, pos, isTop, isBottom ) {
		if ( !this.settings.isHorizontalScroll ) {
			return;
		}
		if ( that.scrollHCurrentX !== pos ) {
			that.scrollHCurrentX = pos;
			evt.scrollD = evt.scrollPositionX = that.scrollHCurrentX;
			evt.maxScrollX = that.maxScrollX;

			that._draw();
			that.handleEvents( "onscrollhorizontal", evt );
		}
		else if ( that.scrollHCurrentX === pos && pos > 0 && !this.reinit && !this.moveble && !this.lock ) {
			evt.pos = pos;
			that.handleEvents( "onscrollHEnd", evt );
		}
	};
	ScrollObject.prototype.scrollByY = function ( delta, isAttack) {
		if ( !this.settings.isVerticalScroll ) {
			return;
		}

		var dPR = AscBrowser.retinaPixelRatio;
		var result = this._correctScrollByYDelta( this, delta );
		if ( result != null && result.isChange === true )
			delta = result.Pos;

		var destY = this.scrollVCurrentY + delta, isTop = false, isBottom = false, vend = false;

		if ( this.scrollVCurrentY + delta < 0 ) {
			destY = 0;
			isTop = true;
			isBottom = false;
		}
		else if ( this.scrollVCurrentY + delta > this.maxScrollY2 ) {
			this.handleEvents( "onscrollVEnd", destY - this.maxScrollY );
			vend = true;
			destY = this.maxScrollY2;
			isTop = false;
			isBottom = true;
		}
		var scrollCoeff = this.scrollCoeff === 0 ? 1 : this.scrollCoeff;
		this.scroller.y = destY / scrollCoeff + this.arrowPosition;
		if ( this.scroller.y < this.dragMinY )
			this.scroller.y = this.dragMinY + 1;
		else if ( this.scroller.y > this.dragMaxY )
			this.scroller.y = this.dragMaxY;

		var arrow = this.settings.showArrows ? this.arrowPosition : 0;
		if ( this.scroller.y + this.scroller.h > this.canvasH - arrow) {
			this.scroller.y -= Math.abs( this.canvasH - arrow - this.scroller.y - this.scroller.h);
		}

		this.scroller.y = Math.round(this.scroller.y);

		if ( vend ) {
			this.moveble = true;
		}
		this._scrollV( this, {}, destY, isTop, isBottom, isAttack);
		if ( vend ) {
			this.moveble = false;
		}
	};
	ScrollObject.prototype.scrollToY = function ( destY ) {
		if ( !this.settings.isVerticalScroll ) {
			return;
		}

		this.scroller.y = destY / Math.max( 1, this.scrollCoeff ) + this.arrowPosition;
		if ( this.scroller.y < this.dragMinY )
			this.scroller.y = this.dragMinY + 1;
		else if ( this.scroller.y > this.dragMaxY )
			this.scroller.y = this.dragMaxY;

		var arrow = this.settings.showArrows ? this.arrowPosition : 0;
		if ( this.scroller.y + this.scroller.h > this.canvasH - arrow ) {
			this.scroller.y -= Math.abs( this.canvasH - arrow - this.scroller.y - this.scroller.h );
		}

		this.scroller.y = Math.round(this.scroller.y);

		this._scrollV( this, {}, destY, false, false );
	};
	ScrollObject.prototype.scrollByX = function ( delta ) {
		if ( !this.settings.isHorizontalScroll ) {
			return;
		}
        var dPR = AscBrowser.retinaPixelRatio;
		var destX = this.scrollHCurrentX + delta, isTop = false, isBottom = false, hend = false;

		if ( destX < 0 ) {
			destX = 0;
			isTop = true;
			isBottom = false;
		}
		else if ( destX > this.maxScrollX2 ) {
			this.handleEvents( "onscrollHEnd", destX - this.maxScrollX );
			hend = true;
			destX = this.maxScrollX2;
			isTop = false;
			isBottom = true;
		}
		var scrollCoeff = this.scrollCoeff === 0 ? 1 : this.scrollCoeff;
		this.scroller.x = destX / scrollCoeff + this.arrowPosition;
		if ( this.scroller.x < this.dragMinX )
			this.scroller.x = this.dragMinX + 1;
		else if ( this.scroller.x > this.dragMaxX )
			this.scroller.x = this.dragMaxX;

		var arrow = this.settings.showArrows ? this.arrowPosition : 0;
		if ( this.scroller.x + this.scroller.w > this.canvasW - arrow) {
			this.scroller.x -= Math.abs( this.canvasW - arrow - this.scroller.x - this.scroller.w);
		}

		this.scroller.x = Math.round(this.scroller.x);

		if ( hend ) {
			this.moveble = true;
		}
		this._scrollH( this, {}, destX, isTop, isBottom );
		if ( hend ) {
			this.moveble = true;
		}
	};
	ScrollObject.prototype.scrollToX = function ( destX ) {
		if ( !this.settings.isHorizontalScroll ) {
			return;
		}
		this.scroller.x = destX / Math.max( 1, this.scrollCoeff ) + this.arrowPosition;
		if ( this.scroller.x < this.dragMinX )
			this.scroller.x = this.dragMinX + 1;
		else if ( this.scroller.x > this.dragMaxX )
			this.scroller.x = this.dragMaxX;

		var arrow = this.settings.showArrows ? this.arrowPosition : 0;
		if ( this.scroller.x + this.scroller.w > this.canvasW - arrow) {
			this.scroller.x -= Math.abs( this.canvasW - arrow - this.scroller.x - this.scroller.w);
		}

		this.scroller.x = Math.round(this.scroller.x);

		this._scrollH( this, {}, destX, false, false );
	};
	ScrollObject.prototype.scrollTo = function ( destX, destY ) {
		this.scrollToX( destX );
		this.scrollToY( destY );
	};
	ScrollObject.prototype.scrollBy = function ( deltaX, deltaY ) {
		this.scrollByX( deltaX );
		this.scrollByY( deltaY );
	};

	ScrollObject.prototype.roundRect = function ( x, y, width, height, radius ) {
		if ( typeof radius === "undefined" ) {
			radius = 1;
		}
		this.context.beginPath();
		this.context.moveTo( x + radius, y );
		this.context.lineTo( x + width - radius, y );
		this.context.quadraticCurveTo( x + width, y, x + width, y + radius );
		this.context.lineTo( x + width, y + height - radius );
		this.context.quadraticCurveTo( x + width, y + height, x + width - radius, y + height );
		this.context.lineTo( x + radius, y + height );
		this.context.quadraticCurveTo( x, y + height, x, y + height - radius );
		this.context.lineTo( x, y + radius );
		this.context.quadraticCurveTo( x, y, x + radius, y );
		this.context.closePath();
	};

	ScrollObject.prototype._clearContent = function () {
		this.context.clearRect( 0, 0, this.canvasW, this.canvasH );
	};

	ScrollObject.prototype._drawArrows = function () {
		var that = this;
		if (!that.settings.showArrows) {
			return;
		}
		var t = that.ArrowDrawer;
		var xDeltaBORDER = 0.5, yDeltaBORDER = 1.5;
		var roundDPR = that._roundForScale(AscBrowser.retinaPixelRatio);
		var x1 = 0, y1 = 0, x2 = 0, y2 = 0;
		var ctx = that.context;

		ctx.beginPath();
		ctx.fillStyle =  that.settings.scrollerColor;
		var arrowImage = that.ArrowDrawer.ImageTop;
		var imgContext = arrowImage.getContext('2d');
		imgContext.globalCompositeOperation = "source-in";
		imgContext.fillStyle = that.settings.arrowColor;
        ctx.lineWidth = roundDPR;
		if (that.settings.isVerticalScroll) {
			for (var i = 0; i < 2; i++) {
				imgContext.fillRect( x1 + xDeltaBORDER * ctx.lineWidth,  0, t.SizeW - roundDPR, t.SizeH - roundDPR);
				ctx.fillRect( x1,  y1 + roundDPR, t.SizeW, t.SizeH);
				ctx.drawImage(arrowImage, x1, y2, t.SizeW, t.SizeH);

				if (t.IsDrawBorders) {
					ctx.strokeStyle = that.settings.strokeStyleNone;
					ctx.rect(x1 + xDeltaBORDER * ctx.lineWidth, y1  + yDeltaBORDER * ctx.lineWidth, t.SizeW - roundDPR, t.SizeH - roundDPR);
					ctx.stroke();
				}

				y1 = that.canvasH - t.SizeH - 2 * roundDPR;
                y2 = that.canvasH - t.SizeH;
				arrowImage = that.ArrowDrawer.ImageBottom;
				imgContext = arrowImage.getContext('2d');
				imgContext.globalCompositeOperation = "source-in";
				imgContext.fillStyle = that.settings.arrowColor;
			}
		}

		var arrowImage = that.ArrowDrawer.ImageLeft;
		var imgContext = arrowImage.getContext('2d');
		imgContext.globalCompositeOperation = "source-in";
		imgContext.fillStyle = that.settings.arrowColor;
		if (that.settings.isHorizontalScroll) {
			for (var i = 0; i < 2; i++) {
				imgContext.fillRect( 0,   y1 + yDeltaBORDER * ctx.lineWidth, t.SizeW - roundDPR, t.SizeH - roundDPR);
				ctx.fillRect( x1 + roundDPR, y1, t.SizeW, t.SizeH);
				ctx.drawImage(arrowImage, x2, y2, t.SizeW, t.SizeH);

				if (t.IsDrawBorders) {
					ctx.strokeStyle = that.settings.strokeStyleNone;
					ctx.lineWidth = roundDPR;
					ctx.rect(x1 + yDeltaBORDER * ctx.lineWidth, y1 + xDeltaBORDER * ctx.lineWidth, t.SizeW - roundDPR, t.SizeH - roundDPR);
					ctx.stroke();
				}

				x1 = that.canvasW - t.SizeW - 2 * roundDPR;
                x2 = that.canvasW - t.SizeW;
                y2 = 0;
				arrowImage = that.ArrowDrawer.ImageRight;
				imgContext = arrowImage.getContext('2d');
				imgContext.globalCompositeOperation = "source-in";
				imgContext.fillStyle = that.settings.arrowColor;
			}
		}
	};

	ScrollObject.prototype._drawScroll = function (fillColor, targetColor, strokeColor) {
		fillColor = Math.round(fillColor);
		targetColor = Math.round(targetColor);
		strokeColor = Math.round(strokeColor);

		var that = this;
		that.context.beginPath();
		var roundDPR = this._roundForScale(AscBrowser.retinaPixelRatio);
		that.context.lineWidth = roundDPR;

		if (that.settings.isVerticalScroll) {
			var _y = that.settings.showArrows ? that.arrowPosition : 0,
				_h = that.canvasH - (_y << 1);

			if (_h > 0) {
				that.context.rect(0, _y, that.canvasW, _h);
			}
		} else if (that.settings.isHorizontalScroll) {
			var _x = that.settings.showArrows ? that.arrowPosition : 0,
				_w = that.canvasW - (_x << 1);

			if (_w > 0) {
				that.context.rect(_x, 0, _w, that.canvasH);
			}
		}

		switch (that.animState) {

			case AnimationType.SCROLL_HOVER: {
				that.context.fillStyle = that.settings.scrollBackgroundColorHover;
				break;
			}
			case AnimationType.SCROLL_ACTIVE: {
				that.context.fillStyle = that.settings.scrollBackgroundColorActive;
				break;
			}
			case AnimationType.NONE:
			default: {
				that.context.fillStyle = that.settings.scrollBackgroundColor;
				break;
			}

		}

		that.context.fill();
		that.context.beginPath();

		if (that.settings.isVerticalScroll && that.maxScrollY != 0) {
			var _y = that.scroller.y >> 0, arrow = that.settings.showArrows ? that.arrowPosition : 0;
			if (_y < arrow) {
				_y = arrow;
			}
			var _b = Math.round(that.scroller.y + that.scroller.h);// >> 0;
			if (_b > (that.canvasH - arrow - 1)) {
				_b = that.canvasH - arrow - 1;
			}

			if (_b > _y) {
				that.roundRect(that.scroller.x - 0.5 * that.context.lineWidth, _y + 0.5 * that.context.lineWidth, that.scroller.w - roundDPR, that.scroller.h - roundDPR, that.settings.cornerRadius * roundDPR);
			}
		} else if (that.settings.isHorizontalScroll && that.maxScrollX != 0) {
			var _x = that.scroller.x >> 0, arrow = that.settings.showArrows ? that.arrowPosition : 0;
			if (_x < arrow) {
				_x = arrow;
			}
			var _r = (that.scroller.x + that.scroller.w) >> 0;
			if (_r > (that.canvasW - arrow - 2)) {
				_r = that.canvasW - arrow - 1;
			}

			if (_r > _x) {
				that.roundRect(_x +  0.5 * that.context.lineWidth, that.scroller.y -  0.5 * that.context.lineWidth, that.scroller.w - roundDPR, that.scroller.h - roundDPR, that.settings.cornerRadius * roundDPR);
			}
		}

		that.context.fillStyle = "rgb(" + fillColor + "," + fillColor + "," + fillColor + ")"
		that.context.fill();



		that.context.strokeStyle = "rgb(" + strokeColor + "," + strokeColor + "," + strokeColor + ")";
		that.strokeColor = strokeColor;
		that.context.stroke();

		var ctx_piperImg, _data, px, img, x, y;

		//drawing scroll stripes
		if (that._checkPiperImagesV()) {

			x = Math.round((that.scroller.w - that.piperImgVert.width) / 2);
			y = Math.floor((that.scroller.y + Math.floor(that.scroller.h / 2) - 6 * AscBrowser.retinaPixelRatio));

			ctx_piperImg = that.piperImgVert.getContext('2d');
			ctx_piperImg.globalCompositeOperation = "source-in";
			ctx_piperImg.fillStyle = "rgb(" + targetColor + "," +
				targetColor + "," +
				targetColor + ")";
			ctx_piperImg.fillRect(0, 0, that.scroller.w - 1, that.scroller.h - 1);

			img = that.piperImgVert;
		} else if (that._checkPiperImagesH()) {
            x = Math.round(that.scroller.x + (that.scroller.w - that.piperImgHor.width) / 2);
            y = Math.round((that.scroller.h - that.piperImgHor.height) / 2);

			ctx_piperImg = that.piperImgHor.getContext('2d');
			ctx_piperImg.globalCompositeOperation = "source-in";
			ctx_piperImg.fillStyle = "rgb(" + targetColor + "," +
				targetColor + "," +
				targetColor + ")";
			ctx_piperImg.fillRect(0, 0, that.scroller.w - 1, that.scroller.h - 1);

			img = that.piperImgHor;
		}

		if (img)
			that.context.drawImage(img, x, y);

		that.scrollColor = fillColor;
		that.targetColor = targetColor;
	};

	ScrollObject.prototype._animateArrow = function (fadeIn, curArrowType, backgroundColorUnfade) {
		var that = this;
		var roundDPR = that._roundForScale(AscBrowser.retinaPixelRatio);
		var sizeW = that.ArrowDrawer.SizeW, sizeH = that.ArrowDrawer.SizeH;

		if (!that.settings.showArrows || !curArrowType) {
			return;
		}
        var cnvs = document.createElement('canvas'), arrowType,
            ctx = cnvs.getContext('2d'), context = that.context,
            hoverColor = _HEXTORGB_(that.settings.scrollerHoverColor).R, defaultColor = _HEXTORGB_(that.settings.scrollerColor).R,
            arrowColor = _HEXTORGB_(that.settings.arrowColor).R,
			arrowHoverColor = _HEXTORGB_(that.settings.arrowHoverColor).R,
		    strokeColor = _HEXTORGB_(that.settings.strokeStyleNone).R,
			strokeHoverColor = _HEXTORGB_(that.settings.strokeStyleOver).R;

		cnvs.width = sizeW;
		cnvs.height = sizeH;

		if (curArrowType === ArrowType.ARROW_TOP || curArrowType === ArrowType.ARROW_LEFT) {
			arrowType = this.firstArrow;
		} else if (curArrowType === ArrowType.ARROW_BOTTOM || curArrowType === ArrowType.ARROW_RIGHT) {
			arrowType = this.secondArrow;
		} else return;


		var x = 0, y = 0, fillRectX = 0, fillRectY = 0, strokeRectX = 0, strokeRectY = 0;
		var arrowImage = that.settings.isVerticalScroll ? that.ArrowDrawer.ImageTop : that.ArrowDrawer.ImageLeft;

		//what type of arrow to draw
		switch (curArrowType) {
			case ArrowType.ARROW_TOP: {
				fillRectY = roundDPR;
				break;
			}
			case ArrowType.ARROW_BOTTOM: {
				y = that.canvasH - sizeH;
				fillRectY = -roundDPR;
				strokeRectY = -2 * roundDPR;
				arrowImage = that.ArrowDrawer.ImageBottom;
				break;
			}
			case ArrowType.ARROW_RIGHT: {
				y = 0;
				x = that.canvasW - sizeW;
				fillRectX = -roundDPR;
				strokeRectX = -roundDPR;
				strokeRectY = -roundDPR;
				arrowImage = that.ArrowDrawer.ImageRight;
				break;
			}
			case ArrowType.ARROW_LEFT: {
				y = 0;
				x = 0;
				fillRectX = roundDPR;
				strokeRectX = roundDPR;
				strokeRectY =  -roundDPR;
				break;
			}
		}

		var stepsCount = 17;
		var step = Math.abs((defaultColor - hoverColor)) / stepsCount;

		//dimming the arrow
		if (fadeIn) {
			if(arrowType.arrowColor === arrowHoverColor && arrowType.arrowBackColor === hoverColor) {
				return;
			}

			if (arrowType.arrowBackColor - hoverColor > step) {
				arrowType.arrowBackColor -= step;
			} else if (arrowType.arrowBackColor - hoverColor < -step) {
				arrowType.arrowBackColor += step;
			} else {
				arrowType.arrowBackColor = hoverColor;
			}

			step = Math.abs((strokeColor - strokeHoverColor)) / stepsCount;

			if (arrowType.arrowStrokeColor - strokeHoverColor > step) {
				arrowType.arrowStrokeColor -= step;
			} else if (arrowType.arrowStrokeColor - strokeHoverColor < -step) {
				arrowType.arrowStrokeColor += step;
			} else {
				arrowType.arrowStrokeColor = strokeHoverColor;
			}

			step = Math.abs((arrowColor - arrowHoverColor)) / stepsCount;

			if (arrowType.arrowColor - arrowHoverColor > step) {
				arrowType.arrowColor -= step;
			} else if (arrowType.arrowColor - arrowHoverColor < -step) {
				arrowType.arrowColor += step;
			} else {
				arrowType.arrowColor = arrowHoverColor;
			}
		} else
			//reverse dimming
		if (fadeIn === false) {
			if(arrowType.arrowColor === arrowColor && arrowType.arrowBackColor === defaultColor) {
				return;
			}

			if (arrowType.arrowBackColor - defaultColor < -step) {
				arrowType.arrowBackColor += step;
			} else if (arrowType.arrowBackColor - defaultColor > step) {
				arrowType.arrowBackColor -= step;
			} else {
				arrowType.arrowBackColor = defaultColor;
			}

			step = Math.abs((arrowColor - arrowHoverColor)) / stepsCount;

			if (arrowType.arrowColor - arrowColor > step) {
				arrowType.arrowColor -= step;
			} else if (arrowType.arrowColor - arrowColor < -step) {
				arrowType.arrowColor += step;
			} else {
				arrowType.arrowColor = arrowColor;
			}

			step = Math.abs((strokeColor - strokeHoverColor)) / stepsCount;

			if (arrowType.arrowStrokeColor - strokeColor > step) {
				arrowType.arrowStrokeColor -= step;
			} else if (arrowType.arrowStrokeColor - strokeColor < -step) {
				arrowType.arrowStrokeColor += step;
			} else {
				arrowType.arrowStrokeColor = strokeColor;
			}
		} else {
			//instant change arrow color
			arrowType.arrowBackColor = backgroundColorUnfade;
			switch(backgroundColorUnfade) {
				case _HEXTORGB_(that.settings.scrollerColor).R:
					arrowType.arrowColor = _HEXTORGB_(that.settings.arrowColor).R;
					arrowType.arrowStrokeColor = _HEXTORGB_(that.settings.strokeStyleNone).R;
					break;
				case _HEXTORGB_(that.settings.scrollerHoverColor).R:
					arrowType.arrowColor = _HEXTORGB_(that.settings.arrowHoverColor).R;
					arrowType.arrowStrokeColor = _HEXTORGB_(that.settings.strokeStyleOver).R;
					break;
				case _HEXTORGB_(that.settings.scrollerActiveColor).R:
					arrowType.arrowColor = _HEXTORGB_(that.settings.arrowActiveColor).R;
					arrowType.arrowStrokeColor = _HEXTORGB_(that.settings.strokeStyleActive).R;
					break;

			}
		}

		ctx = that.context;

        ctx.beginPath();
        var arrowBackColor = Math.round(arrowType.arrowBackColor);
        ctx.fillStyle = "rgb(" + arrowBackColor + "," +
			arrowBackColor + "," +
			arrowBackColor + ")";

        ctx.fillRect( x + fillRectX,  y +  fillRectY, sizeW, sizeH);

		if (that.ArrowDrawer.IsDrawBorders) {
			var arrowStrokeColor = Math.round(arrowType.arrowStrokeColor)
			ctx.strokeStyle = "rgb(" + arrowStrokeColor + "," + arrowStrokeColor + "," + arrowStrokeColor + ")";
			ctx.lineWidth = roundDPR;
            ctx.rect(x + 0.5 * ctx.lineWidth + strokeRectX, y + 1.5 * ctx.lineWidth + strokeRectY, sizeW - roundDPR, sizeH - roundDPR);
            ctx.stroke();
		}

		//drawing arrow icon
		var imgContext = arrowImage.getContext('2d');
		imgContext.globalCompositeOperation = "source-in";
		var arrowColor = Math.round(arrowType.arrowColor);
		imgContext.fillStyle = "rgb(" + arrowColor + "," +
			arrowColor + "," +
			arrowColor + ")";
		imgContext.fillRect(0.5, 1.5, sizeW , sizeH);
        ctx.drawImage(arrowImage,  x, y, sizeW, sizeH);
		ctx.closePath();
		context.drawImage(cnvs, x, y);

		if (fadeIn === undefined)
		return;

		that.fadeTimeoutArrows = setTimeout(function () {
			that._animateArrow(fadeIn, curArrowType, backgroundColorUnfade)
		}, that.settings.fadeInFadeOutDelay);
	};

	ScrollObject.prototype._animateScroll = function (fadeIn) {
		var that = this;
		var hoverColor = _HEXTORGB_(that.settings.scrollerHoverColor).R,
			defaultColor = _HEXTORGB_(that.settings.scrollerColor).R,
			targetDefaultColor = _HEXTORGB_(that.settings.targetColor).R,
			targetHoverColor = _HEXTORGB_(that.settings.targetHoverColor).R,
			strokeHoverColor = _HEXTORGB_(that.settings.strokeStyleOver).R,
			strokeColor = _HEXTORGB_(that.settings.strokeStyleNone).R;

		that.context.beginPath();
		that._drawScroll(that.scrollColor, that.targetColor, that.strokeColor);

		//animation end condition
		if ((fadeIn && that.scrollColor === hoverColor && that.targetColor === targetHoverColor) || (!fadeIn && that.scrollColor === defaultColor && that.targetColor === targetDefaultColor)) {
			return;
		}

		var stepsCount = 17;
		var step = Math.abs((defaultColor - hoverColor)) / stepsCount;
		
		//dimming the scroll
		if (fadeIn) {

			if (that.scrollColor - hoverColor > step) {
				that.scrollColor -= step;
			} else if (that.scrollColor - hoverColor < -step) {
				that.scrollColor += step;
			} else {
				that.scrollColor = hoverColor;
			}

			step =  Math.abs((targetDefaultColor - targetHoverColor)) / stepsCount;

			if (that.targetColor - targetHoverColor > step) {
				that.targetColor -= step;
			} else if (that.targetColor - targetHoverColor < -step) {
				that.targetColor += step;
			} else {
				that.targetColor = targetHoverColor;
			}

			step = Math.abs((strokeColor - strokeHoverColor)) / stepsCount;

			if (that.strokeColor - strokeHoverColor > step) {
				that.strokeColor -= step;
			} else if (that.strokeColor - strokeHoverColor < -step) {
				that.strokeColor += step;
			} else {
				that.strokeColor = strokeHoverColor;
			}

		} else
			//reverse dimming
		if (fadeIn === false) {
			if (that.scrollColor - defaultColor > step) {
				that.scrollColor -= step;
			} else if (that.scrollColor - defaultColor < -step) {
				that.scrollColor += step;
			} else {
				that.scrollColor = defaultColor;
			}

			step = Math.abs((targetDefaultColor - targetHoverColor)) / stepsCount;

			if (that.targetColor - targetDefaultColor > step) {
				that.targetColor -= step;
			} else if (that.targetColor - targetDefaultColor < -step) {
				that.targetColor += step;
			} else {
				that.targetColor = targetDefaultColor;
				that.strokeColor = strokeColor;
			}

			step = Math.abs((strokeColor - strokeHoverColor)) / stepsCount;

			if (that.strokeColor - strokeColor > step) {
				that.strokeColor -= step;
			} else if (that.strokeColor - strokeColor < -step) {
				that.strokeColor += step;
			} else {
				that.strokeColor = strokeColor;
			}

		}

		that.fadeTimeoutScroll = setTimeout(function () {
			that._animateScroll(fadeIn)
		}, that.settings.fadeInFadeOutDelay);
	};

	ScrollObject.prototype._doAnimation = function (lastAnimState) {
		var that = this, secondArrow,
			hoverColor = _HEXTORGB_(that.settings.scrollerHoverColor).R,
			defaultColor = _HEXTORGB_(that.settings.scrollerColor).R,
			activeColor = _HEXTORGB_(that.settings.scrollerActiveColor).R,
		    targetColor = _HEXTORGB_(that.settings.targetColor).R,
			targetHoverColor = _HEXTORGB_(that.settings.targetHoverColor).R,
			targetActiveColor = _HEXTORGB_(that.settings.targetActiveColor).R,
			strokeColor = _HEXTORGB_(that.settings.strokeStyleNone).R,
			strokeHoverColor = _HEXTORGB_(that.settings.strokeStyleOver).R,
			strokeActiveColor = _HEXTORGB_(that.settings.strokeStyleActive).R;

		switch(that.arrowState) {
			case ArrowType.ARROW_TOP:
				secondArrow = ArrowType.ARROW_BOTTOM;
				break;
			case ArrowType.ARROW_BOTTOM:
				secondArrow = ArrowType.ARROW_TOP;
				break;
			case ArrowType.ARROW_LEFT:
				secondArrow = ArrowType.ARROW_RIGHT;
				break;
			case ArrowType.ARROW_RIGHT:
				secondArrow = ArrowType.ARROW_LEFT;
				break;
		}

		//current and previous scroll state
		if (that.animState === AnimationType.NONE && lastAnimState === AnimationType.NONE) {
			that._drawScroll(defaultColor, targetColor, strokeColor);
		} else if (that.animState === AnimationType.SCROLL_HOVER && lastAnimState === AnimationType.SCROLL_HOVER) {
			that._animateArrow(false, that.arrowState);
			that._animateScroll(true);
		} else if (that.animState === AnimationType.ARROW_HOVER && lastAnimState === AnimationType.NONE) {
			that._animateArrow(false, secondArrow);
			that._animateArrow(true, that.arrowState);
			that._animateScroll(true);
		} else if (that.animState === AnimationType.SCROLL_HOVER && lastAnimState === AnimationType.NONE) {
			that._animateArrow(false, that.arrowState);
			that._animateScroll(true);
		} else if (that.animState === AnimationType.NONE && lastAnimState === AnimationType.ARROW_HOVER) {
			that._animateArrow(false, that.arrowState);
			that._animateScroll(false);
		} else if (that.animState === AnimationType.NONE && lastAnimState === AnimationType.SCROLL_ACTIVE) {
			that._animateArrow(false, that.arrowState);
			that._drawScroll(defaultColor, targetColor, strokeColor);
		} else if (that.animState === AnimationType.SCROLL_HOVER && lastAnimState === AnimationType.ARROW_HOVER) {
			that._animateArrow(false, that.arrowState);
			that._animateScroll(true);
		} else if (that.animState === AnimationType.NONE && lastAnimState === AnimationType.SCROLL_HOVER) {
			that._animateArrow(false, that.arrowState);
			that._animateScroll(false);
		} else if (that.animState === AnimationType.ARROW_HOVER && lastAnimState === AnimationType.SCROLL_HOVER) {
			that._animateArrow(false, secondArrow);
			that._animateArrow(true, that.arrowState);
			that._animateScroll(true);
		} else if (this.animState === AnimationType.SCROLL_HOVER && lastAnimState === AnimationType.SCROLL_ACTIVE) {
			that._animateArrow(undefined, that.arrowState, defaultColor);
			that._drawScroll(hoverColor, targetHoverColor, strokeHoverColor);
		} else if (this.animState === AnimationType.SCROLL_HOVER && lastAnimState === AnimationType.ARROW_ACTIVE) {
			that._animateArrow(undefined, that.arrowState, defaultColor);
			that._drawScroll(hoverColor, targetHoverColor, strokeHoverColor);
		} else if (this.animState === AnimationType.ARROW_ACTIVE) {
            that._animateArrow(undefined, that.arrowState, activeColor);
			that._animateScroll(true);
		} else if (this.animState === AnimationType.ARROW_HOVER && lastAnimState === AnimationType.ARROW_ACTIVE) {
			//if different arrows
			if (this.lastArrowState && this.lastArrowState !== this.arrowState) {
				that._animateArrow(undefined, secondArrow, defaultColor);
				that._animateArrow(true, that.arrowState);
				that._animateScroll(true);
			} else {
				that._animateArrow(undefined, that.arrowState, hoverColor);
				that._animateScroll(true);
			}
		} else if (this.animState === AnimationType.NONE && lastAnimState === AnimationType.ARROW_ACTIVE) {
			that._animateArrow(undefined, that.arrowState, defaultColor);
			that._animateScroll(false);
		} else if (this.animState === AnimationType.ARROW_HOVER && lastAnimState === AnimationType.SCROLL_ACTIVE) {
			that._animateArrow(false, secondArrow);
			that._animateArrow(true, that.arrowState);

			if (that.mouseUp && !that.mouseDown) {
				that._drawScroll(hoverColor, targetHoverColor, strokeHoverColor);
			}
			else if(that.mouseDown && that.scrollerMouseDown) {
				that._drawScroll(activeColor, targetActiveColor, strokeActiveColor);
			}
		} else if (this.animState === AnimationType.SCROLL_ACTIVE) {
			that._animateArrow(false, that.arrowState);
			that._drawScroll(activeColor, targetActiveColor, strokeActiveColor);
		} else if (this.animState === AnimationType.ARROW_HOVER && lastAnimState === AnimationType.ARROW_HOVER) {
			    that._animateArrow(true, that.arrowState);
			    that._drawScroll(hoverColor, targetHoverColor, strokeHoverColor);
		} else return;
	};

    ScrollObject.prototype._draw = function () {

		clearTimeout(this.fadeTimeoutScroll);
		this.fadeTimeoutScroll = null;
		clearTimeout(this.fadeTimeoutArrows);
		this.fadeTimeoutArrows = null;

        //scroll animation
        this._doAnimation(this.lastAnimState);
        this.lastAnimState = this.animState;
	};

	ScrollObject.prototype._checkPiperImagesV = function() {
		if ( this.settings.isVerticalScroll && this.maxScrollY != 0 && this.scroller.h >= Math.round(13 * AscBrowser.retinaPixelRatio) )
			return true;
		return false;
	};
	ScrollObject.prototype._checkPiperImagesH = function() {
		if ( this.settings.isHorizontalScroll && this.maxScrollX != 0 && this.scroller.w >= Math.round(13 * AscBrowser.retinaPixelRatio) )
			return true;
		return false;
	};

	ScrollObject.prototype._setDimension = function ( h, w ) {

		var dPR = AscBrowser.retinaPixelRatio;
		if(this.canvasH ===  Math.round(h * dPR) && this.canvasW === Math.round(w * dPR)) {
			this.isResizeArrows = false;
			return;
		}
		this.isResizeArrows = true;
        this.canvasH = Math.round(h * dPR);
        this.canvasW = Math.round(w * dPR);
		this.canvasOriginalH = h;
        this.canvasOriginalW = w;
		this.canvas.height = this.canvasH;
		this.canvas.width =  this.canvasW;
	};
	ScrollObject.prototype._setScrollerHW = function () {
		var dPR = AscBrowser.retinaPixelRatio;
		if ( this.settings.isVerticalScroll ) {
			this.scroller.x = this._roundForScale(dPR);
			this.scroller.w = Math.round((this.canvasOriginalW - 1) * dPR);
			if(this.settings.slimScroll) {
				this.settings.arrowSizeW = this.settings.arrowSizeH = this.scroller.w;
			}
			if ( this.settings.showArrows )
				this.ArrowDrawer.InitSize( this.settings.arrowSizeW, this.settings.arrowSizeH, this.IsRetina );
		}
		else if ( this.settings.isHorizontalScroll ) {
			this.scroller.y = this._roundForScale(dPR);
			this.scroller.h = Math.round((this.canvasOriginalH - 1) * dPR);
			if(this.settings.slimScroll) {
				this.settings.arrowSizeH = this.settings.arrowSizeW = this.scroller.h;
			}
			if ( this.settings.showArrows )
				this.ArrowDrawer.InitSize( this.settings.arrowSizeH, this.settings.arrowSizeW, this.IsRetina );
		}
	};
	ScrollObject.prototype._MouseHoverOnScroller = function ( mp ) {
		if(this.settings.isVerticalScroll && mp.x >= 0 && mp.x <= this.scroller.x + this.scroller.w &&
			mp.y >= this.scroller.y && mp.y <= this.scroller.y + this.scroller.h) {
			return true;
		} else if(this.settings.isHorizontalScroll && mp.x >= this.scroller.x && mp.x <= this.scroller.x + this.scroller.w &&
			mp.y >= 0 && mp.y <= this.scroller.y + this.scroller.h) {
			return true;
		} else {
			return false;
		}
	};

    ScrollObject.prototype._MouseArrowHover = function (mp) {

		var arrowDim =this.settings.arrowDim;
        if (this.settings.isVerticalScroll) {
            if (
                mp.x >= 0 &&
                mp.x <= this.canvasW &&
                mp.y >= 0 &&
                mp.y <= arrowDim
            ) {
                return ArrowType.ARROW_TOP;
            } else if (
                mp.x >= 0 &&
                mp.x <= this.canvasW &&
                mp.y >= this.canvasH - arrowDim &&
                mp.y <= this.canvasH
            ) {
                return ArrowType.ARROW_BOTTOM;
            } else return ArrowType.NONE;
        }
        if (this.settings.isHorizontalScroll) {
            if (
                mp.x >= 0 &&
                mp.x <= arrowDim &&
                mp.y >= 0 &&
                mp.y <= this.canvasH
            ) {
                return ArrowType.ARROW_LEFT;
            } else if (
                mp.x >= this.canvasW - arrowDim &&
                mp.x <= this.canvasW &&
                mp.y >= 0 &&
                mp.y <= this.canvasH
            ) {
                return ArrowType.ARROW_RIGHT;
            } else return ArrowType.NONE;
        }
    };

	ScrollObject.prototype._arrowDownMouseDown = function () {
		var that = this, scrollTimeout, isFirst = true,
			doScroll = function () {
				if ( that.settings.isVerticalScroll )
					that.scrollByY( that.settings.vscrollStep );
				else if ( that.settings.isHorizontalScroll )
					that.scrollByX( that.settings.hscrollStep );

				if(that.mouseDown)
				scrollTimeout = setTimeout( doScroll, isFirst ? that.settings.initialDelay : that.settings.arrowRepeatFreq );
				isFirst = false;
			};
		doScroll();
		this.bind( "mouseup.main mouseout", function () {
			scrollTimeout && clearTimeout( scrollTimeout );
			scrollTimeout = null;
		} );
	};
	ScrollObject.prototype._arrowUpMouseDown = function () {
		var that = this, scrollTimeout, isFirst = true,
			doScroll = function () {
				if ( that.settings.isVerticalScroll )
					that.scrollByY( -that.settings.vscrollStep );
				else if ( that.settings.isHorizontalScroll )
					that.scrollByX( -that.settings.hscrollStep );

                if(that.mouseDown)
				scrollTimeout = setTimeout( doScroll, isFirst ? that.settings.initialDelay : that.settings.arrowRepeatFreq );
				isFirst = false;
			};
		doScroll();
		this.bind( "mouseup.main mouseout", function () {
			scrollTimeout && clearTimeout( scrollTimeout );
			scrollTimeout = null;
		} )
	};

	ScrollObject.prototype.getCurScrolledX = function () {
		return this.scrollHCurrentX;
	};
	ScrollObject.prototype.getCurScrolledY = function () {
		return this.scrollVCurrentY;
	};
	ScrollObject.prototype.getMaxScrolledY = function () {
		return this.maxScrollY;
	};
	ScrollObject.prototype.getMaxScrolledX = function () {
		return this.maxScrollX;
	};
	ScrollObject.prototype.getIsLockedMouse = function () {
		return (this.that.mouseDownArrow || this.that.mouseDown);
	};
	/************************************************************************/
	/*events*/
	ScrollObject.prototype.evt_mousemove = function ( e ) {

        if (this.style)
            this.style.cursor = "default";

        var evt = e || window.event;

		if ( evt.preventDefault )
			evt.preventDefault();
		else
			evt.returnValue = false;

		var mousePos = this.that.getMousePosition( evt );
		this.that.EndMousePosition.x = mousePos.x;
		this.that.EndMousePosition.y = mousePos.y;
		var arrowHover = this.that._MouseArrowHover(mousePos);
		var dPR = AscBrowser.retinaPixelRatio;

		//arrow pressed
		if (this.that.settings.showArrows && this.that.mouseDownArrow) {
		    if (arrowHover && arrowHover === this.that.arrowState) {
				this.that.arrowState = arrowHover;
			}
            this.that.animState = AnimationType.ARROW_ACTIVE;
		} else if (!this.that.mouseDownArrow) {
			if (!arrowHover || !this.that.settings.showArrows) {
				this.that.animState = AnimationType.SCROLL_HOVER;
			} else {
				this.that.animState = AnimationType.ARROW_HOVER;
                this.that.arrowState = arrowHover;
			}
		}

		if ( this.that.mouseDown && this.that.scrollerMouseDown ) {
			this.that.moveble = true;
		}
		else {
			this.that.moveble = false;
		}

		if ( this.that.settings.isVerticalScroll ) {
			if ( this.that.moveble && this.that.scrollerMouseDown ) {
				var isTop = false, isBottom = false;
				if (arrowHover && this.that.settings.showArrows) {
					this.that.animState = AnimationType.ARROW_HOVER;
				} else {
					this.that.animState = AnimationType.SCROLL_ACTIVE;
				}
				var _dlt = this.that.EndMousePosition.y - this.that.StartMousePosition.y;
				_dlt = _dlt >= 0 ? Math.floor(_dlt) : Math.ceil(_dlt);

				if ( this.that.EndMousePosition.y == this.that.StartMousePosition.y ) {
					return;
				}
				else if ( this.that.EndMousePosition.y < this.that.arrowPosition ) {
					this.that.EndMousePosition.y = this.that.arrowPosition;
					this.that.scroller.y = this.that.arrowPosition;
				}
				else if ( this.that.EndMousePosition.y > this.that.canvasH - this.that.arrowPosition ) {
					this.that.EndMousePosition.y = this.that.canvasH - this.that.arrowPosition;
					this.that.scroller.y = this.that.canvasH - this.that.arrowPosition - this.that.scroller.h;
				}
				else {
					var dltY = _dlt > 0 ? this.that.canvasH - this.that.arrowPosition - this.that.scroller.h - this.that.scroller.y : this.that.arrowPosition - this.that.scroller.y;
					_dlt = (_dlt > 0) ? (dltY < _dlt ? dltY : _dlt) : (dltY > _dlt ? dltY : _dlt);

                     if(_dlt > 0 && this.that.scroller.y + this.that.scroller.h + _dlt +  Math.round(dPR) <= this.that.canvasH - this.that.arrowPosition ||
					  	(_dlt < 0 && this.that.scroller.y + _dlt >= this.that.arrowPosition)) {
						 this.that.scroller.y += _dlt;
					 } else if(_dlt > 0) {
						 this.that.scroller.y =  this.that.canvasH - this.that.arrowPosition - this.that.scroller.h;
					 } else if(_dlt < 0) {
						 this.that.scroller.y =  this.that.arrowPosition;
					 }

				}

				var destY = (this.that.scroller.y - this.that.arrowPosition) * this.that.scrollCoeff;
				//var result = editor.WordControl.CorrectSpeedVerticalScroll(destY);
				var result = this.that._correctScrollV( this.that, destY );
				if ( result != null && result.isChange === true ) {
					destY = result.Pos;
				}

				this.that._scrollV( this.that, evt, destY, isTop, isBottom );
				this.that.StartMousePosition.x = this.that.EndMousePosition.x;
				this.that.StartMousePosition.y = this.that.EndMousePosition.y;
			}
		}
		else if ( this.that.settings.isHorizontalScroll ) {
			if ( this.that.moveble && this.that.scrollerMouseDown ) {

				var isTop = false, isBottom = false;
                if (arrowHover && this.that.settings.showArrows) {
                    this.that.animState = AnimationType.ARROW_HOVER;
                } else {
                    this.that.animState = AnimationType.SCROLL_ACTIVE;
                }
				var _dlt =this.that.EndMousePosition.x - this.that.StartMousePosition.x;
                _dlt = _dlt >= 0 ? Math.floor(_dlt) : Math.ceil(_dlt);

				if ( this.that.EndMousePosition.x == this.that.StartMousePosition.x )
					return;
				else if ( this.that.EndMousePosition.x < this.that.arrowPosition ) {
					this.that.EndMousePosition.x = this.that.arrowPosition;
					this.that.scroller.x = this.that.arrowPosition;
				}
				else if ( this.that.EndMousePosition.x > this.that.canvasW - this.that.arrowPosition ) {
					this.that.EndMousePosition.x = this.that.canvasW - this.that.arrowPosition;
					this.that.scroller.x = this.that.canvasW - this.that.arrowPosition - this.that.scroller.w;
				}
				else {
					var dltX = _dlt > 0 ? this.that.canvasW - this.that.arrowPosition - this.that.scroller.w - this.that.scroller.x : this.that.arrowPosition - this.that.scroller.x;
					_dlt = (_dlt > 0) ? (dltX < _dlt ? dltX : _dlt) : (dltX > _dlt ? dltX : _dlt);

					if(_dlt > 0 && this.that.scroller.x + this.that.scroller.w + _dlt <= this.that.canvasW - this.that.arrowPosition ||
						(_dlt < 0 && this.that.scroller.x + _dlt >= this.that.arrowPosition)) {
						this.that.scroller.x += _dlt;
					} else if(_dlt > 0) {
						this.that.scroller.x =  this.that.canvasW - this.that.arrowPosition - this.that.scroller.w;
					} else if(_dlt < 0) {
						this.that.scroller.x =  this.that.arrowPosition;
					}
				}
				var destX = (this.that.scroller.x - this.that.arrowPosition) * this.that.scrollCoeff;

				this.that._scrollH( this.that, evt, destX, isTop, isBottom );

				this.that.StartMousePosition.x = this.that.EndMousePosition.x;
				this.that.StartMousePosition.y = this.that.EndMousePosition.y;
			}
		}

        this.that.moveble = false;

		if ( this.that.lastAnimState != this.that.animState) {
			this.that._draw();
		}

	};
	ScrollObject.prototype.evt_mouseout = function ( e ) {

		var evt = e || window.event;

		if ( this.that.settings.showArrows ) {
			this.that.mouseDown = this.that.mouseDownArrow ? false : this.that.mouseDown;
			this.that.mouseDownArrow = false;
			this.that.handleEvents( "onmouseout", evt );
		}

		if (this.that.mouseDown && this.that.scrollerMouseDown) {
			this.that.animState = AnimationType.SCROLL_ACTIVE;
		} else this.that.animState = AnimationType.NONE;

			this.that._draw();

	};
	ScrollObject.prototype.evt_mouseover = function ( e ) {
		this.that.mouseover = true;
	};

	ScrollObject.prototype.evt_mouseup = function ( e ) {
		var evt = e || window.event;

		this.that.handleEvents( "onmouseup", evt );

		// prevent pointer events on all iframes (while only plugin!)
		if (window.g_asc_plugins)
			window.g_asc_plugins.enablePointerEvents();

		if ( evt.preventDefault )
			evt.preventDefault();
		else
			evt.returnValue = false;

		this.that.mouseDown = false;
		var mousePos = this.that.getMousePosition( evt );
		var arrowHover = this.that._MouseArrowHover(mousePos);
		var mouseHover = this.that._MouseHoverOnScroller( mousePos );
		this.that.scrollTimeout && clearTimeout( this.that.scrollTimeout );
		this.that.scrollTimeout = null;


		if ( this.that.scrollerMouseDown ) {
			this.that.mouseUp = true;
			this.that.scrollerMouseDown = false;
			this.that.mouseDownArrow = false;
			if ( this.that._MouseHoverOnScroller( mousePos ) ) {
				this.that.animState = AnimationType.SCROLL_HOVER;
			}
			else {
				if(arrowHover && this.that.settings.showArrows) {
					this.that.lastAnimState = AnimationType.SCROLL_ACTIVE;
					this.that.animState = AnimationType.ARROW_HOVER;
				} else {
					if (mouseHover)
						this.that.animState = AnimationType.SCROLL_HOVER;
					else
						this.that.animState = AnimationType.NONE;
				}
			}
		} else {
			if(arrowHover && this.that.settings.showArrows) {
				this.that.lastArrowState = this.that.arrowState;
				this.that.arrowState = arrowHover;
				this.that.animState = AnimationType.ARROW_HOVER;
			} else {
				if (this.that._MouseHoverOnScroller(mousePos)) {
					this.that.animState = AnimationType.SCROLL_HOVER;
				} else {
					this.that.animState = AnimationType.NONE;
				}
			}
			this.that.mouseDownArrow = false;
		}
		this.that._draw();
		//for unlock global mouse event
		if ( this.that.onLockMouse && this.that.offLockMouse ) {
			this.that.offLockMouse( evt );
		}
	};
	ScrollObject.prototype.evt_mousedown = function ( e ) {
		var evt = e || window.event;

		// prevent pointer events on all iframes (while only plugin!)
		if (window.g_asc_plugins)
			window.g_asc_plugins.disablePointerEvents();

		// если сделать превент дефолт - перестанет приходить mousemove от window
		/*
		 if (evt.preventDefault)
		 evt.preventDefault();
		 else
		 evt.returnValue = false;
		 */

		var mousePos = this.that.getMousePosition( evt ),
		    arrowHover = this.that._MouseArrowHover(mousePos);

		this.that.mouseDown = true;

		//arrow pressed
		if (this.that.settings.showArrows && arrowHover) {
			this.that.mouseDownArrow = true;
			this.that.arrowState = arrowHover;
			this.that.animState = AnimationType.ARROW_ACTIVE;
			if (arrowHover === ArrowType.ARROW_TOP || arrowHover === ArrowType.ARROW_LEFT) {
				this.that._arrowUpMouseDown();
			} else if (arrowHover === ArrowType.ARROW_BOTTOM || arrowHover === ArrowType.ARROW_RIGHT) {
				this.that._arrowDownMouseDown();
			}
			this.that._draw();
		} else {
			this.that.mouseUp = false;

			if ( this.that._MouseHoverOnScroller( mousePos ) ) {
				this.that.scrollerMouseUp = false;
				this.that.scrollerMouseDown = true;

				if ( this.that.onLockMouse ) {
					this.that.onLockMouse( evt );
				}
				this.that.StartMousePosition.x = mousePos.x;
				this.that.StartMousePosition.y = mousePos.y;
				this.that.animState = AnimationType.SCROLL_ACTIVE;
				this.that._draw();
			}
			else {
				//scroll pressed, but not slider
				if ( this.that.settings.isVerticalScroll ) {
					var _tmp = this,
						direction = mousePos.y - this.that.scroller.y - this.that.scroller.h / 2,
						step = this.that.paneHeight * this.that.settings.scrollPagePercent,
//                        verticalDragPosition = this.that.scroller.y,
						isFirst = true,
						doScroll = function () {
							_tmp.that.lock = true;
							if ( direction > 0 ) {
								if ( _tmp.that.scroller.y + _tmp.that.scroller.h / 2 + step < mousePos.y ) {
									_tmp.that.scrollByY( step * _tmp.that.scrollCoeff );
								}
								else {
									var _step = Math.abs( _tmp.that.scroller.y + _tmp.that.scroller.h / 2 - mousePos.y );
									_tmp.that.scrollByY( _step * _tmp.that.scrollCoeff );
									cancelClick();
									return;
								}
							}
							else if ( direction < 0 ) {
								if ( _tmp.that.scroller.y + _tmp.that.scroller.h / 2 - step > mousePos.y ) {
									_tmp.that.scrollByY( -step * _tmp.that.scrollCoeff );
								}
								else {
									var _step = Math.abs( _tmp.that.scroller.y + _tmp.that.scroller.h / 2 - mousePos.y );
									_tmp.that.scrollByY( -_step * _tmp.that.scrollCoeff );
									cancelClick();
									return;
								}
							}
							_tmp.that.scrollTimeout = setTimeout( doScroll, isFirst ? _tmp.that.settings.initialDelay : _tmp.that.settings.trackClickRepeatFreq );
							isFirst = false;
						},
						cancelClick = function () {
							_tmp.that.scrollTimeout && clearTimeout( _tmp.that.scrollTimeout );
							_tmp.that.scrollTimeout = null;
							_tmp.that.unbind( "mouseup.main", cancelClick );
							_tmp.that.lock = false;
						};

					if ( this.that.onLockMouse ) {
						this.that.onLockMouse( evt );
					}

					doScroll();
					this.that.bind( "mouseup.main", cancelClick );
				}
				if ( this.that.settings.isHorizontalScroll ) {
					var _tmp = this,
						direction = mousePos.x - this.that.scroller.x - this.that.scroller.w / 2,
						step = this.that.paneWidth * this.that.settings.scrollPagePercent,
//                        horizontalDragPosition = this.that.scroller.x,
						isFirst = true,
						doScroll = function () {
							_tmp.that.lock = true;
							if ( direction > 0 ) {
								if ( _tmp.that.scroller.x + _tmp.that.scroller.w / 2 + step < mousePos.x ) {
									_tmp.that.scrollByX( step * _tmp.that.scrollCoeff );
								}
								else {
									var _step = Math.abs( _tmp.that.scroller.x + _tmp.that.scroller.w / 2 - mousePos.x );
									_tmp.that.scrollByX( _step * _tmp.that.scrollCoeff );
									cancelClick();
									return;
								}
							}
							else if ( direction < 0 ) {
								if ( _tmp.that.scroller.x + _tmp.that.scroller.w / 2 - step > mousePos.x ) {
									_tmp.that.scrollByX( -step * _tmp.that.scrollCoeff );
								}
								else {
									var _step = Math.abs( _tmp.that.scroller.x + _tmp.that.scroller.w / 2 - mousePos.x );
									_tmp.that.scrollByX( -_step * _tmp.that.scrollCoeff );
									cancelClick();
									return;
								}
							}
							_tmp.that.scrollTimeout = setTimeout( doScroll, isFirst ? _tmp.that.settings.initialDelay : _tmp.that.settings.trackClickRepeatFreq );
							isFirst = false;
							// _tmp.that._drawArrow( ArrowStatus.arrowHover );
						},
						cancelClick = function () {
							_tmp.that.scrollTimeout && clearTimeout( _tmp.that.scrollTimeout );
							_tmp.that.scrollTimeout = null;
							_tmp.that.unbind( "mouseup.main", cancelClick );
							_tmp.that.lock = false;
						};

					if ( this.that.onLockMouse ) {
						this.that.onLockMouse( evt );
					}

					doScroll();
					this.that.bind( "mouseup.main", cancelClick );
				}
			}
		}
	};
	ScrollObject.prototype.evt_mousewheel = function ( e ) {
		var evt = e || window.event;
		/* if ( evt.preventDefault )
		 evt.preventDefault();
		 else
		 evt.returnValue = false;*/

		var delta = 1;
		if ( this.that.settings.isHorizontalScroll ) return;
		if ( undefined != evt.wheelDelta )
			delta = (evt.wheelDelta > 0) ? -this.that.settings.vscrollStep : this.that.settings.vscrollStep;
		else
			delta = (evt.detail > 0) ? this.that.settings.vscrollStep : -this.that.settings.vscrollStep;
		delta *= this.that.settings.wheelScrollLines;
		this.that.scroller.y += delta;
		if ( this.that.scroller.y < 0 ) {
			this.that.scroller.y = 0;
		}
		else if ( this.that.scroller.y + this.that.scroller.h > this.that.canvasH ) {
			this.that.scroller.y = this.that.canvasH - this.that.arrowPosition - this.that.scroller.h;
		}
		this.that.scrollByY( delta )
	};
	ScrollObject.prototype.evt_click = function ( e ) {
		var evt = e || window.event;
		var mousePos = this.that.getMousePosition( evt );
		if ( this.that.settings.isHorizontalScroll ) {
			if ( mousePos.x > this.arrowPosition && mousePos.x < this.that.canvasW - this.that.arrowPosition ) {
				if ( this.that.scroller.x > mousePos.x ) {
					this.that.scrollByX( -this.that.settings.vscrollStep );
				}
				if ( this.that.scroller.x < mousePos.x && this.that.scroller.x + this.that.scroller.w > mousePos.x ) {
					return false;
				}
				if ( this.that.scroller.x + this.that.scroller.w < mousePos.x ) {
					this.that.scrollByX( this.that.settings.hscrollStep );
				}
			}
		}
		if ( this.that.settings.isVerticalScroll ) {
			if ( mousePos.y > this.that.arrowPosition && mousePos.y < this.that.canvasH - this.that.arrowPosition ) {
				if ( this.that.scroller.y > mousePos.y ) {
					this.that.scrollByY( -this.that.settings.vscrollStep );
				}
				if ( this.that.scroller.y < mousePos.y && this.that.scroller.y + this.that.scroller.h > mousePos.y ) {
					return false;
				}
				if ( this.that.scroller.y + this.that.scroller.h < mousePos.y ) {
					this.that.scrollByY( this.that.settings.hscrollStep );
				}
			}
		}
	};

	/************************************************************************/
	/*for events*/
	ScrollObject.prototype.bind = function ( typesStr, handler ) {
		var types = typesStr.split( " " );
		/*
		 * loop through types and attach event listeners to
		 * each one.  eg. "click mouseover.namespace mouseout"
		 * will create three event bindings
		 */
		for ( var n = 0; n < types.length; n++ ) {
			var type = types[n];
			var event = (type.indexOf( 'touch' ) == -1) ? 'on' + type : type;
			var parts = event.split( "." );
			var baseEvent = parts[0];
			var name = parts.length > 1 ? parts[1] : "";

			if ( !this.eventListeners[baseEvent] ) {
				this.eventListeners[baseEvent] = [];
			}

			this.eventListeners[baseEvent].push( {
				name:name,
				handler:handler
			} );
		}
	};
	ScrollObject.prototype.unbind = function ( typesStr ) {
		var types = typesStr.split( " " );

		for ( var n = 0; n < types.length; n++ ) {
			var type = types[n];
			var event = (type.indexOf( 'touch' ) == -1) ? 'on' + type : type;
			var parts = event.split( "." );
			var baseEvent = parts[0];

			if ( this.eventListeners[baseEvent] && parts.length > 1 ) {
				var name = parts[1];

				for ( var i = 0; i < this.eventListeners[baseEvent].length; i++ ) {
					if ( this.eventListeners[baseEvent][i].name == name ) {
						this.eventListeners[baseEvent].splice( i, 1 );
						if ( this.eventListeners[baseEvent].length === 0 ) {
							this.eventListeners[baseEvent] = undefined;
						}
						break;
					}
				}
			}
			else {
				this.eventListeners[baseEvent] = undefined;
			}
		}
	};
	ScrollObject.prototype.handleEvents = function ( eventType, evt, p ) {
		var that = this;
		// generic events handler
		function handle( obj ) {
			var el = obj.eventListeners;
			if ( el[eventType] ) {
				var events = el[eventType];
				for ( var i = 0; i < events.length; i++ ) {
					events[i].handler.apply( obj, [evt] );
				}
			}
		}

		/*
		 * simulate bubbling by handling shape events
		 * first, followed by group events, folulowed
		 * by layer events
		 */
		handle( that );
	};

	function _HEXTORGB_( colorHEX ) {
		return {
			R:parseInt( colorHEX.substring( 1, 3 ), 16 ),
			G:parseInt( colorHEX.substring( 3, 5 ), 16 ),
			B:parseInt( colorHEX.substring( 5, 7 ), 16 )
		}
	}

	ScrollObject.prototype._roundForScale = function (value) {
		return ((value - Math.floor(value)) <= 0.5) ? Math.floor(value) : Math.round(value);
	};
    //---------------------------------------------------------export---------------------------------------------------
	window["AscCommon"].ScrollSettings = ScrollSettings;
    window["AscCommon"].ScrollObject = ScrollObject;
})(window);
