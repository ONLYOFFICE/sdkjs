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

(function(window, undefined){

	var darkModeEdge = 10;
	function _darkModeCheckColor(r, g, b)
	{
		return (r < darkModeEdge && g < darkModeEdge && b < darkModeEdge) ? true : false;
	}
	function _darkModeCheckColor2(r, g, b)
	{
		if (r < darkModeEdge && g < darkModeEdge && b < darkModeEdge) return true;
		var max = 255 - darkModeEdge;
		if (r > max && g > max && b > max) return true;
		return false;
	}

	AscCommon.darkModeCorrectColor = function(r, g, b)
	{
		if (r < darkModeEdge && g < darkModeEdge && b < darkModeEdge)
			return { R : 255 - r, G: 255 - g, B : 255 - b };
		return { R : r, G : g, B : b };
	};
	AscCommon.darkModeCorrectColor2 = function(r, g, b)
	{
		var oHSL = {}, oRGB = {};
		AscFormat.CColorModifiers.prototype.RGB2HSL(r, g, b, oHSL);
		var dKoefL = (255 - 58) / 255;
		oHSL.L = 255 - ((dKoefL * oHSL.L) >> 0);
		AscFormat.CColorModifiers.prototype.HSL2RGB(oHSL, oRGB);
		return oRGB;
	};

	AscCommon.RendererType = {
		Base          : 0,
		Drawer        : 1,
		PDF           : 2,
		BoundsChecker : 3,
		Track         : 4,
		TextDrawer    : 5,
		NativeDrawer  : 6
	};

	function CGraphicsBase(type, isUseGrState)
	{
		this.type = (undefined === type) ? AscCommon.RendererType.Base : type;
		this.isDarkMode   = false;

		this.shapeDrawCounter = 0;
		this.isFormDraw       = 0;

		this.globalAlpha = 1;
		this.textAlpha = undefined;

		this.GrState = null;

		// Flags
		this.IsNoDrawingEmptyPlaceholder = false;
		this.IsDemonstrationMode = false;
		this.IsThumbnail = false;
		this.IsDrawSmart = false;
		this.IsPrintMode = false;
		this.IsPrintPreview = false;

		if (true === isUseGrState)
		{
			this.GrState = new AscCommon.CGrState();
			this.GrState.Parent = this;
		}
	}

	// TYPE
	CGraphicsBase.prototype.isBoundsChecker = function()
	{
		return (this.type === AscCommon.RendererType.BoundsChecker);
	};

	CGraphicsBase.prototype.isPdf = function()
	{
		return (this.type === AscCommon.RendererType.PDF);
	};

	CGraphicsBase.prototype.isNativeDrawer = function()
	{
		return (this.type === AscCommon.RendererType.NativeDrawer);
	};

	CGraphicsBase.prototype.isTrack = function()
	{
		return (this.type === AscCommon.RendererType.Track);
	};

	CGraphicsBase.prototype.isTextDrawer = function()
	{
		return (this.type === AscCommon.RendererType.TextDrawer);
	};

	// DARK MODE REGION
	CGraphicsBase.prototype._darkMode1 = function()
	{
		this.isDarkMode = true;
		function _darkColor(_this, _func) {
			return function(r, g, b, a) {
				if (_this.isDarkMode && _darkModeCheckColor(r, g, b))
					_func.call(_this, 255 - r, 255 - g, 255 - b, a);
				else
					_func.call(_this, r, g, b, a);
			};
		}

		this.p_color_old = this.p_color; this.p_color = _darkColor(this, this.p_color_old);
		this.b_color1_old = this.b_color1; this.b_color1 = _darkColor(this, this.b_color1_old);
		this.b_color2_old = this.b_color2; this.b_color2 = _darkColor(this, this.b_color2_old);
	};

	CGraphicsBase.prototype._darkMode2 = function()
	{
		this.isDarkMode = true;
		function _darkColor(_this, _func) {
			return function(r, g, b, a) {
				var isCorrect = _this.isDarkMode;
				if (isCorrect && 0 !== this.shapeDrawCounter)
					if (!(1 === this.shapeDrawCounter && this.isFormDraw)) //форму первого уровня не корректируем
						isCorrect = false;
				if (isCorrect && _darkModeCheckColor2(r, g, b))
					_func.call(_this, 255 - r, 255 - g, 255 - b, a);
				else
					_func.call(_this, r, g, b, a);
			};
		}

		this.p_color_old = this.p_color; this.p_color = _darkColor(this, this.p_color_old);
		this.b_color1_old = this.b_color1; this.b_color1 = _darkColor(this, this.b_color1_old);
		this.b_color2_old = this.b_color2; this.b_color2 = _darkColor(this, this.b_color2_old);
	};

	CGraphicsBase.prototype._darkMode3 = function()
	{
		this.isDarkMode = true;
		function _darkColor(_this, _func) {
			return function(r, g, b, a) {
				var isCorrect = _this.isDarkMode;
				if (isCorrect && 0 !== this.shapeDrawCounter)
					if (!(1 === this.shapeDrawCounter && this.isFormDraw)) //форму первого уровня не корректируем
						isCorrect = false;
				if (isCorrect)
				{
					var c = AscCommon.darkModeCorrectColor2(r, g, b);
					_func.call(_this, c.R, c.G, c.B, a);
				}
				else
				{
					_func.call(_this, r, g, b, a);
				}
			};
		}

		this.p_color_old = this.p_color; this.p_color = _darkColor(this, this.p_color_old);
		this.b_color1_old = this.b_color1; this.b_color1 = _darkColor(this, this.b_color1_old);
		this.b_color2_old = this.b_color2; this.b_color2 = _darkColor(this, this.b_color2_old);
	};

	CGraphicsBase.prototype.setDarkMode = function(mode)
	{
		if (undefined === mode)
			this._darkMode3();
		else
		{
			if (mode === 1)
				this._darkMode1();
			else if (mode === 2)
				this.darkMode2();
			else
				this.darkMode3();
		}
	};

	CGraphicsBase.prototype.StartDrawShape = function(type, isForm)
	{
		this.shapeDrawCounter++;
		this.isFormDraw = isForm;
	};
	CGraphicsBase.prototype.EndDrawShape = function()
	{
		this.isFormDraw = false;
		this.shapeDrawCounter--;
	};

	// GLOBAL ALPHA
	CGraphicsBase.prototype.Start_GlobalAlpha = function()
	{
	};
	CGraphicsBase.prototype.End_GlobalAlpha = function()
	{
	};
	CGraphicsBase.prototype.setEndGlobalAlphaColor = function(r, g, b)
	{
		this.endGlobalAlphaColor =  (undefined === r) ? undefined : { R : r, G : g, B : b };
	};

	CGraphicsBase.prototype.put_GlobalAlpha = function(enable, alpha)
	{
		this.globalAlpha = (false === enable) ? 1 : alpha;
	};

	// TEXT ALPHA (placeholders)
	CGraphicsBase.prototype.setTextGlobalAlpha = function(alpha)
	{
		this.textAlpha = alpha;
	};
	CGraphicsBase.prototype.getTextGlobalAlpha = function()
	{
		return this.textAlpha;
	};
	CGraphicsBase.prototype.resetTextGlobalAlpha = function()
	{
		this.textAlpha = undefined;
	};

	// PEN
	CGraphicsBase.prototype.p_color = function(r, g, b, a)
	{
	};
	CGraphicsBase.prototype.p_width = function(w)
	{
	};
	CGraphicsBase.prototype.p_dash = function(params)
	{
	};

	// BRUSH
	CGraphicsBase.prototype.b_color1 = function(r, g, b, a)
	{
	};
	CGraphicsBase.prototype.b_color2 = function(r, g, b, a)
	{
	};

	// TRANSFORM
	CGraphicsBase.prototype.grStateIsUseBaseTransform = function()
	{
		return true;
	};
	CGraphicsBase.prototype.grStateSaveBaseTransform = function()
	{
		this._m_oBaseTransform = this.m_oBaseTransform;
		this.m_oBaseTransform = null;
	};
	CGraphicsBase.prototype.grStateRestoreBaseTransform = function()
	{
		this.m_oBaseTransform = this._m_oBaseTransform;
		delete this._m_oBaseTransform;
	};

	CGraphicsBase.prototype.reset = function()
	{
	};
	CGraphicsBase.prototype.transform = function(sx,shy,shx,sy,tx,ty)
	{
	};
	CGraphicsBase.prototype.transform3 = function(m, isNeedInvert)
	{
	};
	CGraphicsBase.prototype.CalculateFullTransform = function()
	{
	};

	// PATH
	CGraphicsBase.prototype._s = function()
	{
	};
	CGraphicsBase.prototype._e = function()
	{
	};
	CGraphicsBase.prototype._z = function()
	{
	};
	CGraphicsBase.prototype._m = function(x, y)
	{
	};
	CGraphicsBase.prototype._l = function(x, y)
	{
	};
	CGraphicsBase.prototype._c = function(x1,y1,x2,y2,x3,y3)
	{
	};
	CGraphicsBase.prototype._c2 = function(x1,y1,x2,y2)
	{
	};

	CGraphicsBase.prototype.ds = function()
	{
	};
	CGraphicsBase.prototype.df = function()
	{
	};

	// STATE
	CGraphicsBase.prototype.save = function()
	{
	};
	CGraphicsBase.prototype.restore = function()
	{
	};
	CGraphicsBase.prototype.clip = function()
	{
	};

	CGraphicsBase.prototype.StartClipPath = function()
	{
	};
	CGraphicsBase.prototype.EndClipPath = function()
	{
	};

	// GRSTATE
	CGraphicsBase.prototype.SavePen = function()
	{
		this.GrState && this.GrState.SavePen();
	};
	CGraphicsBase.prototype.RestorePen = function()
	{
		this.GrState && this.GrState.RestorePen();
	};
	CGraphicsBase.prototype.SaveBrush = function()
	{
		this.GrState && this.GrState.SaveBrush();
	};
	CGraphicsBase.prototype.RestoreBrush = function()
	{
		this.GrState && this.GrState.RestoreBrush();
	};
	CGraphicsBase.prototype.SavePenBrush = function()
	{
		this.GrState && this.GrState.SavePenBrush();
	};
	CGraphicsBase.prototype.RestorePenBrush = function()
	{
		this.GrState && this.GrState.RestorePenBrush();
	};
	CGraphicsBase.prototype.SaveGrState = function()
	{
		this.GrState && this.GrState.SaveGrState();
	};
	CGraphicsBase.prototype.RestoreGrState = function()
	{
		this.GrState && this.GrState.RestoreGrState();
	};
	CGraphicsBase.prototype.RemoveLastClip = function()
	{
		this.GrState && this.GrState.RemoveLastClip();
	};
	CGraphicsBase.prototype.RestoreLastClip = function()
	{
		this.GrState && this.GrState.RestoreLastClip();
	};

	CGraphicsBase.prototype.AddClipRect = function(x, y, w, h)
	{
		if (!this.GrState)
			return;
		var rect = new AscCommon._rect();
		rect.x = x;
		rect.y = y;
		rect.w = w;
		rect.h = h;
		this.GrState.AddClipRect(rect);
	};
	CGraphicsBase.prototype.RemoveClipRect = function()
	{
		// NOT USED (remove & add, add...)
	};

	CGraphicsBase.prototype.SetClip = function(r)
	{
	};
	CGraphicsBase.prototype.RemoveClip = function()
	{
	};

	// EDITOR
	CGraphicsBase.prototype.isSupportEditFeatures = function()
	{
		return false;
	};

	CGraphicsBase.prototype.DrawHeaderEdit = function(yPos, lock_type, sectionNum, bIsRepeat, type)
	{
	};
	CGraphicsBase.prototype.DrawFooterEdit = function(yPos, lock_type, sectionNum, bIsRepeat, type)
	{
	};

	CGraphicsBase.prototype.DrawLockParagraph = function(lock_type, x, y1, y2)
	{
	};
	CGraphicsBase.prototype.DrawLockObjectRect = function(lock_type, x, y, w, h)
	{
	};
	CGraphicsBase.prototype.DrawEmptyTableLine = function(x1,y1,x2,y2)
	{
	};
	CGraphicsBase.prototype.DrawSpellingLine = function(y0, x0, x1, w)
	{
	};

	CGraphicsBase.prototype.drawCollaborativeChanges = function(x, y, w, h, Color)
	{
	};

	CGraphicsBase.prototype.drawMailMergeField = function(x, y, w, h)
	{
	};

	CGraphicsBase.prototype.drawSearchResult = function(x, y, w, h)
	{
	};

	CGraphicsBase.prototype.drawFlowAnchor = function(x, y)
	{
	};

	CGraphicsBase.prototype.DrawFootnoteRect = function(x, y, w, h)
	{
	};

	CGraphicsBase.prototype.DrawPresentationComment = function(type, x, y, w, h)
	{
	};

	// INTEGER GRID
	CGraphicsBase.prototype.SetIntegerGrid = function(param)
	{
	};
	CGraphicsBase.prototype.GetIntegerGrid = function()
	{
		return false;
	};

	// COMMON FUNCS
	CGraphicsBase.prototype.rect = function(x,y,w,h)
	{
		let r = (x + w);
		let b = (y + h);

		this._s();
		this._m(x, y);
		this._l(r, y);
		this._l(r, b);
		this._l(x, b);
		this._z();
	};

	CGraphicsBase.prototype.TableRect = function(x,y,w,h)
	{
		this.rect(x,y,w,h);
		this.df();
	};

	CGraphicsBase.prototype.drawHorLine = function(align, y, x, r, penW)
	{
		this.p_width(1000 * penW);
		this._s();

		var _y = y;
		switch (align)
		{
			case 0:
			{
				_y = y + penW / 2;
				break;
			}
			case 1:
			{
				break;
			}
			case 2:
			{
				_y = y - penW / 2;
			}
		}
		this._m(x, y);
		this._l(r, y);

		this.ds();

		this._e();
	};

	CGraphicsBase.prototype.drawHorLine2 = function(align, y, x, r, penW)
	{
		this.p_width(1000 * penW);

		var _y = y;
		switch (align)
		{
			case 0:
			{
				_y = y + penW / 2;
				break;
			}
			case 1:
			{
				break;
			}
			case 2:
			{
				_y = y - penW / 2;
				break;
			}
		}

		this._s();
		this._m(x, (_y - penW));
		this._l(r, (_y - penW));
		this.ds();

		this._s();
		this._m(x, (_y + penW));
		this._l(r, (_y + penW));
		this.ds();

		this._e();
	};

	CGraphicsBase.prototype.drawVerLine = function(align, x, y, b, penW)
	{
		this.p_width(1000 * penW);
		this._s();

		var _x = x;
		switch (align)
		{
			case 0:
			{
				_x = x + penW / 2;
				break;
			}
			case 1:
			{
				break;
			}
			case 2:
			{
				_x = x - penW / 2;
			}
		}
		this._m(_x, y);
		this._l(_x, b);

		this.ds();
	};

	CGraphicsBase.prototype.drawHorLineExt = function(align, y, x, r, penW, leftMW, rightMW)
	{
		this.drawHorLine(align, y, x + leftMW, r + rightMW, penW);
	};

	CGraphicsBase.prototype.drawPolygon = function(oPath, lineWidth, shift)
	{
		this.p_width(lineWidth * 1000);
		this._s();

		let Points = oPath.Points;
		let nCount = Points.length;
		// берем предпоследнюю точку, т.к. последняя совпадает с первой
		let PrevX = Points[nCount - 2].X, PrevY = Points[nCount - 2].Y;
		let _x    = Points[nCount - 2].X,    _y = Points[nCount - 2].Y;
		let StartX, StartY;

		for (var nIndex = 0; nIndex < nCount; nIndex++)
		{
			if(PrevX > Points[nIndex].X)
			{
				_y = Points[nIndex].Y - shift;
			}
			else if(PrevX < Points[nIndex].X)
			{
				_y = Points[nIndex].Y + shift;
			}

			if(PrevY < Points[nIndex].Y)
			{
				_x = Points[nIndex].X - shift;
			}
			else if(PrevY > Points[nIndex].Y)
			{
				_x = Points[nIndex].X + shift;
			}

			PrevX = Points[nIndex].X;
			PrevY = Points[nIndex].Y;

			if(nIndex > 0)
			{
				if (1 === nIndex)
				{
					StartX = _x;
					StartY = _y;
					this._m(_x, _y);
				}
				else
				{
					this._l(_x, _y);
				}
			}
		}

		this._l(StartX, StartY);
		this._z();
		this.ds();
		this._e();
	};
	
	CGraphicsBase.prototype.drawPolygonByRects = function(rects, lineWidth, shift)
	{
		let polygon = new AscCommon.CPolygon();
		polygon.fill(rects);
		let paths = polygon.GetPaths(shift ? shift : 0);
		for (let i = 0, count = paths.length; i < count; ++i)
		{
			this.drawPolygon(paths[i], lineWidth, 0);
		}
	};

	/**
	 * made with the use of:
	 * http://html5tutorial.com/how-to-draw-n-grade-bezier-curve-with-canvas-api/
	 * uses de Casteljau's algorithm
	 * @param {{x: Number, y: Number, z? :Number}} startPoint
	 * @param {{x: Number, y: Number, z? :Number}[]} controlPoints
	 * @param {{x: Number, y: Number, z? :Number}} endPoint
	 */
	CGraphicsBase.prototype._cN = function(startPoint, controlPoints, endPoint)
	{
		function sumDistanceBetweenPoints(points)
		{
			function distance(a, b){
				return Math.sqrt(Math.pow(a.x-b.x, 2) + Math.pow(a.y-b.y, 2));
			}
			/**Compute the incremental step*/
			let tLength = 0;
			for(let i=0; i < points.length - 1; i++){
				tLength += distance(points[i], points[i+1]);
			}
			return tLength;
		}

		function transformPoints(points, dpi)
		{
			let pointsCopy = Array(points.length);
			let mmToIch = 0.03937007874;
			for(let i=0; i < pointsCopy.length; i++){
				pointsCopy[i] = {x: null, y: null};
				// pointsCopy[i].x = transform.TransformPointX(points[i].x, points[i].y);
				// pointsCopy[i].y = transform.TransformPointY(points[i].x, points[i].y);
				pointsCopy[i].x = points[i].x * mmToIch * dpi;
				pointsCopy[i].y = points[i].y * mmToIch * dpi;
			}
			return pointsCopy;
		}

		/** Computes a point's coordinates for a value of t
		 * @param {Number} t - a value between o and 1
		 * @param {{x: Number, y: Number, z? :Number}} startPoint
		 * @param {{x: Number, y: Number, z? :Number}[]} controlPoints
		 * @param {{x: Number, y: Number, z? :Number}} endPoint
		 * @return {{x: Number, y: Number}} point
		 **/
		function computeBezierPoint(t, startPoint, controlPoints, endPoint)
		{
			/**Computes Bernstain
			 *@param {Integer} i - the i-th index
			 *@param {Integer} n - the total number of points
			 *@param {Number} t - the value of parameter t , between 0 and 1
			 **/
			function computeBernstainCoef(i,n,t)
			{
				/**Computes factorial*/
				function fact(k){
					if(k==0 || k==1){
						return 1;
					}
					else{
						return k * fact(k-1);
					}
				}

				//if(n < i) throw "Wrong";
				return fact(n) / (fact(i) * fact(n-i))* Math.pow(t, i) * Math.pow(1-t, n-i);
			}

			var points = [].concat(startPoint, controlPoints, endPoint);
			var r = {
				x: 0,
				y: 0
			};
			var n = points.length-1;
			for(var i=0; i <= n; i++){
				r.x += points[i].x * computeBernstainCoef(i, n, t);
				r.y += points[i].y * computeBernstainCoef(i, n, t);
			}
			return r;
		}

		// to calculate points count curve length should be calculated in pixels (approximately)
		// so arguments should be in mm units and graphics.transform should be applied already
		let calculatePointsCount = false;

		let interpolationPointsCount;
		if (calculatePointsCount) {
			// https://www.figma.com/file/FT0m9czNuvK34TK227cQ6e/pointsToCalculatePerOnePixelLengthUnit?type=design&node-id=0%3A1&mode=design&t=0S7e2nxkt2sbCHqw-1
			// not integer more like precision coefficient. can be 0.3 for example
			let pointsToCalculatePerOnePixelLengthUnit = 1;

			let dpi = 96;

			let bezierPoints = [].concat(startPoint, controlPoints, endPoint);

			// this.m_oFullTransform doesn't exist in that Graphics
			// convert length to pixels length units
			// https://www.figma.com/file/FT0m9czNuvK34TK227cQ6e/pointsToCalculatePerOnePixelLengthUnit?type=design&node-id=41-49&mode=design&t=pH5bF1EvWeWjzkS0-0
			// Canvas resize is not considered!
			let bezierPointsCopy = transformPoints(bezierPoints, dpi);

			// As we calculate length of a curve as sum of control points there might be performance issue with
			// high order curves because they have many control points and so length will be considered
			// as long and there will be many control points
			// add + 1 to avoid divide by 0 later when calculating interpolation step which is 1/interpolationPointsCount
			interpolationPointsCount = pointsToCalculatePerOnePixelLengthUnit *
				sumDistanceBetweenPoints(bezierPointsCopy) + 1;
		} else {
			interpolationPointsCount = 1000;
		}

		// this.p_width(lineWidth);
		// this._s(); // beginPath
		// this._m(startPoint.x, startPoint.y);

		// in fact real pointsCount is larger by 1 point. bcs if pointsCount = 2 t will be 0, 1/2 and 1 - 3 times total
		for (let t = 0; t <= 1; t+= 1/interpolationPointsCount) {
			let point = computeBezierPoint(t, startPoint, controlPoints, endPoint);
			this._l(point.x, point.y);
		}

		// https://github.com/ONLYOFFICE/sdkjs/blob/ebcb7401438a8260151cd96f7568d521e04f91e9/word/Drawing/Graphics.js#L438
		// this._z(); // close path
		// this.ds(); // draw stroke
		// this._e(); // beginPath
	};

	// FONT
	CGraphicsBase.prototype.FreeFont = function()
	{
	};
	CGraphicsBase.prototype.ClearLastFont = function()
	{
	};

	CGraphicsBase.prototype.GetFont = function()
	{
	};
	CGraphicsBase.prototype.SetFont = function(font)
	{
	};

	CGraphicsBase.prototype.GetTextPr = function()
	{
	};
	CGraphicsBase.prototype.SetTextPr = function(textPr, theme)
	{
	};

	CGraphicsBase.prototype.SetFontSlot = function(slot, fontSizeKoef)
	{
	};
	CGraphicsBase.prototype.SetFontInternal = function(name, size, style)
	{
	};

	// TEXT
	CGraphicsBase.prototype.isSupportTextDraw = function()
	{
		return true;
	};
	CGraphicsBase.prototype.isSupportTextOutline = function()
	{
		return false;
	};

	CGraphicsBase.prototype.FillText = function(x,y,text)
	{
	};

	CGraphicsBase.prototype.t = function(text,x,y,isBounds)
	{
	};

	CGraphicsBase.prototype.FillText2 = function(x,y,text,cropX,cropW)
	{
	};

	CGraphicsBase.prototype.t2 = function(text,x,y,cropX,cropW)
	{
	};

	CGraphicsBase.prototype.FillTextCode = function(x,y,lUnicode)
	{
	};

	CGraphicsBase.prototype.tg = function(text,x,y,codepoints)
	{
	};

	// IMAGES
	CGraphicsBase.prototype.isVectorImage = function(img)
	{
		if (img.isVectorImage !== undefined)
			return img.isVectorImage;
		if (!img.src)
			return false;
		let fileName = AscCommon.g_oDocumentUrls.getImageLocal(img.src);
		img.isVectorImage = (fileName && fileName.endsWith(".svg")) ? true : false;
		return img.isVectorImage;
	};

	CGraphicsBase.prototype.drawImage2 = function(img,x,y,w,h,alpha,srcRect)
	{
	};

	CGraphicsBase.prototype.drawImage = function(img,x,y,w,h,alpha,srcRect,nativeImage)
	{
	};

	// COMMANDS
	CGraphicsBase.prototype.Start_Command = function(commandId)
	{
	};
	CGraphicsBase.prototype.End_Command = function(commandId)
	{
	};
	
	CGraphicsBase.prototype.SetBorder = function(border)
	{
	};

	//------------------------------------------------------------export----------------------------------------------------
	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].CGraphicsBase = CGraphicsBase;
})(window);
