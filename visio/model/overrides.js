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

"use strict";

// Import
// var CShape = AscFormat.CShape;
// var CGroupShape = AscFormat.CGroupShape;
// var CTheme = AscFormat.CTheme;
// var CreateSolidFillRGBA = AscFormat.CreateSolidFillRGBA;
// var CShapeDrawer = AscCommon.CShapeDrawer;
// var DrawLineEnd = AscCommon.DrawLineEnd;
// var builder_CreateLine = AscFormat.builder_CreateLine;

/**
 * @memberOf AscFormat.CShape
 * @return {{layout: null, slide: null, theme: CTheme, master: null}}
 */
AscFormat.CShape.prototype.getParentObjects = function ()
{
	let oTheme = null;
	if (this.parent) {
		oTheme = this.parent.theme;
	} else {
		AscCommon.consoleLog("Parent was not set for shape/group. GenerateDefaultTheme is used. shape/group:", this);
		oTheme = AscFormat.GenerateDefaultTheme(null, null);
	}
	return {slide: null, layout: null, master: null, theme: oTheme};
};

/**
 * @memberOf AscFormat.CGroupShape
 * @type {function(): {layout: null, slide: null, theme: CTheme, master: null}}
 */
AscFormat.CGroupShape.prototype.getParentObjects = CShape.prototype.getParentObjects;


/**
 * Draw editor.
 * @memberof AscFormat.CShape
 */
AscFormat.CShape.prototype.recalculate = function ()
{
	if(this.bDeleted || !this.parent) {
		AscCommon.consoleLog("no recalculate for bDeleted or no parent");
		return;
	}

	if(this.parent.getObjectType() === AscDFH.historyitem_type_Notes){
		return;
	}

	// var check_slide_placeholder = !this.isPlaceholder() || (this.parent && (this.parent.getObjectType() === AscDFH.historyitem_type_Slide));
	let check_placeholder = !this.isPlaceholder() || (this.parent && this.parent.isVisioDocument);
	AscFormat.ExecuteNoHistory(function(){

		var bRecalcShadow = this.recalcInfo.recalculateBrush ||
			this.recalcInfo.recalculatePen ||
			this.recalcInfo.recalculateTransform ||
			this.recalcInfo.recalculateGeometry ||
			this.recalcInfo.recalculateBounds;
		if (this.recalcInfo.recalculateBrush) {
			this.recalculateBrush();
			this.recalcInfo.recalculateBrush = false;
		}

		if (this.recalcInfo.recalculatePen) {
			this.recalculatePen();
			this.recalcInfo.recalculatePen = false;
		}
		if (this.recalcInfo.recalculateTransform) {
			this.recalculateTransform();
			this.recalculateSnapArrays();
			this.recalcInfo.recalculateTransform = false;
		}

		if (this.recalcInfo.recalculateGeometry) {
			this.recalculateGeometry();
			this.recalcInfo.recalculateGeometry = false;
		}

		if (this.recalcInfo.recalculateContent && check_placeholder) {
			this.recalcInfo.oContentMetrics = this.recalculateContent();
			this.recalcInfo.recalculateContent = false;
		}
		if (this.recalcInfo.recalculateContent2 && check_placeholder) {
			this.recalculateContent2();
			this.recalcInfo.recalculateContent2 = false;
		}

		if (this.recalcInfo.recalculateTransformText && check_placeholder) {
			this.recalculateTransformText();
			this.recalcInfo.recalculateTransformText = false;
		}
		if(this.recalcInfo.recalculateBounds)
		{
			this.recalculateBounds();
			this.recalcInfo.recalculateBounds = false;
		}
		if(bRecalcShadow)
		{
			this.recalculateShdw();
		}

		this.clearCropObject();
	}, this, []);
};

/**
 *
 * @param idx
 * @param unicolor
 * @param {Boolean} isConnectorShape
 * @memberOf AscFormat.CTheme
 * @return {CUniFill|*}
 */
AscFormat.CTheme.prototype.getFillStyle = function (idx, unicolor, isConnectorShape) {
	if (idx === 0 || idx === 1000) {
		return AscFormat.CreateNoFillUniFill();
	}
	var ret;
	let fmtScheme = (isConnectorShape && this.themeElements.themeExt) ?
		this.themeElements.themeExt.fmtConnectorScheme :
		this.themeElements.fmtScheme;
	if (idx >= 1 && idx <= 999) {
		if (fmtScheme.fillStyleLst[idx - 1]) {
			ret = fmtScheme.fillStyleLst[idx - 1].createDuplicate();
			if (ret) {
				ret.checkPhColor(unicolor, false);
				return ret;
			}
		}
	} else if (idx >= 1001) {
		if (fmtScheme.bgFillStyleLst[idx - 1001]) {
			ret = fmtScheme.bgFillStyleLst[idx - 1001].createDuplicate();
			if (ret) {
				ret.checkPhColor(unicolor, false);
				return ret;
			}
		}
	}
	AscCommon.consoleLog("getFillStyle has not found fill and returned transparent fill")
	return AscFormat.CreateSolidFillRGBA(0, 0, 0, 255);
};

/**
 *
 * @param idx
 * @param unicolor
 * @param {Boolean} isConnectorShape
 * @memberOf AscFormat.CTheme
 * @return {CLn|*}
 */
AscFormat.CTheme.prototype.getLnStyle = function (idx, unicolor, isConnectorShape) {
	if (idx === 0) {
		return AscFormat.CreateNoFillLine();
	}
	let fmtScheme = (isConnectorShape && this.themeElements.themeExt) ?
		this.themeElements.themeExt.fmtConnectorScheme :
		this.themeElements.fmtScheme;
	if (fmtScheme.lnStyleLst[idx - 1]) {
		var ret = fmtScheme.lnStyleLst[idx - 1].createDuplicate();
		if (ret.Fill) {
			ret.Fill.checkPhColor(unicolor, false);
		}
		return ret;
	}
	AscCommon.consoleLog("getLnStyle has not found lineStyle and returned new CLn()");
	return new AscFormat.CLn();
};


/**
 *
 * @param {number} nWidth - emus
 * @param oFill
 * @return {CLn}
 */
AscFormat.builder_CreateLine = function(nWidth, oFill) {
	if (nWidth === 0) {
		// return new AscFormat.CreateNoFillLine();
		nWidth = 1000;
	}
	var oLn = new AscFormat.CLn();
	oLn.w = nWidth;
	oLn.Fill = oFill.UniFill;
	return oLn;
}

/**
 * @memberof AscCommon.CShapeDrawer
 */
AscCommon.CShapeDrawer.prototype.ds = function()
{
	if (this.bIsNoStrokeAttack)
		return;

	if (this.Graphics.isTrack())
		this.Graphics.m_oOverlay.ClearAll = true;

	if (null != this.OldLineJoin && !this.IsArrowsDrawing)
	{
		switch (this.Ln.Join.type)
		{
			case AscFormat.LineJoinType.Round:
			{
				this.Graphics.m_oContext.lineJoin = "round";
				break;
			}
			case AscFormat.LineJoinType.Bevel:
			{
				this.Graphics.m_oContext.lineJoin = "bevel";
				break;
			}
			case AscFormat.LineJoinType.Empty:
			{
				this.Graphics.m_oContext.lineJoin = "miter";
				break;
			}
			case AscFormat.LineJoinType.Miter:
			{
				this.Graphics.m_oContext.lineJoin = "miter";
				break;
			}
		}
	}

	var arr = this.Graphics.isTrack() ? this.Graphics.Graphics.ArrayPoints : this.Graphics.ArrayPoints;
	var isArrowsPresent = (arr != null && arr.length > 1 && this.IsCurrentPathCanArrows === true) ? true : false;

	var rgba = this.StrokeUniColor;
	let nAlpha = 0xFF;
	if(!isArrowsPresent && !this.IsArrowsDrawing || Asc.editor.isPdfEditor())
	{
		if (this.Ln && this.Ln.Fill != null && this.Ln.Fill.transparent != null)
			nAlpha = this.Ln.Fill.transparent;
	}

	this.Graphics.p_color(rgba.R, rgba.G, rgba.B, nAlpha);

	if (this.IsRectShape && this.Graphics.AddSmartRect !== undefined)
	{
		if (undefined !== this.Shape.extX)
			this.Graphics.AddSmartRect(0, 0, this.Shape.extX, this.Shape.extY, this.StrokeWidth);
		else
			this.Graphics.ds();
	}
	else
	{
		this.Graphics.ds();
	}

	if (null != this.OldLineJoin && !this.IsArrowsDrawing)
	{
		this.Graphics.m_oContext.lineJoin = this.OldLineJoin;
	}

	if (isArrowsPresent)
	{
		this.IsArrowsDrawing = true;
		this.Graphics.p_dash(null);
		// значит стрелки есть. теперь:
		// определяем толщину линии "как есть"
		// трансформируем точки в окончательные.
		// и отправляем на отрисовку (с матрицей)

		var _graphicsCtx = this.Graphics.isTrack() ? this.Graphics.Graphics : this.Graphics;

		var trans = _graphicsCtx.m_oFullTransform;
		let originalTrans = new AscCommon.CMatrix();
		originalTrans.CopyFrom(trans);

		trans.sx = 1;
		trans.sy = 1;
		trans.shx = 0;
		trans.shy = 0;
		// arrowTrans.SetValues(1, 0, 0, 1, trans.tx, trans.ty);

		var trans1 = AscCommon.global_MatrixTransformer.Invert(trans);

		var x1 = originalTrans.TransformPointX(0, 0);
		var y1 = originalTrans.TransformPointY(0, 0);
		var x2 = originalTrans.TransformPointX(1, 1);
		var y2 = originalTrans.TransformPointY(1, 1);
		var dKoef = Math.sqrt(((x2-x1)*(x2-x1)+(y2-y1)*(y2-y1))/2);
		// var dKoef = 1;
		var _pen_w = this.Graphics.isTrack() ? (this.Graphics.Graphics.m_oContext.lineWidth /* * dKoef*/) : (this.Graphics.m_oContext.lineWidth  /* * dKoef*/);
		var _max_w = undefined;
		if (_graphicsCtx.IsThumbnail === true)
			_max_w = 2;

		var _max_delta_eps2 = 0.001;

		// var arrKoef = this.isArrPix ? (1 / AscCommon.g_dKoef_mm_to_pix) : 1;
		var arrKoef = 1;

		if (this.Ln.headEnd != null)
		{
			var _x1 = trans.TransformPointX(arr[0].x, arr[0].y);
			var _y1 = trans.TransformPointY(arr[0].x, arr[0].y);
			var _x2 = trans.TransformPointX(arr[1].x, arr[1].y);
			var _y2 = trans.TransformPointY(arr[1].x, arr[1].y);

			var _x1Orig = originalTrans.TransformPointX(arr[0].x, arr[0].y);
			var _y1Orig = originalTrans.TransformPointY(arr[0].x, arr[0].y);
			var _x2Orig = originalTrans.TransformPointX(arr[1].x, arr[1].y);
			var _y2Orig = originalTrans.TransformPointY(arr[1].x, arr[1].y);

			var _max_delta_eps = Math.max(this.Ln.headEnd.GetLen(_pen_w) * dKoef, 5);

			var _max_delta = Math.max(Math.abs(_x1Orig - _x2Orig), Math.abs(_y1Orig - _y2Orig));
			var cur_point = 2;
			while (_max_delta < _max_delta_eps && cur_point < arr.length)
			{
				_x2 = trans.TransformPointX(arr[cur_point].x, arr[cur_point].y);
				_y2 = trans.TransformPointY(arr[cur_point].x, arr[cur_point].y);

				_x2Orig = originalTrans.TransformPointX(arr[cur_point].x, arr[cur_point].y);
				_y2Orig = originalTrans.TransformPointY(arr[cur_point].x, arr[cur_point].y);

				_max_delta = Math.max(Math.abs(_x1Orig - _x2Orig), Math.abs(_y1Orig - _y2Orig));
				cur_point++;
			}

			if (_max_delta > _max_delta_eps2)
			{
				_graphicsCtx.ArrayPoints = null;
				AscCommon.DrawLineEnd(_x1, _y1, _x2, _y2, this.Ln.headEnd.type, arrKoef * this.Ln.headEnd.GetWidth(_pen_w, _max_w), arrKoef * this.Ln.headEnd.GetLen(_pen_w, _max_w), this, trans1);
				_graphicsCtx.ArrayPoints = arr;
			}
		}
		if (this.Ln.tailEnd != null)
		{
			var _1 = arr.length-1;
			var _2 = arr.length-2;
			var _x1 = trans.TransformPointX(arr[_1].x, arr[_1].y);
			var _y1 = trans.TransformPointY(arr[_1].x, arr[_1].y);
			var _x2 = trans.TransformPointX(arr[_2].x, arr[_2].y);
			var _y2 = trans.TransformPointY(arr[_2].x, arr[_2].y);

			var _x1Orig = originalTrans.TransformPointX(arr[_1].x, arr[_1].y);
			var _y1Orig = originalTrans.TransformPointY(arr[_1].x, arr[_1].y);
			var _x2Orig = originalTrans.TransformPointX(arr[_2].x, arr[_2].y);
			var _y2Orig = originalTrans.TransformPointY(arr[_2].x, arr[_2].y);

			var _max_delta_eps = Math.max(this.Ln.tailEnd.GetLen(_pen_w) * dKoef, 5);

			var _max_delta = Math.max(Math.abs(_x1Orig - _x2Orig), Math.abs(_y1Orig - _y2Orig));
			var cur_point = _2 - 1;
			while (_max_delta < _max_delta_eps && cur_point >= 0)
			{
				_x2 = trans.TransformPointX(arr[cur_point].x, arr[cur_point].y);
				_y2 = trans.TransformPointY(arr[cur_point].x, arr[cur_point].y);

				_x2Orig = originalTrans.TransformPointX(arr[cur_point].x, arr[cur_point].y);
				_y2Orig = originalTrans.TransformPointY(arr[cur_point].x, arr[cur_point].y);

				_max_delta = Math.max(Math.abs(_x1Orig - _x2Orig), Math.abs(_y1Orig - _y2Orig));
				cur_point--;
			}

			if (_max_delta > _max_delta_eps2)
			{
				_graphicsCtx.ArrayPoints = null;
				AscCommon.DrawLineEnd(_x1, _y1, _x2, _y2, this.Ln.tailEnd.type, arrKoef * this.Ln.tailEnd.GetWidth(_pen_w, _max_w), arrKoef * this.Ln.tailEnd.GetLen(_pen_w, _max_w), this, trans1);
				_graphicsCtx.ArrayPoints = arr;
			}
		}
		this.IsArrowsDrawing = false;
		this.CheckDash();
	}
}

function parseFieldPictureFormat(vsdxFieldValue, vsdxFieldFormat) {
	let res = "@";
	if (vsdxFieldFormat.f) {
		let formatFunction = vsdxFieldFormat.f.toUpperCase();
		let vFieldPicture = parseInt(formatFunction.substring('FIELDPICTURE('.length));

		if (0 === vFieldPicture) {
			res = "General";
		} else if (1 === vFieldPicture) {
			res = "General u";//"General";
		} else if (2 === vFieldPicture) {
			res = "0";
		} else if (3 === vFieldPicture) {
			res = "0 u";
		} else if (4 === vFieldPicture) {
			res = "0.0";
		} else if (5 === vFieldPicture) {
			res = "0.0 u";
		} else if (6 === vFieldPicture) {
			res = "0.00";
		} else if (7 === vFieldPicture) {
			res = "0.00 u";
		} else if (8 === vFieldPicture) {
			res = "0.000";
		} else if (9 === vFieldPicture) {
			res = "0.000 u";
		} else if (9 === vFieldPicture) {
			res = "0.000 u";
		} else if (10 === vFieldPicture) {
			res = "<,FEET/INCH>0.000 u";
		} else if (11 === vFieldPicture) {
			res = "<,rad>0.#### u";
		} else if (12 === vFieldPicture) {
			res = "<,deg>0.# u";
		} else if (13 === vFieldPicture) {
			res = "<,FEET/INCH># #/# u";
		} else if (14 === vFieldPicture) {
			res = "<,FEET/INCH># #/## u";
		} else if (15 === vFieldPicture) {
			res = "0 #/#";
		} else if (16 === vFieldPicture) {
			res = "0 #/# u";
		} else if (17 === vFieldPicture) {
			res = "0 #/##";
		} else if (18 === vFieldPicture) {
			res = "0 #/## u";
		} else if (20 === vFieldPicture) {
			res = "ddddd";
		} else if (21 === vFieldPicture) {
			res = "dddddd";
		} else if (23 === vFieldPicture) {
			res = "MM/dd/yy";
		} else if (24 === vFieldPicture) {
			res = "MMM d, yyyy";
		} else if (25 === vFieldPicture) {
			res = "MMMM d, yyyy";
		} else if (26 === vFieldPicture) {
			res = "d/M/YY";
		} else if (27 === vFieldPicture) {
			res = "dd/MM/yy";
		} else if (28 === vFieldPicture) {
			res = "d MMM, yyyy";
		} else if (29 === vFieldPicture) {
			res = "d MMMM, yyyy";
		} else if (30 === vFieldPicture) {
			res = "T";
		} else if (31 === vFieldPicture) {
			res = "h:mm";
		} else if (32 === vFieldPicture) {
			res = "hh:mm";
		} else if (33 === vFieldPicture) {
			res = "H:mm";
		} else if (34 === vFieldPicture) {
			res = "HH:mm";
		} else if (35 === vFieldPicture) {
			res = "h:mm tt";
		} else if (36 === vFieldPicture) {
			res = "Hh:mm tt";
		} else if (37 === vFieldPicture) {
			res = "@";
		} else if (38 === vFieldPicture) {
			res = "@-";
		} else if (39 === vFieldPicture) {
			res = "@+";
		} else if (40 <= vFieldPicture && vFieldPicture <= 81) {
			res = "M/d/yyyy";
		}
	} else if (vsdxFieldFormat.v) {
		res = vsdxFieldFormat.v;
	}
	return res;
}

AscCommonWord.CPresentationField.prototype.private_GetDateTimeFormat = function(vsdxFieldValue, vsdxFieldFormat)
{
	function getLanguageDependantFormat(aFormat, vsdxFieldFormat) {
		let defaultResult = "@";
		let aFormatIndex = null;
		let result;
		if (vsdxFieldFormat.f) {
			let formatFunction = vsdxFieldFormat.f.toUpperCase();
			let vFieldPicture = parseInt(formatFunction.substring('FIELDPICTURE('.length));
			switch (vFieldPicture) {
				case 0:
					// defaultResult = "General";
					defaultResult = "dd.MM.yyyy";
					// aFormatIndex = 0; // not found for default - english
					break;

				case 30:
					defaultResult = "T";
					break;
				case 31:
					defaultResult = "h:mm";
					break;
				case 32:
					defaultResult = "hh:mm";
					break;
				case 33:
					defaultResult = "H:mm";
					break;
				case 34:
					defaultResult = "HH:mm";
					break;
				case 35:
					defaultResult = "h:mm tt";
					break;
				case 36:
					defaultResult = "Hh:mm tt";
					break;

				case 37:
					// defaultResult = "@";
					defaultResult = "dd.MM.yyyy";
					// aFormatIndex = 0;
					break;
				case 200:
					defaultResult = "M/d/yyyy";
					aFormatIndex = 0;
					break;
				case 201:
					defaultResult = "dddd, MMMM d, yyyy";
					aFormatIndex = 1;
					break;
				case 202:
					defaultResult = "MMMM d, yyyy";
					aFormatIndex = 2;
					break;
				case 203:
					defaultResult = "M/d/yy";
					aFormatIndex = 3;
					break;
				case 204:
					defaultResult = "yyyy-MM-dd";
					aFormatIndex = 4;
					break;
				case 205:
					defaultResult = "d-MMM-yy";
					aFormatIndex = 5;
					break;
				case 206:
					defaultResult = "M.d.yyyy";
					aFormatIndex = 6;
					break;
				case 207:
					defaultResult = "MMM. d, yy";
					aFormatIndex = 7;
					break;
				case 208:
					defaultResult = "d MMMM yyyy";
					aFormatIndex = 8;
					break;
				case 209:
					defaultResult = "MMMM yy";
					aFormatIndex = 9;
					break;
				case 210:
					defaultResult = "MMM-yy";
					aFormatIndex = 10;
					break;
				case 211:
					defaultResult = "M/d/yyyy h:mm am/pm";
					aFormatIndex = 11; // but "M/d/yyyy h:mm AM/PM" in formats
					break;
				case 212:
					defaultResult = "M/d/yyyy h:mm:ss am/pm";
					aFormatIndex = 12; // but "M/d/yyyy h:mm:ss AM/PM" in formats
					break;
				case 213:
					defaultResult = "h:mm am/pm";
					aFormatIndex = 13; // but "h:mm AM/PM" in formats
					break;
				case 214:
					defaultResult = "h:mm:ss am/pm";
					aFormatIndex = 14; // but "h:mm:ss AM/PM" in formats
					break;
				case 215:
					defaultResult = "HH:mm";
					aFormatIndex = 15;
					break;
				case 216:
					defaultResult = "HH:mm:ss";
					aFormatIndex = 16;
					break;
			}
		} else if (vsdxFieldFormat.v) {
			defaultResult = vsdxFieldFormat.v;
		}

		if (Array.isArray(aFormat)) {
			// get language dependant format string if index is set
			result = aFormatIndex !== null ? aFormat[aFormatIndex] : defaultResult;
		} else {
			// use default value
			result = defaultResult;
		}

		return result;

	}

	let oFormat = null;
	const nLang = this.Get_CompiledPr().Lang.Val;
	// let sFormat = AscCommonWord.oDefaultDateTimeFormat[nLang];
	// if(!sFormat)
	// {
		// choose pre-defined format
		// sFormat = oDateTimeFormats[sResultFiledType]
	// }
	// if(sFormat) {
		// get language dependant formats array
	let sFormat;

	let aFormat = Asc.c_oAscDateTimeFormat[nLang];
	if (!Array.isArray(aFormat)) {
		aFormat = Asc.c_oAscDateTimeFormat[lcid_enUS];
	}
	sFormat = getLanguageDependantFormat(aFormat, vsdxFieldFormat);
	oFormat = AscCommon.oNumFormatCache.get(sFormat, AscCommon.NumFormatType.WordFieldDate);
	// }
	return oFormat;
};

/** @constructor */
function cDate(date) {
	// original
	var bind = Function.bind;
	var unbind = bind.bind(bind);
	var date = new (unbind(Date, null).apply(null, arguments));
	date.__proto__ = cDate.prototype;
	return date;
}

Asc.cDate.prototype.getUTCFullYear = function () {
	// original
	var year = Date.prototype.getUTCFullYear.call(this);
	var month = Date.prototype.getUTCMonth.call(this);
	var date = Date.prototype.getUTCDate.call(this);

	if (1899 == year && 11 == month && (30 === date || 31 === date)) {
		return 1900;
	} else {
		return year;
	}
};

AscCommonWord.CPresentationField.prototype.private_GetString = function()
{
	//todo add num formats with units in editor

	// return; returns direct tag value
	// if (!this.isTextInherited) {
	// 	return;
	// }


	/**
	 *
	 * @param valueV
	 * @param {string} valueUnits
	 * @return {number}
	 */
	function convertConsiderUnits(valueV, valueUnits) {
		/**
		 * @type {(number)}
		 */
		let valueInProperUnits;
		const precision = 4;
		if (valueUnits === "CM") {
			valueInProperUnits = Number(valueV) * g_dKoef_in_to_mm / 10;
		} else if (valueUnits === "MM") {
			valueInProperUnits = Number(valueV) * g_dKoef_in_to_mm;
		} else {
			valueInProperUnits = valueV;
		}

		if (typeof valueInProperUnits === "number") {
			valueInProperUnits = Math.round(valueInProperUnits * Math.pow(10, precision)) / Math.pow(10, precision);
		}
		return valueInProperUnits;
	}

	var sStr = null;
	var oStylesObject;
	var oCultureInfo = AscCommon.g_aCultureInfos[this.Get_CompiledPr().Lang.Val];
	if(!oCultureInfo)
	{
		oCultureInfo = AscCommon.g_aCultureInfos[1033];
	}
	var oDateTime;
	if(typeof this.FieldType === 'string')
	{
		// let format;
		// if (this.vsdxFieldFormat) {
		// 	format = parseFieldPictureFormat(this.vsdxFieldValue, this.vsdxFieldFormat);
		// }
		let logicDocument = this.Paragraph && this.Paragraph.GetLogicDocument();
		const sFieldType = this.FieldType.toUpperCase();

		// let's not use formula (valueCell.f) for now
		// first convert value (valueCell.v) which is inches by default to units set in valueCell.u
		let val = convertConsiderUnits(this.vsdxFieldValue.v, this.vsdxFieldValue.u);


		// get known functions and get val otherwise return tag text
		// time formats are not handled correctly for now so date formats are commented

		// todo check INH sFieldType

		if("PAGECOUNT()" === sFieldType)
		{
			if (logicDocument) {
				val = logicDocument.getCountPages();
			}
		}
		else if("PAGENUMBER()" === sFieldType) {
			if (logicDocument) {
				val = logicDocument.getCurrentPage();
			}
		}
 		// else if("NOW()" === sFieldType)
		// {
		// 	let oDateTime = new Asc.cDate();
		// 	val = oDateTime.getExcelDateWithTime(true);
		// }
		// else if("DOCCREATION()" === sFieldType)
		// {
		// 	let oDateTime
		// 	if (logicDocument.core && logicDocument.core.created) {
		// 		oDateTime = new Asc.cDate(logicDocument.core.created);
		// 	} else {
		// 		oDateTime = new Asc.cDate(val);
		// 	}
		// 	val = oDateTime.getExcelDateWithTime(true);
		// }
		else if("CREATOR()" === sFieldType)
		{
			if (logicDocument.core && logicDocument.core.creator) {
				val = logicDocument.core.creator;
			}
		}
		else if("WIDTH" === sFieldType)
		{
			//todo display units
			val = this.vsdxFieldValue.getValueInMM();
		}
		else if ((this.vsdxFieldValue.u === "STR" || !this.vsdxFieldValue.u) && (sFieldType === "INH" || !sFieldType)) {
			// handle simple values. consider is function is INH value is calculated correctly already
			val = this.vsdxFieldValue.v;
		}
		else
		{
			return;
		}

		// if (this.vsdxFieldValue.u === "DATE") {
		// 	// TODO fix 31.12.1899 visio date
		// 	const oFormat = this.private_GetDateTimeFormat(this.vsdxFieldValue,
		// 		this.vsdxFieldFormat);
		// 	if(oFormat)
		// 	{
		// 		let dateString = this.vsdxFieldValue.v;
		// 		dateString = dateString === "" ? dateString : dateString;
		// 		oDateTime = new Asc.cDate(dateString);
		//
		// 		sStr = oFormat.formatToWord(oDateTime.getExcelDate(false) + 1 + (oDateTime.getHours() * 60 * 60 + oDateTime.getMinutes() * 60 + oDateTime.getSeconds()) / AscCommonExcel.c_sPerDay, 15, oCultureInfo);
		// 	}
		// 	// const oFormat = AscCommon.oNumFormatCache.get(format, AscCommon.NumFormatType.Excel);
		// 	// sStr =  format._formatToText(val, AscCommon.CellValueType.String, 15, oCultureInfo);
		// } else {
		let format;
		if (this.vsdxFieldFormat) {
			format = parseFieldPictureFormat(this.vsdxFieldValue, this.vsdxFieldFormat);
		}
		const oFormat = AscCommon.oNumFormatCache.get(format, AscCommon.NumFormatType.Excel);
		sStr =  oFormat._formatToText(val, AscCommon.CellValueType.String, 15, oCultureInfo);
		// sStr = val + "";
		// }
	}
	return sStr;
};

//todo CMobileDelegateEditorDiagram
AscCommon.CMobileDelegateEditorPresentation.prototype.GetObjectTrack = function(x, y, page, bSelected, bText) { return false; }
AscCommon.CMobileDelegateEditorPresentation.prototype.GetSelectionRectsBounds = function () { return null; }
AscCommon.CMobileDelegateEditorPresentation.prototype.GetSelectionRectsBounds = function () { return null; }
AscCommon.CMobileDelegateEditorPresentation.prototype.GetContextMenuType = function () { return AscCommon.MobileTouchContextMenuType.None; }
AscCommon.CMobileDelegateEditorPresentation.prototype.GetContextMenuInfo = function () {}
AscCommon.CMobileDelegateEditorPresentation.prototype.GetContextMenuPosition = function () { return { X : 0, Y : 0, Mode : AscCommon.MobileTouchContextMenuType.None }; }
AscCommon.CMobileDelegateEditorPresentation.prototype.Logic_GetNearestPos = function () { return null; }
