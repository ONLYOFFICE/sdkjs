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
	let fmtScheme = isConnectorShape ?
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
	let fmtScheme = isConnectorShape ?
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
function parseFieldPictureFormat(vsdxFieldFormat) {
	let res = "@";
	if (vsdxFieldFormat.f) {
		let formatFunction = vsdxFieldFormat.f.toUpperCase();
		let vFieldPicture = parseInt(formatFunction.substring('FIELDPICTURE('.length));
		switch (vFieldPicture) {
			case 0:
				res = "General";
				break;
			case 37:
				res = "@";
				break;
			case 200:
				res = "M/d/yyyy";
				break;
			case 212:
				res = "M/d/yyyy h:mm:ss am/pm";
				break;
		}
	} else if (vsdxFieldFormat.v) {
		res = vsdxFieldFormat.v;
	}
	return res;
}
AscCommonWord.CPresentationField.prototype.private_GetString = function()
{
	var sStr = null;
	var oStylesObject;
	var oCultureInfo = AscCommon.g_aCultureInfos[this.Get_CompiledPr().Lang.Val];
	if(!oCultureInfo)
	{
		oCultureInfo = AscCommon.g_aCultureInfos[1033];
	}
	var oDateTime, oFormat;
	if(typeof this.FieldType === 'string')
	{
		let format;
		if (this.vsdxFieldFormat) {
			format = parseFieldPictureFormat(this.vsdxFieldFormat);
		}
		let logicDocument = this.Paragraph && this.Paragraph.GetLogicDocument();
		const sFieldType = this.FieldType.toUpperCase();
		let val = this.vsdxFieldValue.v;
		if("PAGECOUNT()" === sFieldType)
		{
			if (logicDocument) {
				val = logicDocument.getCountPages();
			}
		}
		else if("NOW()" === sFieldType)
		{
			let oDateTime = new Asc.cDate();
			val = oDateTime.getExcelDateWithTime(true);
		}
		else if("DOCCREATION()" === sFieldType)
		{
			let oDateTime
			if (logicDocument.core && logicDocument.core.created) {
				oDateTime = new Asc.cDate(logicDocument.core.created);
			} else {
				oDateTime = new Asc.cDate(val);
			}
			val = oDateTime.getExcelDateWithTime(true);
		}
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
		if (format) {
			const oFormat = AscCommon.oNumFormatCache.get(format, AscCommon.NumFormatType.Excel);
			sStr =  oFormat._formatToText(val, AscCommon.CellValueType.String, 15, oCultureInfo);
		} else {
			sStr = val + "";
		}
	}
	return sStr;
};

