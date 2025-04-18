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
var c_oAscDocumentUnits = Asc.c_oAscDocumentUnits;
var global_mouseEvent = AscCommon.global_mouseEvent;
var g_dKoef_pix_to_mm = AscCommon.g_dKoef_pix_to_mm;
var g_dKoef_mm_to_pix = AscCommon.g_dKoef_mm_to_pix;

function CTab(pos, type, leader)
{
	this.pos    = pos;
	this.type   = type;
	this.leader = leader;
}

var g_array_objects_length = 1;

var RULER_OBJECT_TYPE_PARAGRAPH = 1;
var RULER_OBJECT_TYPE_HEADER    = 2;
var RULER_OBJECT_TYPE_FOOTER    = 4;
var RULER_OBJECT_TYPE_TABLE     = 8;
var RULER_OBJECT_TYPE_COLUMNS   = 16;

function CHorRulerRepaintChecker()
{
    // zoom/section check
    this.Width = 0;
    this.Height = 0;

    // ruler type check
    this.Type = 0;

    // margin check
    this.MarginLeft = 0;
    this.MarginRight = 0;

    // table column check
    this.tableCols = [];
    this.marginsLeft = [];
    this.marginsRight = [];

    this.columns = null; // CColumnsMarkup

    // blit to main params
    this.BlitAttack = false;
    this.BlitLeft = 0;
    this.BlitIndentLeft = 0;
    this.BlitIndentLeftFirst = 0;
    this.BlitIndentRight = 0;
    this.BlitDefaultTab = 0;
    this.BlitTabs = null;
	this.BlitRtl = false;

    this.BlitMarginLeftInd = 0;
    this.BlitMarginRightInd = 0;
}

function CVerRulerRepaintChecker()
{
    // zoom/section check
    this.Width = 0;
    this.Height = 0;

    // ruler type check
    this.Type = 0;

    // margin check
    this.MarginTop = 0;
    this.MarginBottom = 0;

    // header/footer check
    this.HeaderTop = 0;
    this.HeaderBottom = 0;

    // table column check
    this.rowsY = [];
    this.rowsH = [];

    // blit params
    this.BlitAttack = false;
    this.BlitTop = 0;
}

function RulerCorrectPosition(_ruler, val, margin)
{
    if (AscCommon.global_keyboardEvent.AltKey)
        return val;

    var mm_1_4 = 10 / 4;

    if (_ruler.Units == c_oAscDocumentUnits.Inch)
        mm_1_4 = 25.4 / 16;
    else if (_ruler.Units == c_oAscDocumentUnits.Point)
        mm_1_4 = 25.4 / 12;

    var mm_1_8 = mm_1_4 / 2;

    if (undefined === margin)
        return (((val + mm_1_8) / mm_1_4) >> 0) * mm_1_4;

    if (val >= margin)
        return margin + (((val - margin + mm_1_8) / mm_1_4) >> 0) * mm_1_4;

    return margin + (((val - margin - mm_1_8) / mm_1_4) >> 0) * mm_1_4;
}

function RulerCheckSimpleChanges()
{
    this.X = -1;
    this.Y = -1;
    this.IsSimple = true;

    this.IsDown   = false;
}
RulerCheckSimpleChanges.prototype =
{
    Clear : function()
    {
        this.X = -1;
        this.Y = -1;
        this.IsSimple = true;
        this.IsDown = false;
    },

    Reinit : function()
    {
        this.X = global_mouseEvent.X;
        this.Y = global_mouseEvent.Y;
        this.IsSimple = true;
        this.IsDown = true;
    },

    CheckMove : function()
    {
        if (!this.IsDown)
            return;
        if (!this.IsSimple)
            return;

        if (Math.abs(global_mouseEvent.X - this.X) > 0 || Math.abs(global_mouseEvent.Y - this.Y) > 0)
            this.IsSimple = false;
    }
};

(function(){
	AscWord.RULER_DRAG_TYPE = {
		none : 0,
		leftMargin : 1,
		rightMargin : 2,
		leftFirstInd : 3,
		leftInd : 4,
		firstInd : 5,
		rightInd : 6,
		tab : 7,
		table : 8,
		columnSize : 9,
		columnPos : 10
	};
})();

function CHorRuler()
{
    this.m_oPage        = null;     // текущая страница. Нужна для размеров и маргинов в режиме RULER_OBJECT_TYPE_PARAGRAPH
    
    this.m_nTop         = 0;        // начало прямогулольника линейки
    this.m_nBottom      = 0;        // конец прямоугольника линейки

    // реализация tab'ов
    this.m_dDefaultTab              = 12.5;
    this.m_arrTabs                  = [];
    this.m_lCurrentTab              = -1;
    this.m_dCurrentTabNewPosition   = -1;
    this.m_dMaxTab                  = 0;
    this.IsDrawingCurTab            = true; // это подсказка для пользователя - будет ли оставлен таб после отпускания мышки, или будет выкинут

    this.m_dMarginLeft              = 20;
    this.m_dMarginRight             = 190;

    this.m_dIndentLeft          = 10;
    this.m_dIndentRight         = 20;
    this.m_dIndentLeftFirst     = 20;
	this.m_bRtl                 = false;

    this.m_oCanvas              = null;

    this.m_dZoom                = 1;
	
	this.DragType          = AscWord.RULER_DRAG_TYPE.none;
	this.DragTypeMouseDown = AscWord.RULER_DRAG_TYPE.none;

    this.m_dIndentLeft_old      = -10000;
    this.m_dIndentLeftFirst_old = -10000;
    this.m_dIndentRight_old     = -10000;

    // отдельные настройки для текущего объекта линейки
    this.CurrentObjectType  = RULER_OBJECT_TYPE_PARAGRAPH;
    this.m_oTableMarkup     = null;
    this.m_oColumnMarkup    = null;
    this.DragTablePos       = -1;

    this.TableMarginLeft = 0;
    this.TableMarginLeftTrackStart = 0;
    this.TableMarginRight   = 0;

    this.m_oWordControl = null;

    this.RepaintChecker = new CHorRulerRepaintChecker();
    this.m_bIsMouseDown = false;

    // presentations addons
    this.IsCanMoveMargins = true;
    this.IsCanMoveAnyMarkers = true;
    this.IsDrawAnyMarkers = true;

    this.SimpleChanges = new RulerCheckSimpleChanges();

    this.Units = c_oAscDocumentUnits.Millimeter;

    this.DrawTablePict = function() {
        var ctx = g_memory.ctx;
        var dPR = AscCommon.AscBrowser.retinaPixelRatio;
        var isNeedRedraw = (dPR - Math.floor(dPR)) >= 0.5 ? true : false;
        var roundDPR = isNeedRedraw ? Math.floor(dPR) : Math.round(dPR);

        var canvasWidth = (isNeedRedraw && dPR >= 1 ) ? Math.round(7 * (Math.floor(dPR) + 0.5)) : 7 * (isNeedRedraw ? dPR : ( dPR >= 1 ? Math.round(dPR) : dPR)),
            canvasHeight = (isNeedRedraw && dPR >= 1) ? Math.round(8 * (Math.floor(dPR) + 0.5)) : 8 * (isNeedRedraw ? dPR : ( dPR >= 1 ? Math.round(dPR) : dPR));

        if (null != this.tableSprite)
        {
            if (ctx.canvas.width == canvasWidth)
                return;
        }

        ctx.canvas.width = canvasWidth;
        ctx.canvas.height = canvasHeight;


        ctx.beginPath();
        ctx.fillStyle = '#ffffff';
        ctx.strokeStyle = '#ffffff';
        ctx.stroke();
        ctx.rect(0,0,ctx.canvas.width, ctx.canvas.height)
        ctx.fill();

        ctx.beginPath();
        ctx.strokeStyle = '#646464';
        ctx.lineWidth = roundDPR;

        var step = isNeedRedraw ? Math.round(dPR): ctx.lineWidth;
        for (var i = 0; i < 7 * ctx.canvas.width; i += step + ctx.lineWidth) {
            ctx.moveTo(0,  0.5 * ctx.lineWidth + step + i);
            ctx.lineTo(ctx.canvas.width, 0.5 * ctx.lineWidth + step + i);
            ctx.stroke();
        }
        for (var i = 0; i < 8 * ctx.canvas.height; i += step + ctx.lineWidth) {
            ctx.moveTo(0.5 * ctx.lineWidth + step + i, 0);
            ctx.lineTo(0.5 * ctx.lineWidth + step + i, ctx.canvas.height);
            ctx.stroke();
        }

        return ctx.canvas;
    }

    this.tableSprite = null;

    this.CheckCanvas = function()
    {
        this.m_dZoom = this.m_oWordControl.m_nZoomValue / 100;
		var dPR = AscCommon.AscBrowser.retinaPixelRatio;

        var tablePict = this.DrawTablePict();
        if (tablePict)
            this.tableSprite = tablePict;

        var dKoef_mm_to_pix = g_dKoef_mm_to_pix * this.m_dZoom;
        dKoef_mm_to_pix *= dPR;

        var widthNew    = dKoef_mm_to_pix * this.m_oPage.width_mm;

        var _width      = 10 * Math.round(dPR) + widthNew;
        if (dPR > 1)
            _width += Math.round(dPR);

        var _height     = 8 * g_dKoef_mm_to_pix * dPR;

        var intW = _width >> 0;
        var intH = _height >> 0;
        if (null == this.m_oCanvas)
        {
            this.m_oCanvas = document.createElement('canvas');
            this.m_oCanvas.width    = intW;
            this.m_oCanvas.height   = intH;
        }
        else
        {
            var oldW = this.m_oCanvas.width;
            var oldH = this.m_oCanvas.height;

            if ((oldW != intW) || (oldH != intH))
            {
                delete this.m_oCanvas;
                this.m_oCanvas = document.createElement('canvas');
                this.m_oCanvas.width    = intW;
                this.m_oCanvas.height   = intH;
            }
        }
        return widthNew;
    }

    this.CreateBackground = function(cachedPage, isattack)
    {
        if (window["NATIVE_EDITOR_ENJINE"])
            return;
            
        if (null == cachedPage || undefined == cachedPage)
            return;

        this.m_oPage = cachedPage;
        var width = this.CheckCanvas();

        if (0 == this.DragType)
        {
            this.m_dMarginLeft  = cachedPage.margin_left;
            this.m_dMarginRight = cachedPage.margin_right;
        }

        // check old state
        var checker = this.RepaintChecker;
        var markup = this.m_oTableMarkup;
        if (this.CurrentObjectType == RULER_OBJECT_TYPE_COLUMNS)
            markup = this.m_oColumnMarkup;

        if (isattack !== true && this.CurrentObjectType == checker.Type && width == checker.Width)
        {
            if (this.CurrentObjectType == RULER_OBJECT_TYPE_PARAGRAPH)
            {
                if (this.m_dMarginLeft == checker.MarginLeft && this.m_dMarginRight == checker.MarginRight)
                    return;
            }
            else if (this.CurrentObjectType == RULER_OBJECT_TYPE_TABLE)
            {
                var oldcount = checker.tableCols.length;
                var newcount = 1 + markup.Cols.length;

                if (oldcount == newcount)
                {
                    var arr1 = checker.tableCols;
                    var arr2 = markup.Cols;

                    if (arr1[0] == markup.X)
                    {
                        var _break = false;
                        for (var i = 1; i < newcount; i++)
                        {
                            if (arr1[i] != arr2[i - 1])
                            {
                                _break = true;
                                break;
                            }
                        }

                        if (!_break)
                        {
                            --newcount;
                            var _margs = markup.Margins;

                            for (var i = 0; i < newcount; i++)
                            {
                                if (_margs[i].Left != checker.marginsLeft[i] || _margs[i].Right != checker.marginsRight[i])
                                {
                                    _break = true;
                                    break;
                                }
                            }

                            if (!_break)
                                return;
                        }
                    }
                }
            }
            else if (this.CurrentObjectType == RULER_OBJECT_TYPE_COLUMNS)
            {
                if (this.m_oColumnMarkup.X == checker.columns.X)
                {
                    if (markup.EqualWidth == checker.columns.EqualWidth)
                    {
                        if (markup.EqualWidth)
                        {
                            if (markup.Num == checker.columns.Num && markup.Space == checker.columns.Space && markup.R == checker.columns.R)
                                return;
                        }
                        else
                        {
                            var _arr1 = markup.Cols;
                            var _arr2 = checker.columns.Cols;
                            if (_arr1 && _arr2 && _arr1.length == _arr2.length)
                            {
                                var _len = _arr1.length;
                                var _index = 0;
                                for (_index = 0; _index < _len; _index++)
                                {
                                    if (_arr1[_index].W != _arr2[_index].W || _arr1[_index].Space != _arr2[_index].Space)
                                        break;
                                }

                                if (_index == _len)
                                    return;
                            }
                        }
                    }
                }
            }
        }

        //console.log("horizontal");

        checker.Width = width;
        checker.Type = this.CurrentObjectType;
        checker.BlitAttack = true;

        var dPR = AscCommon.AscBrowser.retinaPixelRatio;
        var dKoef_mm_to_pix = g_dKoef_mm_to_pix * this.m_dZoom * dPR;

        // не править !!!
        this.m_nTop     = Math.round(6 * dPR);//(1.8 * g_dKoef_mm_to_pix) >> 0;
        this.m_nBottom  = Math.round(19 * dPR);//(5.2 * g_dKoef_mm_to_pix) >> 0;

        var context = this.m_oCanvas.getContext('2d');
        context.setTransform(1, 0, 0, 1, 5 * Math.round( dPR), 0);


        context.fillStyle = GlobalSkin.BackgroundColor;
        context.fillRect(0, 0, this.m_oCanvas.width, this.m_oCanvas.height);

		// промежуток между маргинами
        var left_margin  = 0;
        var right_margin = 0;
		
		let isRtl = false;
        if (this.CurrentObjectType == RULER_OBJECT_TYPE_PARAGRAPH)
        {
            left_margin  = (this.m_dMarginLeft * dKoef_mm_to_pix) >> 0;
            right_margin = (this.m_dMarginRight * dKoef_mm_to_pix) >> 0;

            checker.MarginLeft = this.m_dMarginLeft;
            checker.MarginRight = this.m_dMarginRight;
			isRtl = this.m_bRtl;
        }
        else if (this.CurrentObjectType == RULER_OBJECT_TYPE_TABLE && null != markup)
        {
            var _cols = checker.tableCols;
            if (0 != _cols.length)
                _cols.splice(0, _cols.length);

            _cols[0] = markup.X;
            var _ml = checker.marginsLeft;
            if (0 != _ml.length)
                _ml.splice(0, _ml.length);

            var _mr = checker.marginsRight;
            if (0 != _mr.length)
                _mr.splice(0, _mr.length);

            var _count_ = markup.Cols.length;

            for (var i = 0; i < _count_; i++)
            {
                _cols[i + 1] = markup.Cols[i];
                _ml[i] = markup.Margins[i].Left;
                _mr[i] = markup.Margins[i].Right;
            }

            if (0 != _count_)
            {
                var _start = 0;
                for (var i = 0; i < _count_; i++)
                {
                    _start += markup.Cols[i];
                }

                left_margin  = ((markup.X + markup.Margins[0].Left) * dKoef_mm_to_pix) >> 0;
                right_margin = ((markup.X + _start - markup.Margins[markup.Margins.length - 1].Right) * dKoef_mm_to_pix) >> 0;
            }
        }
        else if (this.CurrentObjectType == RULER_OBJECT_TYPE_COLUMNS && null != markup)
        {
            left_margin  = (markup.X * dKoef_mm_to_pix) >> 0;
            right_margin = (markup.R * dKoef_mm_to_pix) >> 0;

            checker.MarginLeft = this.m_dMarginLeft;
            checker.MarginRight = this.m_dMarginRight;

            checker.columns = this.m_oColumnMarkup.CreateDuplicate();
        }
        var indent = 0.5 * Math.round(dPR);
        context.fillStyle = GlobalSkin.RulerLight;
        context.fillRect(left_margin + indent, this.m_nTop + indent, right_margin - left_margin, this.m_nBottom - this.m_nTop);

        var intW = width >> 0;

        if (true)
        {
            context.beginPath();
            context.fillStyle = GlobalSkin.RulerDark;

            context.fillRect(indent, this.m_nTop + indent, left_margin, this.m_nBottom - this.m_nTop);
            context.fillRect(right_margin + indent, this.m_nTop + indent, Math.max(intW - right_margin, 1), this.m_nBottom - this.m_nTop);
            context.beginPath();
        }

		// рамка
        //context.shadowBlur = 0;
        //context.shadowColor = "#81878F";

        context.beginPath();
        context.lineWidth = Math.round(dPR);
        context.strokeStyle = GlobalSkin.RulerTextColor;
        context.fillStyle = GlobalSkin.RulerTextColor;
		
		let mm_1_4     = 10 * dKoef_mm_to_pix / 4;
		let inch_1_8   = 25.4 * dKoef_mm_to_pix / 8;
		let point_1_12 = 25.4 * dKoef_mm_to_pix / 12;

        var middleVert = (this.m_nTop + this.m_nBottom) / 2;
        var part1 = 1.5 * Math.round(dPR);
        var part2 = 2.5 * Math.round(dPR);

        context.font = Math.round(7 * dPR) + "pt Arial";
		
		let _bottom = this.m_nBottom;
		function drawLayoutMM(x, step, count)
		{
			let isDraw1_4 = Math.abs(step) > 7;
			let index = 0;
			let num = 0;
			for (let i = 1; i <= count; ++i)
			{
				let lXPos = ((x + i * step) >> 0) + indent;
				index++;
				
				if (index === 4)
					index = 0;
				
				if (0 === index)
				{
					num++;
					// number
					let strNum = "" + num;
					let lWidthText = context.measureText(strNum).width;
					lXPos -= (lWidthText / 2.0);
					context.fillText(strNum, lXPos, _bottom - Math.round(3 * dPR));
				}
				else if (1 === index && isDraw1_4)
				{
					// 1/4
					context.beginPath();
					context.moveTo(lXPos, middleVert - part1);
					context.lineTo(lXPos, middleVert + part1);
					context.stroke();
				}
				else if (2 === index)
				{
					// 1/2
					context.beginPath();
					context.moveTo(lXPos, middleVert - part2);
					context.lineTo(lXPos, middleVert + part2);
					context.stroke();
				}
				else if (isDraw1_4)
				{
					// 1/4
					context.beginPath();
					context.moveTo(lXPos, middleVert - part1);
					context.lineTo(lXPos, middleVert + part1);
					context.stroke();
				}
			}
		}
		function drawLayoutInch(x, step, count)
		{
			let isDraw1_8 = Math.abs(step) > 8;
			let index = 0;
			let num   = 0;
			for (let i = 1; i <= count; ++i)
			{
				let lXPos = ((x + i * step) >> 0) + indent;
				index++;
				
				if (index === 8)
					index = 0;
				
				if (0 === index)
				{
					num++;
					// number
					let strNum     = "" + num;
					let lWidthText = context.measureText(strNum).width;
					lXPos -= (lWidthText / 2.0);
					context.fillText(strNum, lXPos, _bottom - Math.round(3 * dPR));
				}
				else if (4 === index)
				{
					// 1/2
					context.beginPath();
					context.moveTo(lXPos, middleVert - part2);
					context.lineTo(lXPos, middleVert + part2);
					context.stroke();
				}
				else if (isDraw1_8)
				{
					// 1/8
					context.beginPath();
					context.moveTo(lXPos, middleVert - part1);
					context.lineTo(lXPos, middleVert + part1);
					context.stroke();
				}
			}
		}
		function drawLayoutPt(x, step, count)
		{
			let isDraw1_12 = Math.abs(step) > 5;
			let index = 0;
			let num   = 0;
			for (let i = 1; i <= count; ++i)
			{
				let lXPos = ((x + i * step) >> 0) + indent;
				index++;
				
				if (index === 12)
					index = 0;
				
				if (0 === index || 6 === index)
				{
					num++;
					// number
					let strNum     = "" + (num * 36);
					let lWidthText = context.measureText(strNum).width;
					lXPos -= (lWidthText / 2.0);
					context.fillText(strNum, lXPos, _bottom - Math.round(3 * dPR));
				}
				else if (isDraw1_12)
				{
					// 1/12
					context.beginPath();
					context.moveTo(lXPos, middleVert - part1);
					context.lineTo(lXPos, middleVert + part1);
					context.stroke();
				}
			}
		}
		
		let step = point_1_12;
		let drawFunc = null;
		
		if (this.Units === c_oAscDocumentUnits.Millimeter)
		{
			step = mm_1_4;
			drawFunc = drawLayoutMM;
		}
		else if (this.Units === c_oAscDocumentUnits.Inch)
		{
			step = inch_1_8;
			drawFunc = drawLayoutInch;
		}
		else if (this.Units === c_oAscDocumentUnits.Point)
		{
			step = point_1_12;
			drawFunc = drawLayoutPt;
		}
		
		if (drawFunc)
		{
			let zeroX  = isRtl ? right_margin : left_margin;
			let rCount = isRtl ? ((width - right_margin) / step) >> 0 : (((width - left_margin) / step) >> 0) - 1;
			let lCount = isRtl ? ((right_margin / step) >> 0) - 1 : left_margin / step >> 0;
			
			drawFunc(zeroX, step, rCount);
			drawFunc(zeroX, -step, lCount);
		}

        if (null != markup && this.CurrentObjectType == RULER_OBJECT_TYPE_TABLE)
        {
            var _count = markup.Cols.length;
            if (0 != _count)
            {
                context.fillStyle = GlobalSkin.RulerDark;
                context.strokeStyle = GlobalSkin.RulerOutline;

                var _offset = markup.X;
                for (var i = 0; i <= _count; i++)
                {
                    var __xID = 0;
                        __xID = (-2.5 * dPR + _offset * dKoef_mm_to_pix) >> 0;

                    var __yID = this.m_nBottom - Math.round(10 * dPR);

                    if (0 == i)
                    {
                        context.drawImage(this.tableSprite, __xID, __yID);
                        _offset += markup.Cols[i];
                        continue;
                    }
                    if (i == _count)
                    {
                        context.drawImage(this.tableSprite, __xID, __yID);
                        break;
                    }

                    var __x = (((_offset - markup.Margins[i-1].Right) * dKoef_mm_to_pix) >> 0) + indent;
                    var __r = (((_offset + markup.Margins[i].Left) * dKoef_mm_to_pix) >> 0) + indent;

                    context.fillRect(__x, this.m_nTop + indent, __r - __x, this.m_nBottom - this.m_nTop);
                    context.strokeRect(__x, this.m_nTop + indent, __r - __x, this.m_nBottom - this.m_nTop);
                    context.drawImage(this.tableSprite, __xID, __yID);

                    _offset += markup.Cols[i];
                }
            }
        }

        if (null != markup && this.CurrentObjectType == RULER_OBJECT_TYPE_COLUMNS)
        {
            var _array = markup.EqualWidth ? [] : markup.Cols;
            if (markup.EqualWidth)
            {
                var _w = ((markup.R - markup.X) - markup.Space * (markup.Num - 1)) / markup.Num;

                for (var i = 0; i < markup.Num; i++)
                {
                    var _cur = new CColumnsMarkupColumn();
                    _cur.W = _w;
                    _cur.Space = markup.Space;
                    _array.push(_cur);
                }
            }

            var _count = _array.length;
            if (0 != _count)
            {
                context.fillStyle = GlobalSkin.RulerDark;
                context.strokeStyle = GlobalSkin.RulerOutline;

                var _offsetX = markup.X;
                for (var i = 0; i < _count; i++)
                {
                    var __xTmp = _offsetX + _array[i].W;
                    var __rTmp = __xTmp + _array[i].Space;

                    var _offset = (__xTmp + __rTmp) / 2;

                    if (i == (_count - 1))
                        continue;

                    var __xID = 0;
                        __xID = ((2.5 + _offset * dKoef_mm_to_pix)) * dPR >> 0;

                    var __yID = this.m_nBottom - Math.round(10 * dPR);

                    var __x = ((__xTmp * dKoef_mm_to_pix) >> 0) + indent;
                    var __r = ((__rTmp * dKoef_mm_to_pix) >> 0) + indent;

                    context.fillRect(__x, this.m_nTop + indent, __r - __x, this.m_nBottom - this.m_nTop);
                    context.strokeRect(__x, this.m_nTop + indent, __r - __x, this.m_nBottom - this.m_nTop);

                    if (!markup.EqualWidth)
                        context.drawImage(this.tableSprite, __xID, __yID);

                    if ((__r - __x) > 10)
                    {
                        context.fillStyle = GlobalSkin.RulerLight;
                        context.strokeStyle = GlobalSkin.RulerMarkersOutlineColor;

                        var roundV1 = Math.round(3 * dPR);
                        var roundV2 = Math.round(6 * dPR);
                        context.fillRect(__x + roundV1, this.m_nTop + indent + roundV1, roundV1, this.m_nBottom - this.m_nTop - roundV2);
                        context.fillRect(__r - roundV2, this.m_nTop + indent + roundV1, roundV1, this.m_nBottom - this.m_nTop - roundV2);
                        context.strokeRect(__x + roundV1, this.m_nTop + indent + roundV1, roundV1, this.m_nBottom - this.m_nTop - roundV2);
                        context.strokeRect(__r - roundV2, this.m_nTop + indent + roundV1, roundV1, this.m_nBottom - this.m_nTop - roundV2);
                        context.fillStyle = GlobalSkin.RulerDark;
                        context.strokeStyle = GlobalSkin.RulerOutline;
                    }

                    _offsetX += (_array[i].W + _array[i].Space);
                }
            }
        }

        context.beginPath();
        context.strokeStyle = GlobalSkin.RulerOutline;
        context.strokeRect(indent, this.m_nTop + indent, Math.max(intW - 1, 1), this.m_nBottom - this.m_nTop);
        context.beginPath();
        context.moveTo(left_margin + indent, this.m_nTop + indent);
        context.lineTo(left_margin + indent, this.m_nBottom - indent);

        context.moveTo(right_margin + indent, this.m_nTop + indent);
        context.lineTo(right_margin + indent, this.m_nBottom - indent);

        context.stroke();
    }

    this.CorrectTabs = function()
    {
        this.m_dMaxTab = 0;
        var _old_c = this.m_arrTabs.length;
        if (0 == _old_c)
            return;

        var _old = this.m_arrTabs;
        var _new = [];

        for (var i = 0; i < _old_c; i++)
        {
            for (var j = i + 1; j < _old_c; j++)
            {
                if (_old[j].pos < _old[i].pos)
                {
                    var temp = _old[i];
                    _old[i] = _old[j];
                    _old[j] = temp;
                }
            }
        }

        var _new_len = 0;
        _new[_new_len++] = _old[0];
        for (var i = 1; i < _old_c; i++)
        {
            if (_new[_new_len - 1].pos != _old[i].pos)
                _new[_new_len++] = _old[i];
        }
        this.m_arrTabs = null;
        this.m_arrTabs = _new;

        this.m_dMaxTab = this.m_arrTabs[_new_len - 1].pos;
    }

    this.CalculateMargins = function()
    {
        if (this.CurrentObjectType == RULER_OBJECT_TYPE_TABLE)
        {
            this.TableMarginLeft = 0;
            this.TableMarginRight = 0;

            var markup = this.m_oTableMarkup;
            var margin_left = markup.X;

            var _col = markup.CurCol;
            for (var i = 0; i < _col; i++)
                margin_left += markup.Cols[i];

            this.TableMarginLeft = margin_left + markup.Margins[_col].Left;
            this.TableMarginRight = margin_left + markup.Cols[_col] - markup.Margins[_col].Right;
        }
        else if (this.CurrentObjectType == RULER_OBJECT_TYPE_COLUMNS)
        {
            this.TableMarginLeft = 0;
            this.TableMarginRight = 0;

            var markup = this.m_oColumnMarkup;

            if (markup.EqualWidth)
            {
                var _w = ((markup.R - markup.X) - markup.Space * (markup.Num - 1)) / markup.Num;

                this.TableMarginLeft = markup.X + (_w  + markup.Space) * markup.CurCol;
                this.TableMarginRight = this.TableMarginLeft + _w;
            }
            else
            {
                var margin_left = markup.X;
                var _col = markup.CurCol;
                for (var i = 0; i < _col; i++)
                    margin_left += (markup.Cols[i].W + markup.Cols[i].Space);

                this.TableMarginLeft = margin_left;
                this.TableMarginRight = margin_left + markup.Cols[_col].W;

                var _x = markup.X;

                var _len = markup.Cols.length;
                for (var i = 0; i < _len; i++)
                {
                    _x += markup.Cols[i].W;

                    if (i != (_len - 1))
                        _x += markup.Cols[i].Space;
                }

                markup.R = _x;
            }
        }
    }

    this.OnMouseMove = function(left, top, e)
    {
        var word_control = this.m_oWordControl;

        AscCommon.check_MouseMoveEvent(e);

        this.SimpleChanges.CheckMove();

		// теперь определяем позицию относительно самой линейки. Все в миллиметрах
        var hor_ruler = word_control.m_oTopRuler_horRuler;
        var dKoefPxToMM = 100 * g_dKoef_pix_to_mm / word_control.m_nZoomValue;

        var _x = global_mouseEvent.X - left - word_control.X - word_control.GetMainContentBounds().L * g_dKoef_mm_to_pix;
        if (!word_control.m_oApi.isRtlInterface)
            _x -= 5 * g_dKoef_mm_to_pix;

        _x *= dKoefPxToMM;
        var _y = (global_mouseEvent.Y - word_control.Y) * g_dKoef_pix_to_mm;

        var dKoef_mm_to_pix = g_dKoef_mm_to_pix * this.m_dZoom;
        var mm_1_4 = 10 / 4;
        var mm_1_8 = mm_1_4 / 2;

        var _margin_left = this.m_dMarginLeft;
        var _margin_right = this.m_dMarginRight;
        if (this.CurrentObjectType == RULER_OBJECT_TYPE_TABLE || this.CurrentObjectType == RULER_OBJECT_TYPE_COLUMNS)
        {
            _margin_left = this.TableMarginLeft;
            _margin_right = this.TableMarginRight;
        }

        var _presentations = false;
        if (word_control.EditorType === "presentations")
            _presentations = true;

        switch (this.DragType)
        {
            case AscWord.RULER_DRAG_TYPE.none:
            {
                var position = this.CheckMouseType(_x, _y);
                if ((1 == position) || (2 == position) || (8 == position) || (9 == position) || (10 == position))
                    word_control.m_oDrawingDocument.SetCursorType("w-resize");
                else
                    word_control.m_oDrawingDocument.SetCursorType("default");

                break;
            }
            case AscWord.RULER_DRAG_TYPE.leftMargin:
            {
                var newVal = RulerCorrectPosition(this, _x, _margin_left);

                if (newVal < 0)
                    newVal = 0;

                var max = this.m_dMarginRight - 20;
                if (0 < this.m_dIndentRight)
                    max = (this.m_dMarginRight - this.m_dIndentRight - 20);
                if (newVal > max)
                    newVal = max;

                var _max_ind = Math.max(this.m_dIndentLeft, this.m_dIndentLeftFirst);
                if ((newVal + _max_ind) > max)
                    newVal = max - _max_ind;

                this.m_dMarginLeft = newVal;
                word_control.UpdateHorRulerBack();

                var pos = left + this.m_dMarginLeft * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.VertLine(pos);

                word_control.m_oDrawingDocument.SetCursorType("w-resize");
                break;
            }
            case AscWord.RULER_DRAG_TYPE.rightMargin:
            {
                var newVal = RulerCorrectPosition(this, _x, _margin_left);

                var min = this.m_dMarginLeft;
                if ((this.m_dMarginLeft + this.m_dIndentLeft) > min)
                    min = this.m_dMarginLeft + this.m_dIndentLeft;
                if ((this.m_dMarginLeft + this.m_dIndentLeftFirst) > min)
                    min = this.m_dMarginLeft + this.m_dIndentLeftFirst;

                min += 20;
                if (newVal < min)
                    newVal = min;
                if (newVal > this.m_oPage.width_mm)
                    newVal = this.m_oPage.width_mm;

                if ((newVal - this.m_dIndentRight) < min)
                    newVal = min + this.m_dIndentRight;

                this.m_dMarginRight = newVal;
                word_control.UpdateHorRulerBack();

                var pos = left + this.m_dMarginRight * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.VertLine(pos);

                word_control.m_oDrawingDocument.SetCursorType("w-resize");
                break;
            }
			case AscWord.RULER_DRAG_TYPE.leftFirstInd:
			{
				let newVal = RulerCorrectPosition(this, _x, this.m_bRtl ? _margin_right : _margin_left);
				
				if (this.m_bRtl)
				{
					let max = this.m_oPage.width_mm;
					if (_presentations)
						max = _margin_right;
					else if (this.m_dIndentLeftFirst < this.m_dIndentLeft)
						max = this.m_oPage.width_mm - (this.m_dIndentLeft - this.m_dIndentLeftFirst);
					
					if (newVal > max)
						newVal = max;
					
					let min = _margin_left;
					if (this.m_dIndentRight > 0)
						min = _margin_left + this.m_dIndentRight;
					
					if (this.m_dIndentLeftFirst > this.m_dIndentLeft)
						min += this.m_dIndentLeftFirst - this.m_dIndentLeft;
					
					if (newVal < min + 20)
						newVal = Math.min(min + 20, _margin_right - this.m_dIndentLeft_old);
					
					let newIndent = _margin_right - newVal;
					this.m_dIndentLeftFirst = (this.m_dIndentLeftFirst - this.m_dIndentLeft) + newIndent;
					this.m_dIndentLeft      = newIndent;
				}
				else
				{
					let min = 0;
					
					if (_presentations)
						min = _margin_left;
					else if (this.m_dIndentLeftFirst < this.m_dIndentLeft)
						min = this.m_dIndentLeft - this.m_dIndentLeftFirst;
					
					if (newVal < min)
						newVal = min;
					
					let max = _margin_right;
					if (0 < this.m_dIndentRight)
						max = _margin_right - this.m_dIndentRight;
					
					if (this.m_dIndentLeftFirst > this.m_dIndentLeft)
						max = max + (this.m_dIndentLeft - this.m_dIndentLeftFirst);
					
					if (newVal > (max - 20))
						newVal = Math.max(max - 20, (this.m_dIndentLeft_old + _margin_left));
					
					let newIndent           = newVal - _margin_left;
					this.m_dIndentLeftFirst = (this.m_dIndentLeftFirst - this.m_dIndentLeft) + newIndent;
					this.m_dIndentLeft      = newIndent;
				}
				let pos = left + newVal * dKoef_mm_to_pix;
				word_control.UpdateHorRulerBack();
				word_control.m_oOverlayApi.VertLine(pos);
				break;
			}
			case AscWord.RULER_DRAG_TYPE.leftInd:
			{
				let newVal = RulerCorrectPosition(this, _x, this.m_bRtl ? _margin_right : _margin_left);
				if (newVal < 0)
					newVal = 0;
				
				if (this.m_bRtl)
				{
					if (newVal > (this.m_oPage.width_mm))
						newVal = this.m_oPage.width_mm;
					
					let min = Math.max(_margin_left, _margin_left + this.m_dIndentRight) + 20;
					if (_presentations)
					{
						if (newVal > _margin_right)
							newVal = _margin_right;
					}
					
					if (newVal < min)
						newVal = Math.min(min, _margin_right - this.m_dIndentLeft_old);
					
					this.m_dIndentLeft = _margin_right - newVal;
				}
				else
				{
					let max = _margin_right - 20;
					if (0 < this.m_dIndentRight)
						max -= this.m_dIndentRight;
					
					if (_presentations)
					{
						if (newVal < _margin_left)
							newVal = _margin_left;
					}
					
					if (newVal > max)
						newVal = Math.max(max, _margin_left + this.m_dIndentLeft_old);
					
					this.m_dIndentLeft = newVal - _margin_left;
				}
				let pos = left + newVal * dKoef_mm_to_pix;
				word_control.UpdateHorRulerBack();
				word_control.m_oOverlayApi.VertLine(pos);
				break;
			}
			case AscWord.RULER_DRAG_TYPE.firstInd:
			{
				let newVal = RulerCorrectPosition(this, _x, this.m_bRtl ? _margin_right : _margin_left);
		
				if (newVal < 0)
					newVal = 0;
				
				if (this.m_bRtl)
				{
					if (newVal > (this.m_oPage.width_mm))
						newVal = this.m_oPage.width_mm;
					
					let min = 20 + Math.max(_margin_left, _margin_left + this.m_dIndentRight);
					if (_presentations)
					{
						if (newVal > _margin_right)
							newVal = _margin_right;
					}
					
					if (newVal < min)
						newVal = Math.min(min, _margin_right - this.m_dIndentLeftFirst_old);
					
					this.m_dIndentLeftFirst = _margin_right - newVal;
				}
				else
				{
					let max = _margin_right - 20;
					if (0 < this.m_dIndentRight)
						max -= this.m_dIndentRight;
					
					if (_presentations)
					{
						if (newVal < _margin_left)
							newVal = _margin_left;
					}
					
					if (newVal > max)
						newVal = Math.max(max, _margin_left + this.m_dIndentLeftFirst_old);
					
					this.m_dIndentLeftFirst = newVal - _margin_left;
				}
				let pos = left + newVal * dKoef_mm_to_pix;
				word_control.UpdateHorRulerBack();
				word_control.m_oOverlayApi.VertLine(pos);
				break;
			}
			case AscWord.RULER_DRAG_TYPE.rightInd:
			{
				let newVal = RulerCorrectPosition(this, _x, this.m_bRtl ? _margin_right : _margin_left);
				
				if (newVal > (this.m_oPage.width_mm))
					newVal = this.m_oPage.width_mm;
				if (newVal < 0)
					newVal = 0;
				
				if (this.m_bRtl)
				{
					let max = -20 + Math.min(_margin_right, _margin_right - this.m_dIndentLeft, _margin_right - this.m_dIndentLeftFirst);
					if (newVal > max)
						newVal = Math.max(max, _margin_left + this.m_dIndentRight_old);
					
					if (_presentations && newVal < _margin_left)
						newVal = _margin_left;
					
					this.m_dIndentRight = newVal - _margin_left;
				}
				else
				{
					let min = _margin_left;
					if ((_margin_left + this.m_dIndentLeft) > min)
						min = _margin_left + this.m_dIndentLeft;
					if ((_margin_left + this.m_dIndentLeftFirst) > min)
						min = _margin_left + this.m_dIndentLeftFirst;
					
					min += 20;
					
					if (newVal < min)
						newVal = Math.min(min, _margin_right - this.m_dIndentRight_old);
					
					if (_presentations)
					{
						if (newVal > _margin_right)
							newVal = _margin_right;
					}
					this.m_dIndentRight = _margin_right - newVal;
				}
				let pos = left + newVal * dKoef_mm_to_pix;
				word_control.UpdateHorRulerBack();
				word_control.m_oOverlayApi.VertLine(pos);
				break;
			}
			case AscWord.RULER_DRAG_TYPE.tab:
			{
				let newVal = RulerCorrectPosition(this, _x, this.m_bRtl ? _margin_right : _margin_left);
				
				this.m_dCurrentTabNewPosition = this.m_bRtl ? _margin_right - newVal : newVal - _margin_left;
				if (_y <= 3 || _y > 5.6)
				{
					this.IsDrawingCurTab = false;
					word_control.OnUpdateOverlay();
				}
				else
				{
					this.IsDrawingCurTab = true;
					let offset = this.m_bRtl ? _margin_right - this.m_dCurrentTabNewPosition : _margin_left + this.m_dCurrentTabNewPosition;
					let pos    = left + offset * dKoef_mm_to_pix;
					word_control.m_oOverlayApi.VertLine(pos);
				}
				word_control.UpdateHorRulerBack();
				break;
			}
            case AscWord.RULER_DRAG_TYPE.table:
            {
                var newVal = RulerCorrectPosition(this, _x, this.TableMarginLeftTrackStart);

                // сначала определим граничные условия
                var _min = 0;
                var _max = this.m_oPage.width_mm;

                var markup = this.m_oTableMarkup;
                var _left = 0;

                if (this.DragTablePos > 0)
                {
                    var start = markup.X;
                    for (var i = 1; i < this.DragTablePos; i++)
                        start += markup.Cols[i - 1];
                    _left = start;

                    start += markup.Margins[this.DragTablePos - 1].Left;
                    start += markup.Margins[this.DragTablePos - 1].Right;

                    _min = start;
                }

                if (newVal < _min)
                    newVal = _min;
                if (newVal > _max)
                    newVal = _max;

                if (0 == this.DragTablePos)
                {
                    markup.X = newVal;
                }
                else
                {
                    markup.Cols[this.DragTablePos - 1] = newVal - _left;
                }

                this.CalculateMargins();
                word_control.UpdateHorRulerBack();

                var pos = left + newVal * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.VertLine(pos);
                break;
            }
            case AscWord.RULER_DRAG_TYPE.columnSize:
            {
                var newVal = RulerCorrectPosition(this, _x, this.TableMarginLeftTrackStart);

                // сначала определим граничные условия
                var markup = this.m_oColumnMarkup;

                if (markup.EqualWidth)
                {
                    if (0 == this.DragTablePos)
                    {
                        var _min = 0;
                        var _max = markup.R - markup.Num * 10 - (markup.Num - 1) * markup.Space;

                        if (newVal < _min)
                            newVal = _min;
                        if (newVal > _max)
                            newVal = _max;

                        markup.X = newVal;
                    }
                    else if ((2 * markup.Num - 1) == this.DragTablePos)
                    {
                        var _min = markup.X + markup.Num * 10 + (markup.Num - 1) * markup.Space;
                        var _max = this.m_oPage.width_mm;

                        if (newVal < _min)
                            newVal = _min;
                        if (newVal > _max)
                            newVal = _max;

                        markup.R = newVal;
                    }
                    else
                    {
                        var bIsLeftSpace = ((this.DragTablePos & 1) == 1);
                        var _spaceMax = (markup.R - markup.X - 10 * markup.Num) / (markup.Num - 1);

                        var _col = ((this.DragTablePos + 1) >> 1);
                        var _center = _col * (markup.R - markup.X + markup.Space) / markup.Num - markup.Space / 2;

                        newVal -= markup.X;

                        if (bIsLeftSpace)
                        {
                            var _min = _center - _spaceMax / 2;
                            var _max = _center;

                            if (newVal < _min)
                                newVal = _min;
                            if (newVal > _max)
                                newVal = _max;

                            markup.Space = Math.abs((newVal - _center) * 2);
                        }
                        else
                        {
                            var _min = _center;
                            var _max = _center + _spaceMax / 2;

                            if (newVal < _min)
                                newVal = _min;
                            if (newVal > _max)
                                newVal = _max;

                            markup.Space = Math.abs((newVal - _center) * 2);
                        }

                        newVal += markup.X;
                    }
                }
                else
                {
                    var bIsLeftSpace = ((this.DragTablePos & 1) == 1);
                    var nSpaceNumber = ((this.DragTablePos + 1) >> 1);
                    var _min = 0;
                    var _max = this.m_oPage.width_mm;

                    if (0 == nSpaceNumber)
                    {
                        _max = markup.X + markup.Cols[0].W - 10;

                        if (newVal < _min)
                            newVal = _min;
                        if (newVal > _max)
                            newVal = _max;

                        var _delta = markup.X - newVal;
                        markup.X -= _delta;
                        markup.Cols[0].W += _delta;
                    }
                    else
                    {
                        var _offsetX = markup.X;
                        for (var i = 0; i < nSpaceNumber; i++)
                        {
                            _min = _offsetX;

                            _offsetX += markup.Cols[i].W;
                            if (!bIsLeftSpace || i != (nSpaceNumber - 1))
                            {
                                _min = _offsetX;
                                _offsetX += markup.Cols[i].Space;
                            }
                        }

                        if (bIsLeftSpace)
                        {
                            if (nSpaceNumber != markup.Num)
                                _max = _min + markup.Cols[nSpaceNumber - 1].W + markup.Cols[nSpaceNumber - 1].Space;

                            var _natMin = _min + 10;

                            if (newVal < _natMin)
                                newVal = _natMin;
                            if (newVal > _max)
                                newVal = _max;

                            markup.Cols[nSpaceNumber - 1].W = newVal - _min;
                            markup.Cols[nSpaceNumber - 1].Space = _max - newVal;

                            if (nSpaceNumber == markup.Num)
                            {
                                markup.R = newVal;
                            }
                        }
                        else
                        {
                            _max = _min + markup.Cols[nSpaceNumber - 1].Space + markup.Cols[nSpaceNumber].W;
                            var _natMax = _max - 10;

                            if (newVal < _min)
                                newVal = _min;
                            if (newVal > _natMax)
                                newVal = _natMax;

                            markup.Cols[nSpaceNumber - 1].Space = newVal - _min;
                            markup.Cols[nSpaceNumber].W = _max - newVal;
                        }
                    }
                }

                this.CalculateMargins();
                word_control.UpdateHorRulerBack();

                var pos = left + newVal * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.VertLine(pos);
                break;
            }
            case AscWord.RULER_DRAG_TYPE.columnPos:
            {
                var newVal = RulerCorrectPosition(this, _x, this.TableMarginLeftTrackStart);

                // сначала определим граничные условия
                var markup = this.m_oColumnMarkup;

                var _min = markup.X;
                for (var i = 0; i < this.DragTablePos; i++)
                {
                    _min += (markup.Cols[i].W + markup.Cols[i].Space);
                }

                var _max = _min + markup.Cols[this.DragTablePos].W + markup.Cols[this.DragTablePos].Space;
                _max += markup.Cols[this.DragTablePos + 1].W;

                var _space = markup.Cols[this.DragTablePos].Space;

                var _natMin = _min + _space / 2 + 10;
                var _natMax = _max - _space / 2 - 10;

                if (newVal < _natMin)
                    newVal = _natMin;
                if (newVal > _natMax)
                    newVal = _natMax;

                var _delta = newVal - (_min + markup.Cols[this.DragTablePos].W + markup.Cols[this.DragTablePos].Space / 2);
                markup.Cols[this.DragTablePos].W += _delta;
                markup.Cols[this.DragTablePos + 1].W -= _delta;

                this.CalculateMargins();
                word_control.UpdateHorRulerBack();

                var pos = left + newVal * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.VertLine(pos);
                break;
            }
        }
    }

    this.CheckMouseType = function(x, y, isMouseDown, isNegative)
    {
		// проверяем где находимся

        var _top = 1.8;
        var _bottom = 5.2;

        var _margin_left = this.m_dMarginLeft;
        var _margin_right = this.m_dMarginRight;
        if (this.CurrentObjectType == RULER_OBJECT_TYPE_TABLE || this.CurrentObjectType == RULER_OBJECT_TYPE_COLUMNS)
        {
            _margin_left = this.TableMarginLeft;
            _margin_right = this.TableMarginRight;
        }

        var posL = _margin_left;
        if ((_margin_left + this.m_dIndentLeft) > posL)
            posL = _margin_left + this.m_dIndentLeft;
        if ((_margin_left + this.m_dIndentLeftFirst) > posL)
            posL = _margin_left + this.m_dIndentLeftFirst;

        var posR = _margin_right;
        if (this.m_dIndentRight > 0)
            posR = _margin_right - this.m_dIndentRight;

        if (this.IsCanMoveAnyMarkers && posL < posR)
        {
            // tabs
            if (y >= 3 && y <= _bottom)
            {
                for (let i = 0, tabCount = this.m_arrTabs.length; i < tabCount; ++i)
                {
                    var _pos = this.m_bRtl ? _margin_right - this.m_arrTabs[i].pos : _margin_left + this.m_arrTabs[i].pos;
                    if ((x >= (_pos - 1)) && (x <= (_pos + 1)))
                    {
                        if (true === isMouseDown)
                            this.m_lCurrentTab = i;
                        return AscWord.RULER_DRAG_TYPE.tab;
                    }
                }
            }

            // left indent
            var dCenterX = this.m_bRtl ? _margin_right - this.m_dIndentLeft : _margin_left +  this.m_dIndentLeft;

            var var1 = dCenterX - 1;
            var var2 = 1.4;
            var var3 = 1.5;
            var var4 = dCenterX + 1;

            if ((x >= var1) && (x <= var4))
            {
                if ((y >= _bottom) && (y < (_bottom + var2)))
                    return AscWord.RULER_DRAG_TYPE.leftFirstInd;
                else if ((y > (_bottom - var3)) && (y < _bottom))
                    return AscWord.RULER_DRAG_TYPE.leftInd;
            }

            // right indent
            dCenterX = this.m_bRtl ? _margin_left + this.m_dIndentRight : _margin_right -  this.m_dIndentRight;

            var1 = dCenterX - 1;
            var4 = dCenterX + 1;

            if ((x >= var1) && (x <= var4))
            {
                if ((y > (_bottom - var3)) && (y < _bottom))
                    return AscWord.RULER_DRAG_TYPE.rightInd;
            }

            // first line indent
            dCenterX = this.m_bRtl ? _margin_right - this.m_dIndentLeftFirst : _margin_left +  this.m_dIndentLeftFirst;

            var1 = dCenterX - 1;
            var4 = dCenterX + 1;

            if ((x >= var1) && (x <= var4))
            {
                if ((y > (_top - 1)) && (y < (_top + 1.68)))
                {
                    if (0 == this.m_dIndentLeftFirst && 0 == this.m_dIndentLeft && this.CurrentObjectType == RULER_OBJECT_TYPE_PARAGRAPH && this.IsCanMoveMargins)
                    {
                        if (y > (_top + 1))
                            return this.m_bRtl ? AscWord.RULER_DRAG_TYPE.rightMargin : AscWord.RULER_DRAG_TYPE.leftMargin;
                    }
                    return AscWord.RULER_DRAG_TYPE.firstInd;
                }
            }
        }

        var isColumnsInside = false;
		var isColumnsInside2 = false;
        var isTableInside = false;

        if (this.CurrentObjectType == RULER_OBJECT_TYPE_PARAGRAPH && this.IsCanMoveMargins)
        {
            if (y >= _top && y <= _bottom)
            {
				// внутри линейки
                if (Math.abs(x - this.m_dMarginLeft) < 1)
                {
                    return AscWord.RULER_DRAG_TYPE.leftMargin;
                }
                else if (Math.abs(x - this.m_dMarginRight) < 1)
                {
                    return AscWord.RULER_DRAG_TYPE.rightMargin;
                }

                if (isNegative)
                {
                    if (x < this.m_dMarginLeft)
                        return -AscWord.RULER_DRAG_TYPE.leftMargin;
                    if (x > this.m_dMarginRight)
                        return -AscWord.RULER_DRAG_TYPE.rightMargin;
                }
            }
        }
        else if (this.CurrentObjectType == RULER_OBJECT_TYPE_TABLE)
        {
            var nTI_x = 0;
            var nTI_r = 0;
            if (y >= _top && y <= _bottom)
            {
                var markup = this.m_oTableMarkup;
                var pos = markup.X;
                var _count = markup.Cols.length;
                for (var i = 0; i <= _count; i++)
                {
                    if (Math.abs(x - pos) < 1)
                    {
                        this.DragTablePos = i;
                        return AscWord.RULER_DRAG_TYPE.table;
                    }
                    if (i == _count)
                    {
                        nTI_r = pos;
                        break;
					}
					if (i == 0)
                    {
                        nTI_x = pos;
                    }
                    
                    pos += markup.Cols[i];
                }
            }

            if (x > nTI_x && x < nTI_r)
                isTableInside = true;
        }
        else if (this.CurrentObjectType == RULER_OBJECT_TYPE_COLUMNS)
        {
			if (y >= _top && y <= _bottom)
            {
                var markup = this.m_oColumnMarkup;

				var nCI_x = markup.X;
				var nCI_r = nCI_x;

                if (markup.EqualWidth)
                {
                    var _w = ((markup.R - markup.X) - markup.Space * (markup.Num - 1)) / markup.Num;
                    var _x = markup.X;

                    var _index = 0;
                    for (var i = 0; i < markup.Num; i++)
                    {
                        if (0 == i)
                        {
                            if (Math.abs(x - _x) < 1)
                            {
                                this.DragTablePos = _index;
                                return AscWord.RULER_DRAG_TYPE.columnSize;
                            }
                        }
                        else
                        {
                            if (x < _x + 1 && x > _x - 2)
                            {
                                this.DragTablePos = _index;
                                return AscWord.RULER_DRAG_TYPE.columnSize;
                            }
                        }

                        ++_index;
                        _x += _w;
                        nCI_r = _x;

                        if (i == markup.Num - 1)
                        {
                            if (Math.abs(x - _x) < 1)
                            {
                                this.DragTablePos = _index;
                                return AscWord.RULER_DRAG_TYPE.columnSize;
                            }
                        }
                        else
                        {
                            if (x < _x + 2 && x > _x - 1)
                            {
                                this.DragTablePos = _index;
                                return AscWord.RULER_DRAG_TYPE.columnSize;
                            }

							if (x > _x && x < (_x + markup.Space))
								isColumnsInside = true;
                        }

                        ++_index;
                        _x += markup.Space;
                    }
                }
                else
                {
                    var _x = markup.X;

                    var _index = 0;
                    for (var i = 0; i < markup.Cols.length; i++)
                    {
                        if (0 == i)
                        {
                            if (Math.abs(x - _x) < 1)
                            {
                                this.DragTablePos = _index;
                                return AscWord.RULER_DRAG_TYPE.columnSize;
                            }
                        }
                        else
                        {
                            if (x < _x + 1 && x > _x - 2)
                            {
                                this.DragTablePos = _index;
                                return AscWord.RULER_DRAG_TYPE.columnSize;
                            }
                        }

                        ++_index;
                        _x += markup.Cols[i].W;
						nCI_r = _x;

						if (i == markup.Num - 1)
                        {
                            if (Math.abs(x - _x) < 1)
                            {
                                this.DragTablePos = _index;
                                return AscWord.RULER_DRAG_TYPE.columnSize;
                            }
                        }
                        else
                        {
                            if (x < _x + 2 && x > _x - 1)
                            {
                                this.DragTablePos = _index;
                                return AscWord.RULER_DRAG_TYPE.columnSize;
                            }
                        }

                        if (i != markup.Cols.length - 1)
                        {
                            if (Math.abs(x - (_x + markup.Cols[i].Space / 2)) < 1)
                            {
                                this.DragTablePos = i;
                                return AscWord.RULER_DRAG_TYPE.columnPos;
                            }

							if (x > _x && x < (_x + markup.Cols[i].Space))
								isColumnsInside = true;
                        }

                        ++_index;
                        _x += markup.Cols[i].Space;
                    }
                }

				if (x > nCI_x && x < nCI_r)
					isColumnsInside2 = true;
            }
        }

        if (isNegative)
        {
            if (isColumnsInside)
                return -AscWord.RULER_DRAG_TYPE.columnSize;

            // если вникуда - то ВСЕГДА margins
            return -AscWord.RULER_DRAG_TYPE.leftMargin;

            if (isColumnsInside2)
                return 0;
            if ((y >= _top && y <= _bottom) && !isTableInside)
            {
				if (x < _margin_left)
					return -AscWord.RULER_DRAG_TYPE.leftMargin;
				if (x > _margin_right)
					return -AscWord.RULER_DRAG_TYPE.rightMargin;
            }
        }

        return AscWord.RULER_DRAG_TYPE.none;
    }

    this.OnMouseDown = function(left, top, e)
    {
        var word_control = this.m_oWordControl;

        if (true === word_control.m_oApi.isStartAddShape)
        {
			word_control.m_oApi.sync_EndAddShape();
			word_control.m_oApi.sync_StartAddShapeCallback(false);
        }
        if (true === word_control.m_oApi.isInkDrawerOn())
        {
			word_control.m_oApi.stopInkDrawer();
        }

        AscCommon.check_MouseDownEvent(e, true);
        global_mouseEvent.LockMouse();

        this.SimpleChanges.Reinit();

        var dKoefPxToMM = 100 * g_dKoef_pix_to_mm / word_control.m_nZoomValue;
        var dKoef_mm_to_pix = g_dKoef_mm_to_pix * this.m_dZoom;
        
        var _x = global_mouseEvent.X - left - word_control.X - word_control.GetMainContentBounds().L * g_dKoef_mm_to_pix;
        if (!word_control.m_oApi.isRtlInterface)
            _x -= 5 * g_dKoef_mm_to_pix;

        _x *= dKoefPxToMM;
        var _y = (global_mouseEvent.Y - word_control.Y) * g_dKoef_pix_to_mm;

        this.DragType = this.CheckMouseType(_x, _y, true, true);
        if (this.DragType < 0)
        {
			this.DragTypeMouseDown = -this.DragType;
			this.DragType = AscWord.RULER_DRAG_TYPE.none;
		}
		else
        {
            this.DragTypeMouseDown = this.DragType;
		}

		if (global_mouseEvent.ClickCount > 1)
        {
            var eventType = "";
			switch (this.DragTypeMouseDown)
			{
				case AscWord.RULER_DRAG_TYPE.leftMargin:
				case AscWord.RULER_DRAG_TYPE.rightMargin:
				{
					eventType = "margins";
					break;
				}
				case AscWord.RULER_DRAG_TYPE.leftFirstInd:
				case AscWord.RULER_DRAG_TYPE.leftInd:
				case AscWord.RULER_DRAG_TYPE.firstInd:
				case AscWord.RULER_DRAG_TYPE.rightInd:
				{
					eventType = "indents";
					break;
				}
				case AscWord.RULER_DRAG_TYPE.tab:
				{
					eventType = "tabs";
					break;
				}
				case AscWord.RULER_DRAG_TYPE.table:
				{
					eventType = "tables";
					break;
				}
				case AscWord.RULER_DRAG_TYPE.columnSize:
				case AscWord.RULER_DRAG_TYPE.columnPos:
				{
					eventType = "columns";
					break;
				}
			}

            if (eventType != "")
            {
				word_control.m_oApi.sendEvent("asc_onRulerDblClick", eventType);
                this.DragType = AscWord.RULER_DRAG_TYPE.none;
                this.OnMouseUp(left, top, e);
                return;
            }
        }

        var _margin_left = this.m_dMarginLeft;
        var _margin_right = this.m_dMarginRight;
        if (this.CurrentObjectType == RULER_OBJECT_TYPE_TABLE || this.CurrentObjectType == RULER_OBJECT_TYPE_COLUMNS)
        {
            _margin_left = this.TableMarginLeft;
            _margin_right = this.TableMarginRight;
        }

        this.m_bIsMouseDown = true;

        switch (this.DragType)
        {
            case AscWord.RULER_DRAG_TYPE.leftMargin:
            {
                var pos = left + _margin_left * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.VertLine(pos);
                break;
            }
            case AscWord.RULER_DRAG_TYPE.rightMargin:
            {
                var pos = left + _margin_right * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.VertLine(pos);
                break;
            }
			case AscWord.RULER_DRAG_TYPE.leftFirstInd:
			{
				let offset = this.m_bRtl ? (_margin_right - this.m_dIndentLeft) : (_margin_left + this.m_dIndentLeft);
				let pos    = left + offset * dKoef_mm_to_pix;
				word_control.m_oOverlayApi.VertLine(pos);
				
				this.m_dIndentLeft_old      = this.m_dIndentLeft;
				this.m_dIndentLeftFirst_old = this.m_dIndentLeftFirst;
				break;
			}
			case AscWord.RULER_DRAG_TYPE.leftInd:
			{
				let offset = this.m_bRtl ? (_margin_right - this.m_dIndentLeft) : (_margin_left + this.m_dIndentLeft);
				let pos    = left + offset * dKoef_mm_to_pix;
				word_control.m_oOverlayApi.VertLine(pos);
				
				this.m_dIndentLeft_old = this.m_dIndentLeft;
				break;
			}
			case AscWord.RULER_DRAG_TYPE.firstInd:
			{
				let offset = this.m_bRtl ? (_margin_right - this.m_dIndentLeftFirst) : (_margin_left + this.m_dIndentLeftFirst);
				let pos    = left + offset * dKoef_mm_to_pix;
				word_control.m_oOverlayApi.VertLine(pos);
				
				this.m_dIndentLeftFirst_old = this.m_dIndentLeftFirst;
				break;
			}
			case AscWord.RULER_DRAG_TYPE.rightInd:
			{
				let offset = this.m_bRtl ? (_margin_left + this.m_dIndentRight) : (_margin_right - this.m_dIndentRight);
				let pos    = left + offset * dKoef_mm_to_pix;
				word_control.m_oOverlayApi.VertLine(pos);
				
				this.m_dIndentRight_old = this.m_dIndentRight;
				break;
			}
			case AscWord.RULER_DRAG_TYPE.tab:
			{
				let offset = this.m_bRtl ? (_margin_right - this.m_arrTabs[this.m_lCurrentTab].pos) : (_margin_left + this.m_arrTabs[this.m_lCurrentTab].pos);
				let pos    = left + offset * dKoef_mm_to_pix;
				word_control.m_oOverlayApi.VertLine(pos);
				this.m_dCurrentTabNewPosition = this.m_arrTabs[this.m_lCurrentTab].pos;
				break;
			}
            case AscWord.RULER_DRAG_TYPE.table:
            {
                var markup = this.m_oTableMarkup;
                var pos = markup.X;
                var _count = markup.Cols.length;
                for (var i = 0; i < this.DragTablePos; i++)
                {
                    pos += markup.Cols[i];
                }
                pos = left + pos * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.VertLine(pos);

                this.TableMarginLeftTrackStart = this.TableMarginLeft;
                break;
            }
            case AscWord.RULER_DRAG_TYPE.columnSize:
            {
                var markup = this.m_oColumnMarkup;
                var pos = 0;
                if (markup.EqualWidth)
                {
                    var _w = ((markup.R - markup.X) - markup.Space * (markup.Num - 1)) / markup.Num;
                    var _x = markup.X + (this.DragTablePos >> 1) * (_w + markup.Space);
                    if (this.DragTablePos & 1 == 1)
                        _x += _w;

                    pos = _x;
                }
                else
                {
                    var _x = markup.X;

                    var _index = 0;
                    for (var i = 0; i < markup.Cols.length && _index < this.DragTablePos; i++)
                    {
                        if (_index == this.DragTablePos)
                            break;

                        ++_index;
                        _x += markup.Cols[i].W;

                        if (_index == this.DragTablePos)
                            break;

                        ++_index;
                        _x += markup.Cols[i].Space;
                    }

                    pos = _x;
                }

                pos = left + pos * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.VertLine(pos);

                this.TableMarginLeftTrackStart = markup.X;
                break;
            }
            case AscWord.RULER_DRAG_TYPE.columnPos:
            {
                var markup = this.m_oColumnMarkup;
                var pos = markup.X;

                var _index = 0;
                for (var i = 0; i < markup.Cols.length && i < this.DragTablePos; i++)
                {
                    if (_index == this.DragTablePos)
                        break;

                    pos += markup.Cols[i].W;
                    pos += markup.Cols[i].Space;
                }

                if (this.DragTablePos < markup.Cols.length)
                {
                    pos += markup.Cols[this.DragTablePos].W;
                    pos += markup.Cols[this.DragTablePos].Space / 2;
                }

                pos = left + pos * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.VertLine(pos);
                this.TableMarginLeftTrackStart = markup.X;
                break;
            }
        }
		
		if (AscWord.RULER_DRAG_TYPE.none === this.DragType)
		{
			let posT = 3;
			let posB = 5.2;
			let posL = this.m_bRtl ? _margin_left + this.m_dIndentRight : _margin_left + this.m_dIndentLeft;
			let posR = this.m_bRtl ? _margin_right - this.m_dIndentRight : _margin_right - this.m_dIndentLeft;
			if (posT <= _y && _y <= posB && posL <= _x && _x <= posR)
			{
				let newVal  = RulerCorrectPosition(this, _x, this.m_bRtl ? _margin_right : _margin_left);
				let tabPos  = this.m_bRtl ? _margin_right - newVal : newVal - _margin_left;
				let tabType = word_control.m_nTabsType;
				if (this.m_bRtl)
				{
					if (tab_Left === tabType)
						tabType = tab_Right;
					else if (tab_Right === tabType)
						tabType = tab_Left;
				}
				this.m_arrTabs[this.m_arrTabs.length] = new CTab(tabPos, tabType);
				word_control.UpdateHorRuler();
				
				this.m_lCurrentTab = this.m_arrTabs.length - 1;
				
				this.DragType                 = AscWord.RULER_DRAG_TYPE.tab;
				this.m_dCurrentTabNewPosition = tabPos;
				
				let pos = left + newVal * dKoef_mm_to_pix;
				word_control.m_oOverlayApi.VertLine(pos);
			}
		}
		
		word_control.m_oDrawingDocument.LockCursorTypeCur();
	}
    this.OnMouseUp = function(left, top, e)
    {
        var word_control = this.m_oWordControl;
        this.m_oWordControl.OnUpdateOverlay();
        var lockedElement = AscCommon.check_MouseUpEvent(e);

        this.m_dIndentLeft_old      = -10000;
        this.m_dIndentLeftFirst_old = -10000;
        this.m_dIndentRight_old     = -10000;

        if (AscWord.RULER_DRAG_TYPE.tab !== this.DragType)
        {
            word_control.UpdateHorRuler();
            //word_control.m_oOverlayApi.UnShow();
        }

        var _margin_left = this.m_dMarginLeft;
        var _margin_right = this.m_dMarginRight;
        if (this.CurrentObjectType == RULER_OBJECT_TYPE_TABLE || this.CurrentObjectType == RULER_OBJECT_TYPE_COLUMNS)
        {
            _margin_left = this.TableMarginLeft;
            _margin_right = this.TableMarginRight;
        }

        switch (this.DragType)
        {
			case AscWord.RULER_DRAG_TYPE.leftMargin:
			case AscWord.RULER_DRAG_TYPE.rightMargin:
            {
                if (!this.SimpleChanges.IsSimple)
                    this.SetMarginProperties();
                break;
            }
			case AscWord.RULER_DRAG_TYPE.leftFirstInd:
			case AscWord.RULER_DRAG_TYPE.leftInd:
            case AscWord.RULER_DRAG_TYPE.firstInd:
            case AscWord.RULER_DRAG_TYPE.rightInd:
            {
                if (!this.SimpleChanges.IsSimple)
                    this.SetPrProperties();
                else
                    word_control.OnUpdateOverlay();
                break;
            }
            case AscWord.RULER_DRAG_TYPE.tab:
            {
                // смотрим, сохраняем ли таб
				var _y = (global_mouseEvent.Y - word_control.Y) * g_dKoef_pix_to_mm;
				if (_y <= 3 || _y > 5.6 || this.m_dCurrentTabNewPosition < this.m_dIndentLeft || this.m_dCurrentTabNewPosition > _margin_right - _margin_left - this.m_dIndentRight)
				{
					if (-1 !== this.m_lCurrentTab)
						this.m_arrTabs.splice(this.m_lCurrentTab, 1);
				}
				else
				{
					if (this.m_lCurrentTab < this.m_arrTabs.length)
						this.m_arrTabs[this.m_lCurrentTab].pos = this.m_dCurrentTabNewPosition;
				}

                this.m_lCurrentTab = -1;
                this.CorrectTabs();
                this.m_oWordControl.UpdateHorRuler();
                this.SetTabsProperties();
                break;
            }
            case AscWord.RULER_DRAG_TYPE.table:
            {
                if (!this.SimpleChanges.IsSimple)
                    this.SetTableProperties();
                this.DragTablePos = -1;
                break;
            }
            case AscWord.RULER_DRAG_TYPE.columnSize:
			case AscWord.RULER_DRAG_TYPE.columnPos:
            {
                if (!this.SimpleChanges.IsSimple)
                    this.SetColumnsProperties();
                this.DragTablePos = -1;
                break;
            }
        }

        if (AscWord.RULER_DRAG_TYPE.tab === this.DragType)
        {
            word_control.UpdateHorRuler();
            //word_control.m_oOverlayApi.UnShow();
        }

        this.IsDrawingCurTab = true;
        this.DragType = 0;
        this.m_bIsMouseDown = false;

        this.m_oWordControl.m_oDrawingDocument.UnlockCursorType();
        this.SimpleChanges.Clear();
    }

    this.OnMouseUpExternal = function()
    {
        var word_control = this.m_oWordControl;
        this.m_oWordControl.OnUpdateOverlay();

        this.m_dIndentLeft_old      = -10000;
        this.m_dIndentLeftFirst_old = -10000;
        this.m_dIndentRight_old     = -10000;

        if (AscWord.RULER_DRAG_TYPE.tab !== this.DragType)
        {
            word_control.UpdateHorRuler();
            //word_control.m_oOverlayApi.UnShow();
        }

        var _margin_left = this.m_dMarginLeft;
        var _margin_right = this.m_dMarginRight;
        if (this.CurrentObjectType == RULER_OBJECT_TYPE_TABLE || this.CurrentObjectType == RULER_OBJECT_TYPE_COLUMNS)
        {
            _margin_left = this.TableMarginLeft;
            _margin_right = this.TableMarginRight;
        }

        switch (this.DragType)
        {
            case AscWord.RULER_DRAG_TYPE.leftMargin:
            case AscWord.RULER_DRAG_TYPE.rightMargin:
            {
                if (!this.SimpleChanges.IsSimple)
                    this.SetMarginProperties();
                break;
            }
            case AscWord.RULER_DRAG_TYPE.leftFirstInd:
            case AscWord.RULER_DRAG_TYPE.leftInd:
            case AscWord.RULER_DRAG_TYPE.firstInd:
            case AscWord.RULER_DRAG_TYPE.rightInd:
            {
                if (!this.SimpleChanges.IsSimple)
                    this.SetPrProperties();
                break;
            }
            case AscWord.RULER_DRAG_TYPE.tab:
            {
                // смотрим, сохраняем ли таб
                var _y = (global_mouseEvent.Y - word_control.Y) * g_dKoef_pix_to_mm;
                if (_y <= 3 || _y > 5.6 || this.m_dCurrentTabNewPosition < this.m_dIndentLeft || this.m_dCurrentTabNewPosition > _margin_right - _margin_left - this.m_dIndentRight)
                {
					if (-1 !== this.m_lCurrentTab)
						this.m_arrTabs.splice(this.m_lCurrentTab, 1);
                }
                else
                {
					if (this.m_lCurrentTab < this.m_arrTabs.length)
						this.m_arrTabs[this.m_lCurrentTab].pos = this.m_dCurrentTabNewPosition;
                }
                this.m_lCurrentTab = -1;
                this.CorrectTabs();
                this.m_oWordControl.UpdateHorRuler();
                this.SetTabsProperties();
                break;
            }
            case AscWord.RULER_DRAG_TYPE.table:
            {
                if (!this.SimpleChanges.IsSimple)
                    this.SetTableProperties();
                this.DragTablePos = -1;
                break;
            }
            case AscWord.RULER_DRAG_TYPE.columnSize:
            case AscWord.RULER_DRAG_TYPE.columnPos:
            {
                if (!this.SimpleChanges.IsSimple)
                    this.SetColumnsProperties();
                this.DragTablePos = -1;
                break;
            }
        }

        if (AscWord.RULER_DRAG_TYPE.tab === this.DragType)
        {
            word_control.UpdateHorRuler();
            //word_control.m_oOverlayApi.UnShow();
        }

        this.IsDrawingCurTab = true;
        this.DragType = 0;
        this.m_bIsMouseDown = false;

        this.m_oWordControl.m_oDrawingDocument.UnlockCursorType();
        this.SimpleChanges.Clear();
    }

    this.SetTabsProperties = function()
	{
		// потом заменить на объекты CTab (когда Илюха реализует не только левые табы)
		var _arr = new CParaTabs();
		var _c   = this.m_arrTabs.length;
		for (var i = 0; i < _c; i++)
		{
			if (this.m_arrTabs[i].type == tab_Left || this.m_arrTabs[i].type == tab_Right || this.m_arrTabs[i].type == tab_Center)
				_arr.Add(new CParaTab(this.m_arrTabs[i].type, this.m_arrTabs[i].pos, this.m_arrTabs[i].leader));
		}

		if (false === this.m_oWordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Properties))
		{
			this.m_oWordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphTabs);
			this.m_oWordControl.m_oLogicDocument.SetParagraphTabs(_arr);
			this.m_oWordControl.m_oLogicDocument.FinalizeAction();
		}
		this.m_oWordControl.m_oLogicDocument.UpdateSelection();
		this.m_oWordControl.m_oLogicDocument.UpdateInterface();
	}

    this.SetPrProperties = function(isTemporary)
	{
		let left      = this.m_dIndentLeft;
		let right     = this.m_dIndentRight;
		let firstLine = this.m_dIndentLeftFirst - this.m_dIndentLeft;
		
		let logicDocument = this.m_oWordControl.m_oLogicDocument;
		if (logicDocument.IsSelectionLocked(AscCommon.changestype_Paragraph_Properties))
		{
			logicDocument.UpdateSelection();
			logicDocument.UpdateInterface();
			return;
		}
		
		isTemporary = isTemporary && logicDocument.IsDocumentEditor();
		
		if (isTemporary)
			logicDocument.TurnOff_InterfaceEvents();
		
		logicDocument.StartAction(AscDFH.historydescription_Document_SetParagraphIndentFromRulers);
		logicDocument.SetParagraphIndent({
			Left      : left,
			Right     : right,
			FirstLine : firstLine
		});
		logicDocument.UpdateInterface();
		logicDocument.Recalculate();
		logicDocument.FinalizeAction();
		
		if (isTemporary)
		{
			AscCommon.History.SetLastPointTemporary();
			logicDocument.TurnOn_InterfaceEvents();
		}
	}
    this.SetMarginProperties = function()
    {
        if ( false === this.m_oWordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Document_SectPr) )
        {
            this.m_oWordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetDocumentMargin_Hor);
            this.m_oWordControl.m_oLogicDocument.SetDocumentMargin( { Left : this.m_dMarginLeft, Right : this.m_dMarginRight }, true);
			this.m_oWordControl.m_oLogicDocument.FinalizeAction();
        }
		this.m_oWordControl.m_oLogicDocument.UpdateSelection();
		this.m_oWordControl.m_oLogicDocument.UpdateInterface();
    }

    this.SetTableProperties = function()
    {
        if ( false === this.m_oWordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Table_Properties) )
        {
            this.m_oWordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTableMarkup_Hor);

            this.m_oTableMarkup.CorrectTo();
            this.m_oTableMarkup.Table.Update_TableMarkupFromRuler(this.m_oTableMarkup, true, this.DragTablePos);
			if (this.m_oTableMarkup)
			    this.m_oTableMarkup.CorrectFrom();

            this.m_oWordControl.m_oLogicDocument.UpdateRulers();
			this.m_oWordControl.m_oLogicDocument.FinalizeAction();
        }
		this.m_oWordControl.m_oLogicDocument.UpdateSelection();
		this.m_oWordControl.m_oLogicDocument.UpdateInterface();
	}

    this.SetColumnsProperties = function()
    {
        this.m_oWordControl.m_oLogicDocument.Update_ColumnsMarkupFromRuler(this.m_oColumnMarkup);
    }

    this.BlitToMain = function(left, top, htmlElement)
    {
        var dPR = AscCommon.AscBrowser.retinaPixelRatio;
        left = Math.round(dPR * left);
        var _margin_left = this.m_dMarginLeft;
        var _margin_right = this.m_dMarginRight;
        if (this.CurrentObjectType == RULER_OBJECT_TYPE_TABLE || this.CurrentObjectType == RULER_OBJECT_TYPE_COLUMNS)
        {
            _margin_left = this.TableMarginLeft;
            _margin_right = this.TableMarginRight;
        }

        var checker = this.RepaintChecker;
        if (!checker.BlitAttack && left == checker.BlitLeft && !this.m_bIsMouseDown)
        {
            if (checker.BlitIndentLeft == this.m_dIndentLeft
				&& checker.BlitIndentLeftFirst == this.m_dIndentLeftFirst
                && checker.BlitIndentRight == this.m_dIndentRight
				&& checker.BlitRtl === this.m_bRtl
				&& checker.BlitDefaultTab == this.m_dDefaultTab
				&& _margin_left == checker.BlitMarginLeftInd
				&& _margin_right == checker.BlitMarginRightInd)
            {
                // осталось проверить только табы кастомные
                var _count1 = 0;
                if (null != checker.BlitTabs)
                    _count1 = checker.BlitTabs.length;

                var _count2 =  this.m_arrTabs.length;
                if (_count1 == _count2)
                {
                    var bIsBreak = false;
                    for (var ii = 0; ii < _count1; ii++)
                    {
                        if ((checker.BlitTabs[ii].type != this.m_arrTabs[ii].type) || (checker.BlitTabs[ii].pos != this.m_arrTabs[ii].pos))
                        {
                            bIsBreak = true;
                            break;
                        }
                    }
                    if (false === bIsBreak)
                        return;
                }
            }
        }

        checker.BlitAttack = false;
        htmlElement.width = htmlElement.width;
        var context = htmlElement.getContext('2d');
        context.setTransform(1, 0, 0, 1, 0, 0);

        if (null != this.m_oCanvas)
        {
            checker.BlitLeft = left;
			checker.BlitRtl = this.m_bRtl;
            checker.BlitIndentLeft = this.m_dIndentLeft;
            checker.BlitIndentLeftFirst = this.m_dIndentLeftFirst;
            checker.BlitIndentRight = this.m_dIndentRight;
            checker.BlitDefaultTab = this.m_dDefaultTab;
            checker.BlitTabs = null;
            if (0 != this.m_arrTabs.length)
            {
                checker.BlitTabs = [];
                var _len = this.m_arrTabs.length;
                for (var ii = 0; ii < _len; ii++)
                {
                    checker.BlitTabs[ii] = { type: this.m_arrTabs[ii].type, pos: this.m_arrTabs[ii].pos };
                }
            }
            var roundDPR = Math.round(dPR);
            var indent = 0.5 * roundDPR;
            //context.drawImage(this.m_oCanvas, left - 5, 0, this.m_oCanvas.width, this.m_oCanvas.height,
            //    0, 0, this.m_oCanvas.width, this.m_oCanvas.height);

            context.drawImage(this.m_oCanvas, 5 * roundDPR, 0, this.m_oCanvas.width - 10 * roundDPR, this.m_oCanvas.height,
                    left, 0, this.m_oCanvas.width - 10 * roundDPR, this.m_oCanvas.height);

            if (!this.IsDrawAnyMarkers)
                return;
            var dKoef_mm_to_pix = g_dKoef_mm_to_pix * this.m_dZoom * dPR;
            var dCenterX = 0;
            var var2 = 0;
            var var3 = 0;

            // не менять!!!
            var2 = 5 * dPR;//(1.4 * g_dKoef_mm_to_pix) >> 0;
            var3 = 3 * dPR;//(1 * g_dKoef_mm_to_pix) >> 0;

            checker.BlitMarginLeftInd = _margin_left;
            checker.BlitMarginRightInd = _margin_right;

			var _1mm_to_pix = g_dKoef_mm_to_pix * dPR;
			let top    = this.m_nTop;
			let bottom = this.m_nBottom;
			let isRtl  = this.m_bRtl;
			
			function blitLeftInd(ind, isRtl)
			{
				let offset   = isRtl ? (_margin_right - ind) : (_margin_left + ind);
				let dCenterX = left + offset * dKoef_mm_to_pix;
				
				let var1 = parseInt(dCenterX - _1mm_to_pix) - indent + Math.round(dPR) - 1;
				let var4 = parseInt(dCenterX + _1mm_to_pix) + indent + Math.round(dPR) - 1;
				
				if (0 !== ((var1 - var4 + Math.round(dPR)) & 1))
					var4 += 1;
				
				context.beginPath();
				context.lineWidth = Math.round(dPR);
				context.moveTo(var1, bottom + indent);
				context.lineTo(var4, bottom + indent);
				context.lineTo(var4, bottom + indent + Math.round(var2));
				context.lineTo(var1, bottom + indent + Math.round(var2));
				context.lineTo(var1, bottom + indent);
				context.lineTo(var1, bottom + indent - Math.round(var3));
				context.lineTo((var1 + var4) / 2, bottom - Math.round(var2 * 1.2));
				context.lineTo(var4, bottom + indent - Math.round(var3));
				context.lineTo(var4, bottom + indent);
				
				context.fill();
				context.stroke();
			}
			
			function blitFirstInd(ind, isRtl)
			{
				let offset = isRtl ? (_margin_right - ind) : (_margin_left + ind);
				dCenterX   = left + offset * dKoef_mm_to_pix;
				
				let var1 = parseInt(dCenterX - _1mm_to_pix) - indent + Math.round(dPR) - 1;
				let var4 = parseInt(dCenterX + _1mm_to_pix) + indent + Math.round(dPR) - 1;
				
				if (0 !== ((var1 - var4 + Math.round(dPR)) & 1))
					var4 += 1;
				
				// first line indent
				context.beginPath();
				context.lineWidth = Math.round(dPR);
				context.moveTo(var1, top + indent);
				context.lineTo(var1, top + indent - Math.round(var3));
				context.lineTo(var4, top + indent - Math.round(var3));
				context.lineTo(var4, top + indent);
				context.lineTo((var1 + var4) / 2, top + Math.round(var2 * 1.2));
				context.closePath();
				
				context.fill();
				context.stroke();
			}
			
			function blitRightInd(ind, isRtl)
			{
				let offset = isRtl ? (_margin_left + ind) : (_margin_right - ind);
				dCenterX   = left + offset * dKoef_mm_to_pix;
				
				let var1 = parseInt(dCenterX - _1mm_to_pix) - indent + Math.round(dPR) - 1;
				let var4 = parseInt(dCenterX + _1mm_to_pix) + indent + Math.round(dPR) - 1;
				
				if (0 !== ((var1 - var4 + Math.round(dPR)) & 1))
					var4 += 1;
				
				context.beginPath();
				context.lineWidth = Math.round(dPR);
				context.moveTo(var1, bottom + indent);
				context.lineTo(var4, bottom + indent);
				context.lineTo(var4, bottom + indent - Math.round(var3));
				context.lineTo((var1 + var4) / 2, bottom - Math.round(var2 * 1.2));
				context.lineTo(var1, bottom + indent - Math.round(var3));
				context.closePath();
				
				context.fill();
				context.stroke();
			}
			
			function blitLeftTab(x, y)
			{
				context.beginPath();
				context.moveTo(x, y);
				context.lineTo(x, y + Math.round(5 * dPR));
				context.lineTo(x + Math.round(5 * dPR), y + Math.round(5 * dPR));
				context.stroke();
			}
			
			function blitRightTab(x, y)
			{
				context.beginPath();
				context.moveTo(x, y);
				context.lineTo(x, y + Math.round(5 * dPR));
				context.lineTo(x - Math.round(5 * dPR), y + Math.round(5 * dPR));
				context.stroke();
			}
			
			function blitCenterTab(x, y)
			{
				context.beginPath();
				context.moveTo(x, y);
				context.lineTo(x, y + Math.round(5 * dPR));
				context.moveTo(x - Math.round(5 * dPR), y + Math.round(5 * dPR));
				context.lineTo(x + Math.round(5 * dPR), y + Math.round(5 * dPR));
				context.stroke();
			}
			
			function blitTab(tabPos, tabType)
			{
				let x = parseInt((isRtl ? (_margin_right - tabPos) : (_margin_left + tabPos)) * dKoef_mm_to_pix) + left;
				let y = bottom - 5 * dPR;
				
				let lineW = context.lineWidth;
				context.lineWidth = 2 * roundDPR;
				
				if (tab_Left === tabType)
				{
					if (isRtl)
						blitRightTab(x, y);
					else
						blitLeftTab(x, y);
				}
				else if (tab_Right === tabType)
				{
					if (isRtl)
						blitLeftTab(x, y);
					else
						blitRightTab(x, y);
				}
				else if (tab_Center === tab_Center)
				{
					blitCenterTab(x, y);
				}
				
				context.lineWidth = lineW;
			}
			
			// old position --------------------------------------
			context.strokeStyle = GlobalSkin.RulerMarkersOutlineColorOld;
			context.fillStyle   = GlobalSkin.RulerMarkersFillColorOld;
			if (-10000 !== this.m_dIndentLeft_old && this.m_dIndentLeft_old !== this.m_dIndentLeft)
				blitLeftInd(this.m_dIndentLeft_old, this.m_bRtl);
			
			if (-10000 !== this.m_dIndentLeftFirst_old && this.m_dIndentLeftFirst_old !== this.m_dIndentLeftFirst)
				blitFirstInd(this.m_dIndentLeftFirst_old, this.m_bRtl);
			
			if (-10000 !== this.m_dIndentRight_old && this.m_dIndentRight_old !== this.m_dIndentRight)
				blitRightInd(this.m_dIndentRight_old, this.m_bRtl);
			
			context.strokeStyle = GlobalSkin.RulerTabsColorOld;
			if (-1 !== this.m_lCurrentTab && this.m_lCurrentTab < this.m_arrTabs.length)
			{
				var _tab = this.m_arrTabs[this.m_lCurrentTab];
				blitTab(_tab.pos, _tab.type);
			}
            
            // ---------------------------------------------------

            // рисуем инденты, только если они корректны
            var posL = _margin_left;
            if ((_margin_left + this.m_dIndentLeft) > posL)
                posL = _margin_left + this.m_dIndentLeft;
            if ((_margin_left + this.m_dIndentLeftFirst) > posL)
                posL = _margin_left + this.m_dIndentLeftFirst;

            var posR = _margin_right;
            if (this.m_dIndentRight > 0)
                posR = _margin_right - this.m_dIndentRight;

            if (posL < posR)
			{
				context.strokeStyle = GlobalSkin.RulerMarkersOutlineColor;
				context.fillStyle   = GlobalSkin.RulerMarkersFillColor;
				
				blitLeftInd(this.m_dIndentLeft, this.m_bRtl);
				blitRightInd(this.m_dIndentRight, this.m_bRtl)
				blitFirstInd(this.m_dIndentLeftFirst, this.m_bRtl);
			}

            // теперь рисуем табы ----------------------------------------
            // default
            var position_default_tab = this.m_dDefaultTab;
            let _positon_y = this.m_nBottom + Math.round(1.5 * dPR);

            var _min_default_value = Math.max(0, this.m_dMaxTab);
			if (this.m_dDefaultTab > 0.01)
			{
				if (this.m_bRtl)
				{
					while (_margin_right - position_default_tab > this.m_dMarginLeft)
					{
						if (position_default_tab < _min_default_value)
						{
							position_default_tab += this.m_dDefaultTab;
							continue;
						}
						
						let _x = parseInt((_margin_right - position_default_tab) * dKoef_mm_to_pix) + left + indent;
						context.beginPath();
						context.moveTo(_x, _positon_y);
						context.lineTo(_x, _positon_y + Math.round(3 * dPR));
						context.stroke();
						
						position_default_tab += this.m_dDefaultTab;
					}
				}
				else
				{
					while (_margin_left + position_default_tab < this.m_dMarginRight)
					{
						if (position_default_tab < _min_default_value)
						{
							position_default_tab += this.m_dDefaultTab;
							continue;
						}
						
						let _x = parseInt((_margin_left + position_default_tab) * dKoef_mm_to_pix) + left + indent;
						context.beginPath();
						context.moveTo(_x, _positon_y);
						context.lineTo(_x, _positon_y + Math.round(3 * dPR));
						context.stroke();
						
						position_default_tab += this.m_dDefaultTab;
					}
				}
			}
			
			// custom tabs
			context.strokeStyle = GlobalSkin.RulerTabsColor;
			for (let i = 0, tabCount = this.m_arrTabs.length; i < tabCount; ++i)
			{
				let tab    = this.m_arrTabs[i];
				let tabPos = tab.pos;
				
				if (i === this.m_lCurrentTab)
				{
					if (!this.IsDrawingCurTab)
						continue;
					// рисуем вместо него - позицию нового
					tabPos = this.m_dCurrentTabNewPosition;
				}
				else if (tab.pos < this.m_dIndentLeft)
				{
					continue;
				}
				blitTab(tabPos, tab.type);
			}
			// -----------------------------------------------------------
		}
	}
	
	this.UpdateParaInd = function(paraInd, isRtl)
	{
		if (!paraInd)
			return 0;
		
		let left      = undefined !== paraInd.Left ? paraInd.Left : 0;
		let right     = undefined !== paraInd.Right ? paraInd.Right : 0;
		let firstLine = undefined !== paraInd.FirstLine ? left + paraInd.FirstLine : 0;
		
		let update = 0;
		
		if (Math.abs(this.m_dIndentLeft - left) > AscWord.EPSILON)
		{
			this.m_dIndentLeft = left;
			update |= 1;
		}
		
		if (Math.abs(this.m_dIndentLeftFirst - firstLine) > AscWord.EPSILON)
		{
			this.m_dIndentLeftFirst = firstLine;
			update |= 1;
		}
		
		if (Math.abs(this.m_dIndentRight - right) > AscWord.EPSILON)
		{
			this.m_dIndentRight = right;
			update |= 1;
		}
		
		if (this.m_bRtl !== isRtl)
		{
			this.m_bRtl = isRtl;
			update |= 1;
			update |= 2;
		}
		
		return update;
	};
}

function CVerRuler()
{
    this.m_oPage        = null;

    this.m_nLeft         = 0;        // значения в пикселах - смещение до самой линейки
    this.m_nRight      = 0;          // значения в пикселах - смещение до самой линейки
                                     // (т.е. ширина линейки в пикселах = (this.m_nRight - this.m_nLeft))

    this.m_dMarginTop           = 20;
    this.m_dMarginBottom        = 250;

    this.m_oCanvas              = null;

    this.m_dZoom                = 1;

    this.DragType = 0;          // 0 - none
                                // 1 - top margin, 2 - bottom margin
                                // 3 - header_top, 4 - header_bottom
                                // 5 - table rows

    // отдельные настройки для текущего объекта линейки
    this.CurrentObjectType  = RULER_OBJECT_TYPE_PARAGRAPH;
    this.m_oTableMarkup     = null;
    this.header_top         = 0;
    this.header_bottom      = 0;

    this.DragTablePos       = -1;

    this.RepaintChecker     = new CVerRulerRepaintChecker();

    // presentations addons
    this.IsCanMoveMargins = true;

    this.m_oWordControl = null;

    this.SimpleChanges = new RulerCheckSimpleChanges();

    this.Units = c_oAscDocumentUnits.Millimeter;

    this.CheckCanvas = function()
    {
        this.m_dZoom = this.m_oWordControl.m_nZoomValue / 100;
        var dPR = AscCommon.AscBrowser.retinaPixelRatio;
        var dKoef_mm_to_pix = g_dKoef_mm_to_pix * this.m_dZoom * dPR;

        var heightNew    = dKoef_mm_to_pix * this.m_oPage.height_mm;

        var _height      = Math.round(10 * dPR) + heightNew;
        if (dPR > 1)
            _height += Math.round(dPR);

        var _width       = 5 * g_dKoef_mm_to_pix * dPR;

        var intW = _width >> 0;
        var intH = _height >> 0;
        if (null == this.m_oCanvas)
        {
            this.m_oCanvas = document.createElement('canvas');
            this.m_oCanvas.width    = intW;
            this.m_oCanvas.height   = intH;
        }
        else
        {
            var oldW = this.m_oCanvas.width;
            var oldH = this.m_oCanvas.height;

            if ((oldW != intW) || (oldH != intH))
            {
                delete this.m_oCanvas;
                this.m_oCanvas = document.createElement('canvas');
                this.m_oCanvas.width    = intW;
                this.m_oCanvas.height   = intH;
            }
        }
        return heightNew;
    }

    this.CreateBackground = function(cachedPage, isattack)
    {
        if (window["NATIVE_EDITOR_ENJINE"])
            return;
            
        if (null == cachedPage || undefined == cachedPage)
            return;
        
        this.m_oPage = cachedPage;
        var height = this.CheckCanvas();

        if (0 == this.DragType)
        {
            this.m_dMarginTop       = cachedPage.margin_top;
            this.m_dMarginBottom    = cachedPage.margin_bottom;
        }

        // check old state
        var checker = this.RepaintChecker;
        var markup = this.m_oTableMarkup;

        if (isattack !== true && this.CurrentObjectType == checker.Type && height == checker.Height)
        {
            if (this.CurrentObjectType == RULER_OBJECT_TYPE_PARAGRAPH)
            {
                if (this.m_dMarginTop == checker.MarginTop && this.m_dMarginBottom == checker.MarginBottom)
                    return;
            }
            else if (this.CurrentObjectType == RULER_OBJECT_TYPE_HEADER || this.CurrentObjectType == RULER_OBJECT_TYPE_FOOTER)
            {
                if (this.header_top == checker.HeaderTop && this.header_bottom == checker.HeaderBottom)
                    return;
            }
            else if (this.CurrentObjectType == RULER_OBJECT_TYPE_TABLE)
            {
                var oldcount = checker.rowsY.length;
                var newcount = markup.Rows.length;

                if (oldcount == newcount)
                {
                    var arr1 = checker.rowsY;
                    var arr2 = checker.rowsH;

                    var rows = markup.Rows;

                    var _break = false;
                    for (var i = 0; i < oldcount; i++)
                    {
                        if ((arr1[i] != rows[i].Y) || (arr2[i] != rows[i].H))
                        {
                            _break = true;
                            break;
                        }
                    }

                    if (!_break)
                        return;
                }
            }
        }

        //console.log("vertical");

        checker.Height = height;
        checker.Type = this.CurrentObjectType;
        checker.BlitAttack = true;

        var dPR = AscCommon.AscBrowser.retinaPixelRatio;
        var dKoef_mm_to_pix = g_dKoef_mm_to_pix * this.m_dZoom * dPR;

        // не править !!!
        this.m_nLeft   = Math.round(3 * dPR);//(0.8 * g_dKoef_mm_to_pix) >> 0;
        this.m_nRight  = Math.round(15 * dPR);//(4.2 * g_dKoef_mm_to_pix) >> 0;

        var context = this.m_oCanvas.getContext('2d');
        context.setTransform(1, 0, 0, 1, 0, Math.round(5 * dPR));

        context.fillStyle = GlobalSkin.BackgroundColor;
        context.fillRect(0, 0, this.m_oCanvas.width, this.m_oCanvas.height);

        var top_margin = 0;
        var bottom_margin = 0;

        if (RULER_OBJECT_TYPE_PARAGRAPH == this.CurrentObjectType)
        {
            top_margin = (this.m_dMarginTop * dKoef_mm_to_pix) >> 0;
            bottom_margin = (this.m_dMarginBottom * dKoef_mm_to_pix) >> 0;

            checker.MarginTop = this.m_dMarginTop;
            checker.MarginBottom = this.m_dMarginBottom;
        }
        else if (RULER_OBJECT_TYPE_HEADER == this.CurrentObjectType || RULER_OBJECT_TYPE_FOOTER == this.CurrentObjectType)
        {
            top_margin = (this.header_top * dKoef_mm_to_pix) >> 0;
            bottom_margin = (this.header_bottom * dKoef_mm_to_pix) >> 0;

            checker.HeaderTop = this.header_top;
            checker.HeaderBottom = this.header_bottom;
        }
        else if (RULER_OBJECT_TYPE_TABLE == this.CurrentObjectType)
        {
            var _arr1 = checker.rowsY;
            var _arr2 = checker.rowsH;

            if (0 != _arr1.length)
                _arr1.splice(0, _arr1.length);

            if (0 != _arr2.length)
                _arr2.splice(0, _arr2.length);

            var _count = this.m_oTableMarkup.Rows.length;

            for (var i = 0; i < _count; i++)
            {
                _arr1[i] = markup.Rows[i].Y;
                _arr2[i] = markup.Rows[i].H;
            }

            if (_count != 0)
            {
                top_margin = (markup.Rows[0].Y * dKoef_mm_to_pix) >> 0;
                bottom_margin = ((markup.Rows[_count - 1].Y + markup.Rows[_count - 1].H) * dKoef_mm_to_pix) >> 0;
            }
        }

        var indent = 0.5 * Math.round(dPR);

        if (bottom_margin > top_margin)
        {
            context.fillStyle = GlobalSkin.RulerLight;
            context.fillRect(this.m_nLeft + indent, top_margin + indent, this.m_nRight - this.m_nLeft, bottom_margin - top_margin);
        }

        var intH = height >> 0;

        if (true)
        {
            context.beginPath();
            context.fillStyle = GlobalSkin.RulerDark;

            context.fillRect(this.m_nLeft + indent, indent, this.m_nRight - this.m_nLeft, top_margin);
            context.fillRect(this.m_nLeft + indent, bottom_margin + indent, this.m_nRight - this.m_nLeft, Math.max(intH - bottom_margin, 1));
            context.beginPath();
        }

        context.beginPath();
        context.lineWidth = Math.round(dPR);
        context.strokeStyle = GlobalSkin.RulerTextColor;
        context.fillStyle = GlobalSkin.RulerTextColor;

        var mm_1_4 = 10 * dKoef_mm_to_pix / 4;
		var isDraw1_4 = (mm_1_4 > 7) ? true : false;
        var inch_1_8 = 25.4 * dKoef_mm_to_pix / 8;

        var middleHor = (this.m_nLeft + this.m_nRight) / 2;
        var part1 = dPR;
        var part2 = 2.5 * dPR;

		var l_part1 = Math.floor(middleHor - part1);
		var r_part1 = Math.ceil(middleHor + part1);
		var l_part2 = Math.floor(middleHor - part2);
		var r_part2 = Math.ceil(middleHor + part2);

        context.font = Math.round(7 * dPR) + "pt Arial";

        if (this.Units == c_oAscDocumentUnits.Millimeter)
        {
            var lCount1 = ((height - top_margin) / mm_1_4) >> 0;
            var lCount2 = (top_margin / mm_1_4) >> 0;

        var index = 0;
        var num = 0;
        for (var i = 1; i < lCount1; i++)
        {
            var lYPos = ((top_margin + i * mm_1_4) >> 0) + indent;
            index++;

            if (index == 4)
                index = 0;

            if (0 == index)
            {
                num++;
                // number

                var strNum = "" + num;
                var lWidthText = context.measureText(strNum).width;

                context.translate(middleHor, lYPos);
                context.rotate(-Math.PI / 2);
                context.fillText(strNum, -lWidthText / 2.0, Math.round(4 * dPR));
                context.setTransform(1, 0, 0, 1, 0, Math.round(5 * dPR));
            }
            else if (1 == index && isDraw1_4)
            {
                // 1/4
                context.beginPath();
                context.moveTo(l_part1, lYPos);
                context.lineTo(r_part1, lYPos);
                context.stroke();
            }
            else if (2 == index)
            {
                // 1/2
                context.beginPath();
                context.moveTo(l_part2, lYPos);
                context.lineTo(r_part2, lYPos);
                context.stroke();
            }
            else if (isDraw1_4)
            {
                // 1/4
                context.beginPath();
                context.moveTo(l_part1, lYPos);
                context.lineTo(r_part1, lYPos);
                context.stroke();
            }
        }

        index = 0;
        num = 0;
        for (var i = 1; i <= lCount2; i++)
        {
            var lYPos = ((top_margin - i * mm_1_4) >> 0) + indent;
            index++;

            if (index == 4)
                index = 0;

            if (0 == index)
            {
                num++;
                // number
                var strNum = "" + num;
                var lWidthText = context.measureText(strNum).width;

                context.translate(middleHor, lYPos);
                context.rotate(-Math.PI / 2);
                context.fillText(strNum, -lWidthText / 2.0, Math.round(4 * dPR));
                context.setTransform(1, 0, 0, 1, 0, Math.round(5 * dPR));
            }
            else if (1 == index && isDraw1_4)
            {
                // 1/4
                context.beginPath();
                context.moveTo(l_part1, lYPos);
                context.lineTo(r_part1, lYPos);
                context.stroke();
            }
            else if (2 == index)
            {
                // 1/2
                context.beginPath();
                context.moveTo(l_part2, lYPos);
                context.lineTo(r_part2, lYPos);
                context.stroke();
            }
            else if (isDraw1_4)
            {
                // 1/4
                context.beginPath();
                context.moveTo(l_part1, lYPos);
                context.lineTo(r_part1, lYPos);
                context.stroke();
            }
        }
        }
        else if (this.Units == c_oAscDocumentUnits.Inch)
        {
            var lCount1 = ((height - top_margin) / inch_1_8) >> 0;
            var lCount2 = (top_margin / inch_1_8) >> 0;

            var index = 0;
            var num = 0;
            for (var i = 1; i < lCount1; i++)
            {
                var lYPos = ((top_margin + i * inch_1_8) >> 0) + indent;
                index++;

                if (index == 8)
                    index = 0;

                if (0 == index)
                {
                    num++;
                    // number

                    var strNum = "" + num;
                    var lWidthText = context.measureText(strNum).width;

                    context.translate(middleHor, lYPos);
                    context.rotate(-Math.PI / 2);
                    context.fillText(strNum, -lWidthText / 2.0, Math.round(4 * dPR));
                    context.setTransform(1, 0, 0, 1, 0, Math.round(5 * dPR));
                }
                else if (4 == index)
                {
                    // 1/2
                    context.beginPath();
                    context.moveTo(l_part2, lYPos);
                    context.lineTo(r_part2, lYPos);
                    context.stroke();
                }
                else if (inch_1_8 > 8)
                {
                    // 1/8
                    context.beginPath();
                    context.moveTo(l_part1, lYPos);
                    context.lineTo(r_part1, lYPos);
                    context.stroke();
                }
            }

            index = 0;
            num = 0;
            for (var i = 1; i <= lCount2; i++)
            {
                var lYPos = ((top_margin - i * inch_1_8) >> 0) + indent;
                index++;

                if (index == 8)
                    index = 0;

                if (0 == index)
                {
                    num++;
                    // number
                    var strNum = "" + num;
                    var lWidthText = context.measureText(strNum).width;

                    context.translate(middleHor, lYPos);
                    context.rotate(-Math.PI / 2);
                    context.fillText(strNum, -lWidthText / 2.0, Math.round(4 * dPR));
                    context.setTransform(1, 0, 0, 1, 0, Math.round(5 * dPR));

                }
                else if (4 == index)
                {
                    // 1/2
                    context.beginPath();
                    context.moveTo(middleHor - part2, lYPos);
                    context.lineTo(middleHor + part2, lYPos);
                    context.stroke();
                }
                else if (inch_1_8 > 8)
                {
                    // 1/8
                    context.beginPath();
                    context.moveTo(middleHor - part1, lYPos);
                    context.lineTo(middleHor + part1, lYPos);
                    context.stroke();
                }
            }
        }
        else if (this.Units == c_oAscDocumentUnits.Point)
        {
            var point_1_12 = 25.4 * dKoef_mm_to_pix / 12;

            var lCount1 = ((height - top_margin) / point_1_12) >> 0;
            var lCount2 = (top_margin / point_1_12) >> 0;

            var index = 0;
            var num = 0;
            for (var i = 1; i < lCount1; i++)
            {
                var lYPos = ((top_margin + i * point_1_12) >> 0) + indent;
                index++;

                if (index == 12)
                    index = 0;

                if (0 == index || 6 == index)
                {
                    num++;
                    // number

                    var strNum = "" + (num * 36);
                    var lWidthText = context.measureText(strNum).width;

                    context.translate(middleHor, lYPos);
                    context.rotate(-Math.PI / 2);
                    context.fillText(strNum, -lWidthText / 2.0, Math.round(4 * dPR));
                    context.setTransform(1, 0, 0, 1, 0, Math.round(5 * dPR));
                }
                else if (point_1_12 > 5)
                {
                    // 1/8
                    context.beginPath();
                    context.moveTo(l_part1, lYPos);
                    context.lineTo(r_part1, lYPos);
                    context.stroke();
                }
            }

            index = 0;
            num = 0;
            for (var i = 1; i <= lCount2; i++)
            {
                var lYPos = ((top_margin - i * point_1_12) >> 0) + indent;
                index++;

                if (index == 12)
                    index = 0;

                if (0 == index || 6 == index)
                {
                    num++;
                    // number
                    var strNum = "" + (num * 36);
                    var lWidthText = context.measureText(strNum).width;

                    context.translate(middleHor, lYPos);
                    context.rotate(-Math.PI / 2);
                    context.fillText(strNum, -lWidthText / 2.0, Math.round(4 * dPR));
                    context.setTransform(1, 0, 0, 1, 0, Math.round(5 * dPR));
                }
                else if (point_1_12 > 5)
                {
                    // 1/8
                    context.beginPath();
                    context.moveTo(l_part1, lYPos);
                    context.lineTo(r_part1, lYPos);
                    context.stroke();
                }
            }
        }

        if ((this.CurrentObjectType == RULER_OBJECT_TYPE_TABLE) && (null != markup))
        {
            var dPR = AscCommon.AscBrowser.retinaPixelRatio;
            var dKoef_mm_to_pix = g_dKoef_mm_to_pix * this.m_dZoom * dPR;

            // не будет нулевых таблиц.
            var _count = markup.Rows.length;

            if (0 == _count)
                return;

            var start_dark = (((markup.Rows[0].Y + markup.Rows[0].H) * dKoef_mm_to_pix) >> 0) + indent;
            var end_dark = 0;

            context.fillStyle = GlobalSkin.RulerDark;
            context.strokeStyle = GlobalSkin.RulerOutline;

            var _x = this.m_nLeft + indent;
            var _w = this.m_nRight - this.m_nLeft;
            for (var i = 1; i < _count; i++)
            {
                end_dark = ((markup.Rows[i].Y * dKoef_mm_to_pix) >> 0) + indent;
                context.fillRect(_x, start_dark, _w, Math.max(end_dark - start_dark, Math.round(7 * dPR)));
                context.strokeRect(_x, start_dark, _w, Math.max(end_dark - start_dark, Math.round(7 * dPR)));

                start_dark = (((markup.Rows[i].Y + markup.Rows[i].H) * dKoef_mm_to_pix) >> 0) + indent;
            }
        }

        // рамка
        context.beginPath();
        context.strokeStyle = GlobalSkin.RulerOutline;
        context.strokeRect(this.m_nLeft + indent, indent, this.m_nRight - this.m_nLeft, Math.max(intH - 1, 1));
        context.beginPath();
        context.moveTo(this.m_nLeft + indent, top_margin + indent);
        context.lineTo(this.m_nRight - indent, top_margin + indent);

        context.moveTo(this.m_nLeft + indent, bottom_margin + indent);
        context.lineTo(this.m_nRight - indent, bottom_margin + indent);
        context.stroke();
    }

    this.OnMouseMove = function(left, top, e)
    {
        var word_control = this.m_oWordControl;
        AscCommon.check_MouseMoveEvent(e);

        this.SimpleChanges.CheckMove();
        var ver_ruler = word_control.m_oLeftRuler_vertRuler;
        var dKoefPxToMM = 100 * g_dKoef_pix_to_mm / word_control.m_nZoomValue;

		// теперь определяем позицию относительно самой линейки. Все в миллиметрах
        var _y = global_mouseEvent.Y - 7 * g_dKoef_mm_to_pix - top - word_control.Y;
        _y *= dKoefPxToMM;
        var _x = left * g_dKoef_pix_to_mm;

        var dKoef_mm_to_pix = g_dKoef_mm_to_pix * this.m_dZoom;
        var mm_1_4 = 10 / 4;
        var mm_1_8 = mm_1_4 / 2;

        switch (this.DragType)
        {
            case 0:
            {
                if (this.CurrentObjectType == RULER_OBJECT_TYPE_PARAGRAPH)
                {
                    if (this.IsCanMoveMargins && ((Math.abs(_y - this.m_dMarginTop) < 1) || (Math.abs(_y - this.m_dMarginBottom) < 1)))
                        word_control.m_oDrawingDocument.SetCursorType("s-resize");
                    else
                        word_control.m_oDrawingDocument.SetCursorType("default");
                }
                else if (this.CurrentObjectType == RULER_OBJECT_TYPE_HEADER)
                {
                    if ((Math.abs(_y - this.header_top) < 1) || (Math.abs(_y - this.header_bottom) < 1))
                        word_control.m_oDrawingDocument.SetCursorType("s-resize");
                    else
                        word_control.m_oDrawingDocument.SetCursorType("default");
                }
                else if (this.CurrentObjectType == RULER_OBJECT_TYPE_FOOTER)
                {
                    if (Math.abs(_y - this.header_top) < 1)
                        word_control.m_oDrawingDocument.SetCursorType("s-resize");
                    else
                        word_control.m_oDrawingDocument.SetCursorType("default");
                }
                else if (this.CurrentObjectType == RULER_OBJECT_TYPE_TABLE)
                {
                    var type = this.CheckMouseType(2, _y);
                    if (type == 5)
                        word_control.m_oDrawingDocument.SetCursorType("s-resize");
                    else
                        word_control.m_oDrawingDocument.SetCursorType("default");
                }

                break;
            }
            case 1:
            {
                var newVal = RulerCorrectPosition(this, _y, this.m_dMarginTop);

                if (newVal > (this.m_dMarginBottom - 30))
                    newVal = this.m_dMarginBottom - 30;
                if (newVal < 0)
                    newVal = 0;

                this.m_dMarginTop = newVal;
                word_control.UpdateVerRulerBack();

                var pos = top + this.m_dMarginTop * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.HorLine(pos);

                word_control.m_oDrawingDocument.SetCursorType("s-resize");

                break;
            }
            case 2:
            {
                var newVal = RulerCorrectPosition(this, _y, this.m_dMarginTop);

                if (newVal < (this.m_dMarginTop + 30))
                    newVal = this.m_dMarginTop + 30;
                if (newVal > this.m_oPage.height_mm)
                    newVal = this.m_oPage.height_mm;

                this.m_dMarginBottom = newVal;
                word_control.UpdateVerRulerBack();

                var pos = top + this.m_dMarginBottom * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.HorLine(pos);

                word_control.m_oDrawingDocument.SetCursorType("s-resize");

                break;
            }
            case 3:
            {
                var newVal = RulerCorrectPosition(this, _y, this.m_dMarginTop);

                if (newVal > this.header_bottom)
                    newVal = this.header_bottom;
                if (newVal < 0)
                    newVal = 0;

                this.header_top = newVal;
                word_control.UpdateVerRulerBack();

                var pos = top + this.header_top * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.HorLine(pos);

                word_control.m_oDrawingDocument.SetCursorType("s-resize");

                break;
            }
            case 4:
            {
                var newVal = RulerCorrectPosition(this, _y, this.m_dMarginTop);

                if (newVal < 0)
                    newVal = 0;
                if (newVal > this.m_oPage.height_mm)
                    newVal = this.m_oPage.height_mm;

                this.header_bottom = newVal;
                word_control.UpdateVerRulerBack();

                var pos = top + this.header_bottom * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.HorLine(pos);

                word_control.m_oDrawingDocument.SetCursorType("s-resize");

                break;
            }
            case 5:
            {
                // сначала нужно определить минимум и максимум сдвига
                var _min = 0;
                var _max = this.m_oPage.height_mm;

                if (0 < this.DragTablePos)
                {
                    _min = this.m_oTableMarkup.Rows[this.DragTablePos - 1].Y;
                }
                if (this.DragTablePos < this.m_oTableMarkup.Rows.length)
                {
                    _max = this.m_oTableMarkup.Rows[this.DragTablePos].Y + this.m_oTableMarkup.Rows[this.DragTablePos].H;
                }

                var newVal = RulerCorrectPosition(this, _y, this.m_dMarginTop);

                if (newVal < _min)
                    newVal = _min;
                if (newVal > _max)
                    newVal = _max;

                if (0 == this.DragTablePos)
                {
                    var _bottom = this.m_oTableMarkup.Rows[0].Y + this.m_oTableMarkup.Rows[0].H;
                    this.m_oTableMarkup.Rows[0].Y = newVal;
                    this.m_oTableMarkup.Rows[0].H = _bottom - newVal;
                }
                else
                {
                    var oldH = this.m_oTableMarkup.Rows[this.DragTablePos - 1].H;
                    this.m_oTableMarkup.Rows[this.DragTablePos - 1].H = newVal - this.m_oTableMarkup.Rows[this.DragTablePos - 1].Y;

                    var delta = this.m_oTableMarkup.Rows[this.DragTablePos - 1].H - oldH;
                    for (var i = this.DragTablePos; i < this.m_oTableMarkup.Rows.length; i++)
                    {
                        this.m_oTableMarkup.Rows[i].Y += delta;
                    }
                }

                word_control.UpdateVerRulerBack();

                var pos = top + newVal * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.HorLine(pos);

                word_control.m_oDrawingDocument.SetCursorType("s-resize");
            }
        }
    }

    this.CheckMouseType = function(x, y)
    {
		// проверяем где находимся

        if (this.IsCanMoveMargins === false)
            return 0;

        if (x >= 0.8 && x <= 4.2)
        {
			// внутри линейки
            if (this.CurrentObjectType == RULER_OBJECT_TYPE_PARAGRAPH)
            {
                if (Math.abs(y - this.m_dMarginTop) < 1)
                {
                    return 1;
                }
                else if (Math.abs(y - this.m_dMarginBottom) < 1)
                {
                    return 2;
                }
            }
            else if (this.CurrentObjectType == RULER_OBJECT_TYPE_HEADER)
            {
                if (Math.abs(y - this.header_top) < 1)
                {
                    return 3;
                }
                else if (Math.abs(y - this.header_bottom) < 1)
                {
                    return 4;
                }
            }
            else if (this.CurrentObjectType == RULER_OBJECT_TYPE_FOOTER)
            {
                if (Math.abs(y - this.header_top) < 1)
                {
                    return 3;
                }
            }
            else if (this.CurrentObjectType == RULER_OBJECT_TYPE_TABLE && null != this.m_oTableMarkup)
            {
                // не будет нулевых таблиц.
                var markup = this.m_oTableMarkup;
                var _count = markup.Rows.length;

                if (0 == _count)
                    return 0;

                var _start = markup.Rows[0].Y;
                var _end = _start - 2;

                for (var i = 0; i <= _count; i++)
                {
                    if (i == _count)
                    {
                        _end = markup.Rows[i - 1].Y + markup.Rows[i - 1].H;
                        _start = _end + 2;
                    }
                    else if (i != 0)
                    {
                        _end = markup.Rows[i - 1].Y + markup.Rows[i - 1].H;
                        _start = markup.Rows[i].Y;
                    }

                    if ((_end - 1) < y && y < (_start + 1))
                    {
                        this.DragTablePos = i;
                        return 5;
                    }
                }
            }
        }
        return 0;
    }

    this.OnMouseDown = function(left, top, e)
    {
        var word_control = this.m_oWordControl;

		if (true === word_control.m_oApi.isStartAddShape)
		{
			word_control.m_oApi.sync_EndAddShape();
			word_control.m_oApi.sync_StartAddShapeCallback(false);
		}
		if (true === word_control.m_oApi.isInkDrawerOn())
		{
			word_control.m_oApi.stopInkDrawer();
		}

        AscCommon.check_MouseDownEvent(e, true);

        this.SimpleChanges.Reinit();
        global_mouseEvent.LockMouse();

        var dKoefPxToMM = 100 * g_dKoef_pix_to_mm / word_control.m_nZoomValue;
        var dKoef_mm_to_pix = g_dKoef_mm_to_pix * this.m_dZoom;

        var _y = global_mouseEvent.Y - 7 * g_dKoef_mm_to_pix - top - word_control.Y;
        _y *= dKoefPxToMM;
        var _x = (global_mouseEvent.X - word_control.X) * g_dKoef_pix_to_mm - word_control.GetMainContentBounds().L - word_control.GetVertRulerLeft();

        this.DragType = this.CheckMouseType(_x, _y);
		this.DragTypeMouseDown = this.DragType;

		if (global_mouseEvent.ClickCount > 1)
		{
			var eventType = "";
			switch (this.DragTypeMouseDown)
			{
				case 5:
				    eventType = "tables";
				    break;
				default:
					eventType = "margins";
					break;
			}

			if (eventType != "")
			{
				word_control.m_oApi.sendEvent("asc_onRulerDblClick", eventType);
				this.DragType = 0;
				this.OnMouseUp(left, top, e);
				return;
			}
		}

        switch (this.DragType)
        {
            case 1:
            {
                var pos = top + this.m_dMarginTop * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.HorLine(pos);
                break;
            }
            case 2:
            {
                var pos = top + this.m_dMarginBottom * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.HorLine(pos);
                break;
            }
            case 3:
            {
                var pos = top + this.header_top * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.HorLine(pos);
                break;
            }
            case 4:
            {
                var pos = top + this.header_bottom * dKoef_mm_to_pix;
                word_control.m_oOverlayApi.HorLine(pos);
                break;
            }
            case 5:
            {
                var pos = 0;
                if (0 == this.DragTablePos)
                {
                    pos = top + this.m_oTableMarkup.Rows[0].Y * dKoef_mm_to_pix;
                    word_control.m_oOverlayApi.HorLine(pos);
                }
                else
                {
                    pos = top + (this.m_oTableMarkup.Rows[this.DragTablePos - 1].Y + this.m_oTableMarkup.Rows[this.DragTablePos - 1].H) * dKoef_mm_to_pix;
                    word_control.m_oOverlayApi.HorLine(pos);
                }
            }
        }

        word_control.m_oDrawingDocument.LockCursorTypeCur();
    }

    this.OnMouseUp = function(left, top, e)
    {
        var lockedElement = AscCommon.check_MouseUpEvent(e);

        //this.m_oWordControl.m_oOverlayApi.UnShow();
        this.m_oWordControl.OnUpdateOverlay();

        switch (this.DragType)
        {
            case 1:
            case 2:
            {
                if (!this.SimpleChanges.IsSimple)
                    this.SetMarginProperties();
                break;
            }
            case 3:
            case 4:
            {
                if (!this.SimpleChanges.IsSimple)
                    this.SetHeaderProperties();
                break;
            }
            case 5:
            {
                if (!this.SimpleChanges.IsSimple)
                    this.SetTableProperties();
                this.DragTablePos = -1;
                break;
            }
        }

        this.DragType = 0;
        this.m_oWordControl.m_oDrawingDocument.UnlockCursorType();

        this.SimpleChanges.Clear();
    }

    this.OnMouseUpExternal = function()
    {
        //this.m_oWordControl.m_oOverlayApi.UnShow();
        this.m_oWordControl.OnUpdateOverlay();

        switch (this.DragType)
        {
            case 1:
            case 2:
            {
                if (!this.SimpleChanges.IsSimple)
                    this.SetMarginProperties();
                break;
            }
            case 3:
            case 4:
            {
                if (!this.SimpleChanges.IsSimple)
                    this.SetHeaderProperties();
                break;
            }
            case 5:
            {
                if (!this.SimpleChanges.IsSimple)
                    this.SetTableProperties();
                this.DragTablePos = -1;
                break;
            }
        }

        this.DragType = 0;
        this.m_oWordControl.m_oDrawingDocument.UnlockCursorType();

        this.SimpleChanges.Clear();
    }

    this.BlitToMain = function(left, top, htmlElement)
    {
        if (!this.RepaintChecker.BlitAttack && top == this.RepaintChecker.BlitTop)
            return;
        var dPR = AscCommon.AscBrowser.retinaPixelRatio;
        top = Math.round(top * dPR);
        this.RepaintChecker.BlitTop = top;
        this.RepaintChecker.BlitAttack = false;

        htmlElement.width = htmlElement.width;
        var context = htmlElement.getContext('2d');

        if (null != this.m_oCanvas)
        {
                context.drawImage(this.m_oCanvas, 0, Math.round(5 * dPR), this.m_oCanvas.width, this.m_oCanvas.height - Math.round(10 * dPR),
                    0, top, this.m_oCanvas.width, this.m_oCanvas.height - Math.round(10 * dPR));
        }
    }

    // выставление параметров логическому документу
    this.SetMarginProperties = function()
    {
        if ( false === this.m_oWordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Document_SectPr) )
        {
            this.m_oWordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetDocumentMargin_Ver);
            this.m_oWordControl.m_oLogicDocument.SetDocumentMargin( { Top : this.m_dMarginTop, Bottom : this.m_dMarginBottom }, true);
			this.m_oWordControl.m_oLogicDocument.FinalizeAction();
        }
    }
    this.SetHeaderProperties = function()
    {
        if ( false === this.m_oWordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_HdrFtr) )
        {
            // TODO: в данной функции при определенных параметрах может меняться верхнее поле. Поэтому, надо
            //       вставить проверку на залоченность с типом changestype_Document_SectPr

            this.m_oWordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetHdrFtrBounds);
            this.m_oWordControl.m_oLogicDocument.Document_SetHdrFtrBounds(this.header_top, this.header_bottom);
			this.m_oWordControl.m_oLogicDocument.FinalizeAction();
        }
    }
    this.SetTableProperties = function()
    {
        if ( false === this.m_oWordControl.m_oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Table_Properties) )
        {
            this.m_oWordControl.m_oLogicDocument.StartAction(AscDFH.historydescription_Document_SetTableMarkup_Ver);

            this.m_oTableMarkup.CorrectTo();
            this.m_oTableMarkup.Table.Update_TableMarkupFromRuler(this.m_oTableMarkup, false, this.DragTablePos);
            if (this.m_oTableMarkup)
                this.m_oTableMarkup.CorrectFrom();

			this.m_oWordControl.m_oLogicDocument.FinalizeAction();
        }
    }
}
