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
var FontStyle = AscFonts.FontStyle;
var g_fontApplication = AscFonts.g_fontApplication;

var CColor = AscCommon.CColor;
var CAscMathCategory = AscCommon.CAscMathCategory;
var g_oTextMeasurer = AscCommon.g_oTextMeasurer;
var global_mouseEvent = AscCommon.global_mouseEvent;
var global_MatrixTransformer = AscCommon.global_MatrixTransformer;
var g_dKoef_pix_to_mm = AscCommon.g_dKoef_pix_to_mm;
var g_dKoef_mm_to_pix = AscCommon.g_dKoef_mm_to_pix;

function CStylesPainter()
{
	this.mergedStyles = null;

	this.STYLE_THUMBNAIL_WIDTH = GlobalSkin.STYLE_THUMBNAIL_WIDTH;
	this.STYLE_THUMBNAIL_HEIGHT = GlobalSkin.STYLE_THUMBNAIL_HEIGHT;

	this.previewGenerator = new AscCommon.StylePreviewGenerator(this);
}
CStylesPainter.prototype.CheckStylesNames = function(_api, ds)
{
	var DocumentStyles = _api.WordControl.m_oLogicDocument.Get_Styles();
	for (var i in ds)
	{
		var style = ds[i];
		if (style.IsExpressStyle(DocumentStyles) && null === DocumentStyles.GetStyleIdByName(style.Name))
		{
			AscFonts.FontPickerByCharacter.getFontsByString(style.Name);
		}
	}

	AscFonts.FontPickerByCharacter.getFontsByString("Aa");

	var styles = DocumentStyles.Style;

	if (!styles)
		return;

	for (var i in styles)
	{
		var style = styles[i];
		if (style.IsExpressStyle(DocumentStyles))
		{
			AscFonts.FontPickerByCharacter.getFontsByString(style.Name);
			AscFonts.FontPickerByCharacter.getFontsByString(AscCommon.translateManager.getValue(style.Name));
		}
	}
};
CStylesPainter.prototype.GenerateStyles = function(_api)
{
	if (this.previewGenerator.IsInProgress())
		this.previewGenerator.Stop();
	
	this.previewGenerator.Begin(_api);
};
CStylesPainter.prototype.OnEndGenerate = function(styles, _api)
{
	this.mergedStyles = styles;
	
	// export
	this["STYLE_THUMBNAIL_WIDTH"] = this.STYLE_THUMBNAIL_WIDTH;
	this["STYLE_THUMBNAIL_HEIGHT"] = this.STYLE_THUMBNAIL_HEIGHT;

	// теперь просто отдаем евент наверх
	_api.sync_InitEditorStyles(this);
};
CStylesPainter.prototype.get_MergedStyles = function ()
{
	return this.mergedStyles;
};

(function(window)
{
	/**
	 * @param stylePainter {AscCommonWord.CStylesPainter}
	 * @constructor
	 * @extends AscCommon.CActionOnTimerBase
	 */
	function StylePreviewGenerator(stylePainter)
	{
		AscCommon.CActionOnTimerBase.call(this);
		
		this.stylePainter = stylePainter;
		this.api          = null;
		this.styleManager = null;
		
		this.canvas   = null;
		this.graphics = null;
		
		this.STYLE_THUMBNAIL_WIDTH  = GlobalSkin.STYLE_THUMBNAIL_WIDTH;
		this.STYLE_THUMBNAIL_HEIGHT = GlobalSkin.STYLE_THUMBNAIL_HEIGHT;
		
		this.defaultStyles = [];
		this.docStyles     = [];
		
		this.CurrentTranslate = null;
	}
	StylePreviewGenerator.prototype = Object.create(AscCommon.CActionOnTimerBase.prototype);
	StylePreviewGenerator.prototype.OnBegin = function(_api)
	{
		this.api = _api;
		if (!_api.WordControl.m_oLogicDocument)
			return;
		
		this.styleManager = _api.WordControl.m_oLogicDocument.Get_Styles();
		if (!this.styleManager.Style)
			return;
		
		this.styles = [];
		for (let i in this.styleManager.Style)
		{
			let style = this.styleManager.Style[i];
			if (style.IsExpressStyle(this.styleManager))
				this.styles.push(this.styleManager.Style[i]);
		}
		
		this.styles.sort(function(st1, st2){
			let p1 = st1.GetUiPriority();
			let p2 = st2.GetUiPriority();
			if (null === p1 || undefined === p1)
				return -1;
			if (null === p2 || undefined === p2)
				return 1;
			
			return p1 === p2 ? 0 : p1 < p2 ? 1 : -1;
		});
		
		this.stylePainter.docStyles = [];
		this.index  = 0;
		
		this.STYLE_THUMBNAIL_WIDTH 	= AscCommon.AscBrowser.convertToRetinaValue(this.stylePainter.STYLE_THUMBNAIL_WIDTH, true);
		this.STYLE_THUMBNAIL_HEIGHT = AscCommon.AscBrowser.convertToRetinaValue(this.stylePainter.STYLE_THUMBNAIL_HEIGHT, true);
		
		this.CurrentTranslate = _api.CurrentTranslate;
		
		this.InitCanvas();
		
		this.defaultStyles = [];
		this.docStyles     = [];
		
		this.GenerateDefaultStyles();
	};
	StylePreviewGenerator.prototype.OnEnd = function()
	{
		var _count_default = this.defaultStyles.length;
		var _count_doc = 0;
		if (null != this.docStyles)
			_count_doc = this.docStyles.length;
		
		var aPriorityStyles = [];
		var fAddToPriorityStyles = function (style)
		{
			var index = style.uiPriority;
			if (null == index)
				index = 0;
			var aSubArray = aPriorityStyles[index];
			if (null == aSubArray)
			{
				aSubArray = [];
				aPriorityStyles[index] = aSubArray;
			}
			aSubArray.push(style);
		};
		var _map_document = {};
		
		for (var i = 0; i < _count_doc; i++)
		{
			var style = this.docStyles[i];
			_map_document[style.Name] = 1;
			fAddToPriorityStyles(style);
		}
		
		for (var i = 0; i < _count_default; i++)
		{
			var style = this.defaultStyles[i];
			if (null == _map_document[style.Name])
				fAddToPriorityStyles(style);
		}
		
		let mergedStyles = [];
		for (var index in aPriorityStyles)
		{
			var aSubArray = aPriorityStyles[index];
			aSubArray.sort(function (a, b)
			{
				if (a.Name < b.Name)
					return -1;
				else if (a.Name > b.Name)
					return 1;
				else
					return 0;
			});
			for (var i = 0, length = aSubArray.length; i < length; ++i)
			{
				mergedStyles.push(aSubArray[i]);
			}
		}
		
		this.stylePainter.OnEndGenerate(mergedStyles, this.api);
	};
	StylePreviewGenerator.prototype.IsContinue = function()
	{
		return (this.index < this.styles.length);
	};
	StylePreviewGenerator.prototype.DoAction = function()
	{
		let style = this.styles[this.index++];
		if (!style)
			return;
		
		this.docStyles.push(this.GeneratePreview(style));
	};
	StylePreviewGenerator.prototype.OnEndTimer = function()
	{
		// TODO: Добавить обработку, если сделаем постепенное заполнение
	};
	StylePreviewGenerator.prototype.InitCanvas = function()
	{
		var _canvas = document.createElement('canvas');
		_canvas.width  = this.STYLE_THUMBNAIL_WIDTH;
		_canvas.height = this.STYLE_THUMBNAIL_HEIGHT;
		var ctx = _canvas.getContext('2d');
		
		if (window["flat_desine"] !== true)
		{
			ctx.fillStyle = "#FFFFFF";
			ctx.fillRect(0, 0, _canvas.width, _canvas.height);
		}
		
		var graphics = new AscCommon.CGraphics();
		var koef = AscCommon.g_dKoef_pix_to_mm / AscCommon.AscBrowser.retinaPixelRatio;
		graphics.init(ctx, _canvas.width, _canvas.height, _canvas.width * koef, _canvas.height * koef);
		graphics.m_oFontManager = AscCommon.g_fontManager;
		
		this.canvas   = _canvas;
		this.graphics = graphics;
	};
	StylePreviewGenerator.prototype.GeneratePreview = function(style)
	{
		let uiPriority = style.GetUiPriority();
		let styleId    = style.GetId();
		let styleName  = style.GetName();
		
		// как только меняется сериалайзер - меняется и код здесь. Да, не очень удобно,
		// зато быстро делается
		var formalStyle = styleId.toLowerCase().replace(/\s/g, "");
		var res = formalStyle.match(/^heading([1-9][0-9]*)$/);
		var index = (res) ? res[1] - 1 : -1;
		
		if (style.Default)
		{
			switch (style.Default)
			{
				case 1:
					break;
				case 2:
					styleName = "No List";
					break;
				case 3:
					styleName = "Normal";
					break;
				case 4:
					styleName = "Normal Table";
					break;
			}
		}
		else if (-1 !== index)
		{
			styleName = "Heading ".concat(index + 1);
		}
		
		var _dr_style  = this.styleManager.Get_Pr(styleId, AscWord.styletype_Paragraph);
		_dr_style.Name = styleName;
		_dr_style.Id   = styleId;
		
		this.drawStyle(this.api, this.graphics, _dr_style, this.styleManager.IsStyleDefaultByName(styleName) ? AscCommon.translateManager.getValue(styleName) : styleName);
		return new AscCommon.CStyleImage(styleName, AscCommon.c_oAscStyleImage.Document, this.canvas.toDataURL("image/png"), uiPriority);
	};
	StylePreviewGenerator.prototype.drawStyle = function(_api, graphics, style, styleName)
	{
		let logicDocument = _api.getLogicDocument();
		if (!logicDocument)
			return;
		
		let isShowParaMarks = _api.get_ShowParaMarks();
		if (isShowParaMarks)
			_api.put_ShowParaMarks(false);
		
		let oldTabStop = AscCommonWord.Default_Tab_Stop;
		AscCommonWord.Default_Tab_Stop = 1;
		
		let _t = this;
		AscCommon.executeNoRevisions(function()
		{
			AscCommon.ExecuteNoHistory(function()
			{
				_t._drawStyle(_api, graphics, style, styleName);
			}, logicDocument);
		}, logicDocument);
		
		AscCommonWord.Default_Tab_Stop = oldTabStop;
		
		if (isShowParaMarks)
			_api.put_ShowParaMarks(isShowParaMarks);
	};
	StylePreviewGenerator.prototype._drawStyle = function(_api, graphics, style, styleName)
	{
		var ctx = graphics.m_oContext;
		ctx.fillStyle = "#FFFFFF";
		ctx.fillRect(0, 0, this.STYLE_THUMBNAIL_WIDTH, this.STYLE_THUMBNAIL_HEIGHT);
		
		var font = {
			FontFamily: {Name: "Times New Roman", Index: -1},
			Color: {r: 0, g: 0, b: 0},
			Bold: false,
			Italic: false,
			FontSize: 10
		};
		
		var textPr = style.TextPr;
		if (textPr.FontFamily !== undefined)
		{
			font.FontFamily.Name = textPr.FontFamily.Name;
			font.FontFamily.Index = textPr.FontFamily.Index;
		}
		
		if (textPr.Bold !== undefined)
			font.Bold = textPr.Bold;
		if (textPr.Italic !== undefined)
			font.Italic = textPr.Italic;
		
		if (textPr.FontSize !== undefined)
			font.FontSize = textPr.FontSize;
		
		graphics.SetFont(font);
		
		if (textPr.Color === undefined)
			graphics.b_color1(0, 0, 0, 255);
		else
			graphics.b_color1(textPr.Color.r, textPr.Color.g, textPr.Color.b, 255);
		
		var dKoefToMM = AscCommon.g_dKoef_pix_to_mm;
		dKoefToMM /= AscCommon.AscBrowser.retinaPixelRatio;
		
		if (window["flat_desine"] !== true)
		{
			var y = 0;
			var b = dKoefToMM * this.STYLE_THUMBNAIL_HEIGHT;
			var w = dKoefToMM * this.STYLE_THUMBNAIL_WIDTH;
			
			graphics.transform(1, 0, 0, 1, 0, 0);
			graphics.save();
			graphics._s();
			graphics._m(-0.5, y);
			graphics._l(w, y);
			graphics._l(w, b);
			graphics._l(0, b);
			graphics._z();
			graphics.clip();
			
			graphics.t(this.CurrentTranslate.StylesText, 0.5, (y + b) / 2);
			
			ctx.setTransform(1, 0, 0, 1, 0, 0);
			ctx.fillStyle = "#E8E8E8";
			
			var _b = this.STYLE_THUMBNAIL_HEIGHT - 1.5;
			var _x = 2;
			var _w = this.STYLE_THUMBNAIL_WIDTH - 4;
			var _h = (this.STYLE_THUMBNAIL_HEIGHT / 3) >> 0;
			ctx.beginPath();
			ctx.moveTo(_x, _b - _h);
			ctx.lineTo(_x + _w, _b - _h);
			ctx.lineTo(_x + _w, _b);
			ctx.lineTo(_x, _b);
			ctx.closePath();
			ctx.fill();
			
			ctx.lineWidth = 1;
			ctx.strokeStyle = "#D8D8D8";
			ctx.beginPath();
			ctx.rect(0.5, 0.5, this.STYLE_THUMBNAIL_WIDTH - 1, this.STYLE_THUMBNAIL_HEIGHT - 1);
			
			ctx.stroke();
			
			graphics.restore();
		}
		else
		{
			var hdr = new CHeaderFooter(editor.WordControl.m_oLogicDocument.HdrFtr, editor.WordControl.m_oLogicDocument, editor.WordControl.m_oDrawingDocument, AscCommon.hdrftr_Header);
			var _dc = hdr.Content;//new CDocumentContent(editor.WordControl.m_oLogicDocument, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, false, true, false);
			
			var par = new AscWord.Paragraph(_dc, false);
			var run = new ParaRun(par, false);
			run.AddText(styleName);
			
			_dc.Internal_Content_Add(0, par, false);
			par.Add_ToContent(0, run);
			par.Style_Add(style.Id, false);
			par.Set_Align(AscCommon.align_Left);
			par.Set_Tabs(new CParaTabs());
			
			if (!textPr.Color || (255 === textPr.Color.r && 255 === textPr.Color.g && 255 === textPr.Color.b))
				run.Set_Color(new CDocumentColor(0, 0, 0, false));
			
			var _brdL = style.ParaPr.Brd.Left;
			if (undefined !== _brdL && null !== _brdL)
			{
				var brdL = new CDocumentBorder();
				brdL.Set_FromObject(_brdL);
				brdL.Space = 0;
				par.Set_Border(brdL, AscDFH.historyitem_Paragraph_Borders_Left);
			}
			
			var _brdT = style.ParaPr.Brd.Top;
			if (undefined !== _brdT && null !== _brdT)
			{
				var brd = new CDocumentBorder();
				brd.Set_FromObject(_brdT);
				brd.Space = 0;
				par.Set_Border(brd, AscDFH.historyitem_Paragraph_Borders_Top);
			}
			
			var _brdB = style.ParaPr.Brd.Bottom;
			if (undefined !== _brdB && null !== _brdB)
			{
				var brd = new CDocumentBorder();
				brd.Set_FromObject(_brdB);
				brd.Space = 0;
				par.Set_Border(brd, AscDFH.historyitem_Paragraph_Borders_Bottom);
			}
			
			var _brdR = style.ParaPr.Brd.Right;
			if (undefined !== _brdR && null !== _brdR)
			{
				var brd = new CDocumentBorder();
				brd.Set_FromObject(_brdR);
				brd.Space = 0;
				par.Set_Border(brd, AscDFH.historyitem_Paragraph_Borders_Right);
			}
			
			var _ind = new CParaInd();
			_ind.FirstLine = 0;
			_ind.Left = 0;
			_ind.Right = 0;
			par.Set_Ind(_ind, false);
			
			var _sp = new CParaSpacing();
			_sp.Line = 1;
			_sp.LineRule = Asc.linerule_Auto;
			_sp.Before = 0;
			_sp.BeforeAutoSpacing = false;
			_sp.After = 0;
			_sp.AfterAutoSpacing = false;
			par.Set_Spacing(_sp, false);
			
			_dc.Reset(0, 0, 10000, 10000);
			_dc.Recalculate_Page(0, true);
			
			_dc.Reset(0, 0, par.Lines[0].Ranges[0].W + 0.001, 10000);
			_dc.Recalculate_Page(0, true);
			
			var y = 0;
			var b = dKoefToMM * this.STYLE_THUMBNAIL_HEIGHT;
			var w = dKoefToMM * this.STYLE_THUMBNAIL_WIDTH;
			var off = 10 * dKoefToMM;
			var off2 = 5 * dKoefToMM;
			var off3 = 1 * dKoefToMM;
			
			graphics.transform(1, 0, 0, 1, 0, 0);
			graphics.save();
			graphics._s();
			graphics._m(off2, y + off3);
			graphics._l(w - off, y + off3);
			graphics._l(w - off, b - off3);
			graphics._l(off2, b - off3);
			graphics._z();
			graphics.clip();
			
			//graphics.t(style.Name, off + 0.5, y + 0.75 * (b - y));
			var baseline = par.Lines[0].Y;
			par.Shift(0, off + 0.5, y + 0.75 * (b - y) - baseline);
			par.Draw(0, graphics);
			
			graphics.restore();
		}
	};
	StylePreviewGenerator.prototype.GenerateDefaultStyles = function ()
	{
		for (let i = 0; i < this.styles.length; ++i)
		{
			let style = this.styles[i];
			let styleName = style.GetName();
			if (null === this.styleManager.GetStyleIdByName(styleName))
			{
				this.drawStyle(this.api, this.graphics, style, AscCommon.translateManager.getValue(styleName));
				this.defaultStyles.push(new AscCommon.CStyleImage(styleName, AscCommon.c_oAscStyleImage.Default, this.canvas.toDataURL("image/png"), style.uiPriority));
			}
		}
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].StylePreviewGenerator = StylePreviewGenerator;
})(window);

window['AscCommonWord'].CStylesPainter = CStylesPainter;
CStylesPainter.prototype['get_MergedStyles'] = CStylesPainter.prototype.get_MergedStyles;
