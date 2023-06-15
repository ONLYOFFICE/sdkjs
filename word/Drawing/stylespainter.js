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
var FontStyle = AscFonts.FontStyle;
var g_fontApplication = AscFonts.g_fontApplication;

var CColor = AscCommon.CColor;
var CAscMathCategory = AscCommon.CAscMathCategory;
var g_oTableId = AscCommon.g_oTableId;
var g_oTextMeasurer = AscCommon.g_oTextMeasurer;
var global_mouseEvent = AscCommon.global_mouseEvent;
var History = AscCommon.History;
var global_MatrixTransformer = AscCommon.global_MatrixTransformer;
var g_dKoef_pix_to_mm = AscCommon.g_dKoef_pix_to_mm;
var g_dKoef_mm_to_pix = AscCommon.g_dKoef_mm_to_pix;

function CStylesPainter()
{
	this.defaultStyles = null;
	this.docStyles = null;

	this.mergedStyles = null;

	this.STYLE_THUMBNAIL_WIDTH = GlobalSkin.STYLE_THUMBNAIL_WIDTH;
	this.STYLE_THUMBNAIL_HEIGHT = GlobalSkin.STYLE_THUMBNAIL_HEIGHT;

	this.CurrentTranslate = null;
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

CStylesPainter.prototype.GenerateStyles = function (_api, ds)
{
	var _oldX = this.STYLE_THUMBNAIL_WIDTH;
	var _oldY = this.STYLE_THUMBNAIL_HEIGHT;

	this.STYLE_THUMBNAIL_WIDTH 	= AscCommon.AscBrowser.convertToRetinaValue(this.STYLE_THUMBNAIL_WIDTH, true);
	this.STYLE_THUMBNAIL_HEIGHT = AscCommon.AscBrowser.convertToRetinaValue(this.STYLE_THUMBNAIL_HEIGHT, true);

	this.CurrentTranslate = _api.CurrentTranslate;

	this.GenerateDefaultStyles(_api, ds);
	this.GenerateDocumentStyles(_api);

	// стили сформированы. осталось просто сформировать единый список
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

	this.mergedStyles = [];
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
			this.mergedStyles.push(aSubArray[i]);
		}
	}

	this.STYLE_THUMBNAIL_WIDTH = _oldX;
	this.STYLE_THUMBNAIL_HEIGHT = _oldY;

	// export
	this["STYLE_THUMBNAIL_WIDTH"] = this.STYLE_THUMBNAIL_WIDTH;
	this["STYLE_THUMBNAIL_HEIGHT"] = this.STYLE_THUMBNAIL_HEIGHT;

	// теперь просто отдаем евент наверх
	_api.sync_InitEditorStyles(this);
};

CStylesPainter.prototype.GenerateDefaultStyles = function (_api, ds)
{
	var styles = ds;

	// добавили переводы => нельзя кэшировать
	var _canvas = document.createElement('canvas');
	_canvas.width = this.STYLE_THUMBNAIL_WIDTH;
	_canvas.height = this.STYLE_THUMBNAIL_HEIGHT;
	var ctx = _canvas.getContext('2d');

	ctx.fillStyle = "#FFFFFF";
	ctx.fillRect(0, 0, _canvas.width, _canvas.height);

	var koef = AscCommon.g_dKoef_pix_to_mm / AscCommon.AscBrowser.retinaPixelRatio;

	var graphics = new AscCommon.CGraphics();
	graphics.init(ctx, _canvas.width, _canvas.height, _canvas.width * koef, _canvas.height * koef);
	graphics.m_oFontManager = AscCommon.g_fontManager;

	var DocumentStyles = _api.WordControl.m_oLogicDocument.Get_Styles();
	this.defaultStyles = [];
	for (var i in styles)
	{
		var style = styles[i];
		if (style.IsExpressStyle(DocumentStyles) && null === DocumentStyles.GetStyleIdByName(style.Name))
		{
			this.drawStyle(_api, graphics, style, AscCommon.translateManager.getValue(style.Name));
			this.defaultStyles.push(new AscCommon.CStyleImage(style.Name, AscCommon.c_oAscStyleImage.Default,
				_canvas.toDataURL("image/png"), style.uiPriority));
		}
	}
};

CStylesPainter.prototype.GenerateDocumentStyles = function(_api)
{
	if (!_api.WordControl.m_oLogicDocument)
		return;

	var DocumentStyles = _api.WordControl.m_oLogicDocument.Get_Styles();
	var styles = DocumentStyles.Style;

	if (!styles)
		return;

	var cur_index = 0;

	var _canvas = document.createElement('canvas');
	_canvas.width = this.STYLE_THUMBNAIL_WIDTH;
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

	this.docStyles = [];
	for (var i in styles)
	{
		var style = styles[i];
		if (style.IsExpressStyle(DocumentStyles))
		{
			// как только меняется сериалайзер - меняется и код здесь. Да, не очень удобно,
			// зато быстро делается
			var formalStyle = i.toLowerCase().replace(/\s/g, "");
			var res = formalStyle.match(/^heading([1-9][0-9]*)$/);
			var index = (res) ? res[1] - 1 : -1;

			var _dr_style = DocumentStyles.Get_Pr(i, styletype_Paragraph);
			_dr_style.Name = style.Name;
			_dr_style.Id = i;

			this.drawStyle(_api, graphics, _dr_style,
				DocumentStyles.IsStyleDefaultByName(style.Name) ? AscCommon.translateManager.getValue(style.Name) : style.Name);
			this.docStyles[cur_index] = new AscCommon.CStyleImage(style.Name, AscCommon.c_oAscStyleImage.Document,
				_canvas.toDataURL("image/png"), style.uiPriority);

			// алгоритм смены имени
			if (style.Default)
			{
				switch (style.Default)
				{
					case 1:
						break;
					case 2:
						this.docStyles[cur_index].Name = "No List";
						break;
					case 3:
						this.docStyles[cur_index].Name = "Normal";
						break;
					case 4:
						this.docStyles[cur_index].Name = "Normal Table";
						break;
				}
			}
			else if (index != -1)
			{
				this.docStyles[cur_index].Name = "Heading ".concat(index + 1);
			}

			cur_index++;
		}
	}
};

CStylesPainter.prototype.drawStyle = function(_api, graphics, style, styleName)
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
		g_oTableId.m_bTurnOff = true;
		History.TurnOff();

		var oldDefTabStop = AscCommonWord.Default_Tab_Stop;
		AscCommonWord.Default_Tab_Stop = 1;
		var isShowParaMarks = false;
		if (_api)
		{
			isShowParaMarks = _api.get_ShowParaMarks();
			_api.put_ShowParaMarks(false);
		}

		var hdr = new CHeaderFooter(editor.WordControl.m_oLogicDocument.HdrFtr, editor.WordControl.m_oLogicDocument, editor.WordControl.m_oDrawingDocument, AscCommon.hdrftr_Header);
		var _dc = hdr.Content;//new CDocumentContent(editor.WordControl.m_oLogicDocument, editor.WordControl.m_oDrawingDocument, 0, 0, 0, 0, false, true, false);

		var par = new Paragraph(editor.WordControl.m_oDrawingDocument, _dc, false);
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
		//par.Reset(0, 0, 10000, 10000, 0);
		//par.Recalculate_Page(0);

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

		if (_api)
			_api.put_ShowParaMarks(isShowParaMarks);

		AscCommonWord.Default_Tab_Stop = oldDefTabStop;

		g_oTableId.m_bTurnOff = false;
		History.TurnOn();
	}
};

CStylesPainter.prototype.get_MergedStyles = function ()
{
	return this.mergedStyles;
};

window['AscCommonWord'].CStylesPainter = CStylesPainter;
CStylesPainter.prototype['get_MergedStyles'] = CStylesPainter.prototype.get_MergedStyles;
