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

// Import
var CellValueType = AscCommon.CellValueType;
var c_oAscBorderWidth = AscCommon.c_oAscBorderWidth;
var c_oAscBorderStyles = Asc.c_oAscBorderStyles;
var FormulaTablePartInfo = AscCommon.FormulaTablePartInfo;
var parserHelp = AscCommon.parserHelp;
var gc_nMaxRow0 = AscCommon.gc_nMaxRow0;
var gc_nMaxCol0 = AscCommon.gc_nMaxCol0;
	var c_oAscPrintDefaultSettings = AscCommon.c_oAscPrintDefaultSettings;

var UndoRedoDataTypes = AscCommonExcel.UndoRedoDataTypes;
var UndoRedoData_IndexSimpleProp = AscCommonExcel.UndoRedoData_IndexSimpleProp;
var UndoRedoData_FromTo = AscCommonExcel.UndoRedoData_FromTo;

var UndoRedoData_Layout = AscCommonExcel.UndoRedoData_Layout;


var c_oAscCustomAutoFilter = Asc.c_oAscCustomAutoFilter;
var c_oAscAutoFilterTypes = Asc.c_oAscAutoFilterTypes;

let cDate = Asc.cDate;

var c_maxOutlineLevel = 7;

var g_oColorManager = null;
	
var g_nHSLMaxValue = 255;
var g_nColorTextDefault = 1;
var g_nColorHyperlink = 10;
var g_nColorHyperlinkVisited = 11;

var previewConditionalFormattingNum = 38718;

var g_oThemeColorsDefaultModsSpreadsheet = [
    [0, -4.9989318521683403E-2, -0.14999847407452621, -0.249977111117893, -0.34998626667073579, -0.499984740745262],
    [0, -9.9978637043366805E-2, -0.249977111117893, -0.499984740745262, -0.749992370372631, -0.89999084444715716],
    [0, 0.79998168889431442, 0.59999389629810485, 0.39997558519241921, -0.249977111117893, -0.499984740745262],
    [0, 0.89999084444715716, 0.749992370372631, 0.499984740745262, 0.249977111117893, 9.9978637043366805E-2],
    [0, 0.499984740745262, 0.34998626667073579, 0.249977111117893, 0.14999847407452621, 4.9989318521683403E-2]];

var map_themeExcel_to_themePresentation = {
	0: 12,
	1: 8,
	2: 13,
	3: 9,
	4: 0,
	5: 1,
	6: 2,
	7: 3,
	8: 4,
	9: 5,
	10: 11,
	11: 10
};
function shiftGetBBox(bbox, bHor)
{
	var bboxGet = null;
	if(bHor)
		bboxGet = new Asc.Range(bbox.c1, bbox.r1, gc_nMaxCol0, bbox.r2);
	else
		bboxGet = new Asc.Range(bbox.c1, bbox.r1, bbox.c2, gc_nMaxRow0);
	return bboxGet;
}
function shiftSort(a, b, offset)
{
	var nRes = 0;
	if(null == a.to || null == b.to)
	{
		if(null == a.to && null == b.to)
			nRes = 0;
		else if(null == a.to)
			nRes = -1;
		else if(null == b.to)
			nRes = 1;
	}
	else
	{
	    if (0 != offset.row) {
	        if (offset.row > 0)
	            nRes = b.to.r1 - a.to.r1;
	        else
	            nRes = a.to.r1 - b.to.r1;
	    }
	    if (0 == nRes && 0 != offset.col) {
	        if (offset.col > 0)
	            nRes = b.to.c1 - a.to.c1;
	        else
	            nRes = a.to.c1 - b.to.c1;
	    }
	}
	return nRes;
}
function createRgbColor(r, g, b) {
	return new RgbColor((r << 16) + (g << 8) + b);
}

function compareValues(val1, val2) {
	if (val1 === val2) {
		return true;
	}
	if ((val1 == null || val1 == "") && (val2 == null || val2 == "")) {
		return true;
	}

	return false;
}
var g_oRgbColorProperties = {
		rgb : 0
	};
function RgbColor(rgb)
{
	this.rgb = rgb;

	this._hash;
}
RgbColor.prototype =
{
	Properties: g_oRgbColorProperties,
	getHash: function() {
		if (!this._hash) {
			this._hash = this.rgb;
		}
		return this._hash;
	},
	clone : function()
	{
		return new RgbColor(this.rgb);
	},
	getType : function()
	{
		return UndoRedoDataTypes.RgbColor;
	},
	getProperties : function()
	{
		return this.Properties;
	},

    isEqual: function(oColor)
    {
        if(!oColor || !(oColor instanceof RgbColor)){
            return false;
        }
        if(this.rgb !== oColor.rgb){
            return false;
        }
        return true;
    },

	getProperty : function(nType)
	{
		switch(nType)
		{
		case this.Properties.rgb:return this.rgb;break;
		}
	},
	setProperty : function(nType, value)
	{
		switch(nType)
		{
		case this.Properties.rgb: this.rgb = value;break;
		}
	},
	Write_ToBinary2 : function(oBinaryWriter)
	{
		oBinaryWriter.WriteLong(this.rgb);
	},
	Read_FromBinary2 : function(oBinaryReader)
	{
		this.rgb = oBinaryReader.GetULongLE();
	},
	getRgb : function()
	{
		return this.rgb;
	},
	getR : function()
	{
		return (this.rgb >> 16) & 0xff;
	},
	getG : function()
	{
		return (this.rgb >> 8) & 0xff;
	},
	getB : function()
	{
		return this.rgb & 0xff;
	},
	getA : function () {
		return 1;
	}
};
var g_oThemeColorProperties = {
		rgb: 0,
		theme: 1,
		tint: 2
	};
function ThemeColor()
{
	this.rgb = null;
	this.theme = null;
	this.tint = null;

	this._hash;
}
ThemeColor.prototype =
{
	Properties: g_oThemeColorProperties,
	getHash: function() {
		if (!this._hash) {
			this._hash = this.theme + ';' + this.tint;
		}
		return this._hash;
	},
	clone : function()
	{
		//ThemeColor must be created by g_oColorManager for correct rebuild
		//no need getThemeColor because it return same object
		return this;
	},
	getType : function()
	{
		return UndoRedoDataTypes.ThemeColor;
	},
	getProperties : function()
	{
		return this.Properties;
	},
	getProperty : function(nType)
	{
		switch(nType)
		{
		case this.Properties.rgb:return this.rgb;break;
		case this.Properties.theme:return this.theme;break;
		case this.Properties.tint:return this.tint;break;
		}
	},
	setProperty : function(nType, value)
	{
		switch(nType)
		{
		case this.Properties.rgb: this.rgb = value;break;
		case this.Properties.theme: this.theme= value;break;
		case this.Properties.tint: this.tint = value;break;
		}
	},
    isEqual: function(oColor)
    {
        if(!oColor){
            return false;
        }
        if(this.theme !== oColor.theme){
            return false;
        }
        if(!AscFormat.fApproxEqual(this.tint, oColor.tint)){
            return false;
        }
        return true;
    },
	Write_ToBinary2 : function(oBinaryWriter)
	{
		oBinaryWriter.WriteByte(this.theme);
		if(null != this.tint)
		{
			oBinaryWriter.WriteByte(true);
			oBinaryWriter.WriteDouble2(this.tint);
		}
		else
		{
			oBinaryWriter.WriteBool(false);
		}
	},
	Read_FromBinary2AndReplace : function(oBinaryReader)
	{
		this.theme = oBinaryReader.GetUChar();
		var bTint = oBinaryReader.GetBool();
		if(bTint)
			this.tint = oBinaryReader.GetDoubleLE();
		return g_oColorManager.getThemeColor(this.theme, this.tint);
	},
	getRgb : function()
	{
		return this.rgb;
	},
	getR : function()
	{
		return (this.rgb >> 16) & 0xff;
	},
	getG : function()
	{
		return (this.rgb >> 8) & 0xff;

	},
	getB : function()
	{
		return this.rgb & 0xff;
	},
	getA : function () {
		return 1;
	},
	rebuild : function(theme)
	{
		var nRes = 0;
		var r = 0;
		var g = 0;
		var b = 0;
		if(null != this.theme && null != theme)
		{
			var oUniColor = theme.themeElements.clrScheme.colors[map_themeExcel_to_themePresentation[this.theme]];
			if(null != oUniColor)
			{
				var rgba = oUniColor.color.RGBA;
				if(null != rgba)
				{
					r = rgba.R;
					g = rgba.G;
					b = rgba.B;
				}
			}
			if(null != this.tint && 0 != this.tint)
			{
				var oCColorModifiers = new AscFormat.CColorModifiers();
				var HSL = {H: 0, S: 0, L: 0};
				oCColorModifiers.RGB2HSL(r, g, b, HSL);
				if (this.tint < 0)
					HSL.L = HSL.L * (1 + this.tint);
				else
					HSL.L = HSL.L * (1 - this.tint) + (g_nHSLMaxValue - g_nHSLMaxValue * (1 - this.tint));
				HSL.L >>= 0;
				var RGB = {R: 0, G: 0, B: 0};
				oCColorModifiers.HSL2RGB(HSL, RGB);
				r = RGB.R;
				g = RGB.G;
				b = RGB.B;
			}
			nRes |= b;
			nRes |= g << 8;
			nRes |= r << 16;
		}
		this.rgb = nRes;
	}
};
function CorrectAscColor(asc_color)
{
	if (null == asc_color || asc_color.asc_getAuto())
		return null;

	var ret = null;

	var _type = asc_color.asc_getType();
	switch (_type)
	{
		case Asc.c_oAscColor.COLOR_TYPE_SCHEME:
		{
			// тут выставляется ТОЛЬКО из меню. поэтому:
			var _index = asc_color.asc_getValue() >> 0;
			var _id = (_index / 6) >> 0;
			var _pos = _index - _id * 6;
			var basecolor = g_oColorManager.getThemeColor(_id);
			var aTints = g_oThemeColorsDefaultModsSpreadsheet[AscCommon.GetDefaultColorModsIndex(basecolor.getR(), basecolor.getG(), basecolor.getB())];
			var tint = aTints[_pos];
			ret = g_oColorManager.getThemeColor(_id, tint);
			break;
		}
		default:
		{
			ret = createRgbColor(asc_color.asc_getR(), asc_color.asc_getG(), asc_color.asc_getB());
		}
	}
	return ret;
}
function ColorManager()
{
	this.theme = null;
	this.aColors = new Array(12);
}
ColorManager.prototype =
{
	isEqual : function(color1, color2, byRgb)
	{
		var bRes = false;
		if(null == color1 && null == color2)
			bRes = true;
		else if(null != color1 && null != color2)
		{
			if((color1 instanceof ThemeColor && color2 instanceof ThemeColor) || (color1 instanceof RgbColor && color2 instanceof RgbColor) || byRgb)
				bRes =  color1.getRgb() == color2.getRgb();
		}
		return bRes;
	},
	setTheme : function(theme)
	{
		this.theme = theme;
		this.rebuildColors();
	},
	getThemeColor : function(theme, tint)
	{
		if(null == tint)
			tint = null;
		var oColorObj = this.aColors[theme];
		if(null == oColorObj)
		{
			oColorObj = {};
			this.aColors[theme] = oColorObj;
		}
		var oThemeColor = oColorObj[tint];
		if(null == oThemeColor)
		{
			oThemeColor = new ThemeColor();
			oThemeColor.theme = theme;
			oThemeColor.tint = tint;
			if(null != this.theme)
				oThemeColor.rebuild(this.theme);
			oColorObj[tint] = oThemeColor;
		}
		return oThemeColor;
	},
	rebuildColors : function()
	{
		if(null != this.theme)
		{
			for(var i = 0, length = this.aColors.length; i < length; ++i)
			{
				var oColorObj = this.aColors[i];
				for(var j in oColorObj)
				{
					var oThemeColor = oColorObj[j];
					oThemeColor.rebuild(this.theme);
				}
			}
		}
	}
};
g_oColorManager = new ColorManager();

	var g_oDefaultFormat = {
		XfId: null,
		Font: null,
		Fill: null,
		Num: null,
		Border: null,
		Align: null,
		NumAbs: null,
		AlignAbs: null,
		ColorAuto: new RgbColor(0),

		xfs: new CellXfs()
	};

	/** @constructor */
	function Fragment(val) {
		this.text = null;
		this.format = null;
		//для отрисовки ввожу дополнительный массив
		this.charCodes = null;//[]

		if (null != val) {
			this.set(val);
		}
	}

	Fragment.prototype.clone = function () {
		return new Fragment(this);
	};
	Fragment.prototype.set = function (oVal) {
		if (null != oVal.text) {
			this.setFragmentText(oVal.text);
		}
		if (null != oVal.format) {
			this.format = oVal.format;
		}
		if (null != oVal.charCodes) {
			this.setCharCodes(oVal.charCodes && oVal.charCodes.slice());
		}
	};
	Fragment.prototype.checkVisitedHyperlink = function (row, col, hyperlinkManager) {
		var color = this.format.getColor();
		if (color instanceof AscCommonExcel.ThemeColor && g_nColorHyperlink === color.theme && null === color.tint) {
			//для посещенных гиперссылок
			var hyperlink = hyperlinkManager.getByCell(row, col);
			if (hyperlink && hyperlink.data.getVisited()) {
				this.format.setColor(g_oColorManager.getThemeColor(g_nColorHyperlinkVisited, null));
			}
		}
	};
	Fragment.prototype.getCharCode = function (index) {
		if (!this.isInitCharCodes()) {
			this.initCharCodes();
		}
		return this.charCodes && this.charCodes[index];
	};
	Fragment.prototype.getCharCodesLength = function () {
		if (!this.isInitCharCodes()) {
			this.initCharCodes();
		}
		return this.charCodes ? this.charCodes.length : 0;
	};
	Fragment.prototype.initCharCodes = function () {
		var test2 = this.getFragmentText();
		if (test2) {
			this.setCharCodes(AscCommon.convertUTF16toUnicode(test2), true);
		} else {
			this.setCharCodes([], true);
		}
	};
	Fragment.prototype.initText = function () {
		this.setFragmentText(this.charCodes ? AscCommon.convertUnicodeToUTF16(this.charCodes) : "", true);
	};
	Fragment.prototype.isInitCharCodes = function () {
		return this.charCodes !== null;
	};
	Fragment.prototype.getCharCodes = function () {
		if (!this.isInitCharCodes()) {
			this.initCharCodes();
		}
		return this.charCodes;
	};
	Fragment.prototype.setCharCodes = function (val, isInit) {
		//если выставляем charCodes, контент меняется, нужно занулять текстовое поле
		if (!isInit) {
			this.text = null;
		}
		this.charCodes = val;
	};
	Fragment.prototype.getFragmentText = function () {
		if (null === this.text) {
			this.initText();
		}
		return this.text;
	};
	Fragment.prototype.setFragmentText = function (val, isInit) {
		//если выставляем текстовое поле, контент меняется, нужно занулять charCodes
		if (!isInit) {
			this.charCodes = null;
		}
		this.text = val;
	};
	Fragment.prototype.convertPositionToText = function (codePos) {
		var diff = 0;
		for (var i = 0; i < codePos; i++) {
			if (this.charCodes[i] >= 0x10000) {
				diff++;
			}
		}
		return codePos + diff;
	};
	Fragment.prototype.convertPositionFromText = function (textPos) {
		var count = 0;
		for (var i = 0; i < textPos; i++) {
			if (this.charCodes[i] >= 0x10000) {
				count++;
			} else {
				count += 2;
			}
			if (count >= textPos) {
				return i;
			}
		}
		return textPos;
	};


var g_oFontProperties = {
		fn: 0,
		scheme: 1,
		fs: 2,
		b: 3,
		i: 4,
		u: 5,
		s: 6,
		c: 7,
		va: 8
	};

	/** @constructor */
	function Font() {
		this.fn = null;
		this.scheme = null;
		this.fs = null;
		this.b = null;
		this.i = null;
		this.u = null;
		this.s = null;
		this.c = null;
		this.va = null;
		this.skip = null;
		this.repeat = null;

		this._hash;
		this._index;
	}
	Font.prototype.Properties = g_oFontProperties;
	Font.prototype.getHash = function() {
		if(!this._hash){
			var color = this.c ? this.c.getHash() : '';
			this._hash = this.fn + '|' +this.scheme + '|' +this.fs + '|' +this.b + '|' +this.i + '|' +this.u + '|' +
				this.s + '|' + color + '|' +this.va + '|' + this.skip + '|' +this.repeat;
		}
		return this._hash;
	};
	Font.prototype.getIndexNumber = function() {
		return this._index;
	};
	Font.prototype.setIndexNumber = function(val) {
		return this._index = val;
	};
	Font.prototype.initDefault = function(wb) {
		if (!this.fn) {
			var sThemeFont = null;
			if (null != wb.theme.themeElements && null != wb.theme.themeElements.fontScheme) {
				if (Asc.EFontScheme.fontschemeMinor == this.scheme && wb.theme.themeElements.fontScheme.minorFont) {
					sThemeFont = wb.theme.themeElements.fontScheme.minorFont.latin;
				} else if (Asc.EFontScheme.fontschemeMajor == this.scheme && wb.theme.themeElements.fontScheme.majorFont) {
					sThemeFont = wb.theme.themeElements.fontScheme.majorFont.latin;
				}
			}
			this.fn = sThemeFont ? sThemeFont : "Calibri";
		}
		if (!this.fs) {
			this.fs = 11;
		}
		if (!this.c) {
			this.c = AscCommonExcel.g_oColorManager.getThemeColor(AscCommonExcel.g_nColorTextDefault);
		}
	};
	Font.prototype.assign = function(font) {
		this.fn = font.fn;
		this.scheme = font.scheme;
		this.fs = font.fs;
		this.b = font.b;
		this.i = font.i;
		this.u = font.u;
		this.s = font.s;
		this.c = font.c;
		this.va = font.va;
		this.skip = font.skip;
		this.repeat = font.repeat;
	};
	Font.prototype.assignFromObject = function(font) {
		if (null != font.fn) {
			this.setName(font.fn);
		}
		if (null != font.scheme) {
			this.setScheme(font.scheme);
		}
		if (null != font.fs) {
			this.setSize(font.fs);
		}
		if (null != font.b) {
			this.setBold(font.b);
		}
		if (null != font.i) {
			this.setItalic(font.i);
		}
		if (null != font.u) {
			this.setUnderline(font.u);
		}
		if (null != font.s) {
			this.setStrikeout(font.s);
		}
		if (null != font.c) {
			this.setColor(font.c);
		}
		if (null != font.va) {
			this.setVerticalAlign(font.va);
		}
		if (null != font.skip) {
			this.setSkip(font.skip);
		}
		if (null != font.repeat) {
			this.setRepeat(font.repeat);
		}
	};
    Font.prototype.assignFromTextPr = function (textPr) {
        if (textPr.FontFamily) {
            this.setName(textPr.FontFamily.Name);
        }
        this.setSize(textPr.FontSize);
        this.setBold(textPr.Bold);
        this.setItalic(textPr.Italic);
        this.setUnderline(textPr.Underline === true ? Asc.EUnderline.underlineSingle : Asc.EUnderline.underlineNone);
        this.setStrikeout(textPr.Strikeout);
        this.setVerticalAlign(textPr.VertAlign);
        if (textPr.Unifill) {
            var color = textPr.Unifill.getRGBAColor();
            this.setColor(createRgbColor(color.R, color.G, color.B));
        } else if (textPr.Color) {
            this.setColor(createRgbColor(textPr.Color.r, textPr.Color.g, textPr.Color.b));
        }
    };
	Font.prototype.merge = function (font, isTable, isTableColor) {
		var oRes = new Font();
		oRes.fn = this.fn || font.fn;
		oRes.scheme = this.scheme || font.scheme;
		oRes.fs = this.fs || font.fs;

		oRes.b = null !== this.b ? this.b : font.b;
		if (isTable) {
			oRes.i = null !== font.i ? font.i : this.i;
			oRes.s = null !== font.s ? font.s : this.s;
			oRes.u = null !== font.u ? font.u : this.u;
			oRes.va = font.va || this.va;
		} else {
			oRes.i = null !== this.i ? this.i : font.i;
			oRes.s = null !== this.s ? this.s : font.s;
			oRes.u = null !== this.u ? this.u : font.u;
			oRes.va = this.va || font.va;
		}
		if (isTable) {
			if (isTableColor) {
				oRes.c = font.c || this.c;
			} else {
				oRes.c = this.c;
			}
		} else {
			oRes.c = this.c || font.c;
		}
		oRes.skip = this.skip || font.skip;
		oRes.repeat = this.repeat || font.repeat;
		return oRes;
	};
	Font.prototype.isEqual = function (font) {
		var bRes = this.fs == font.fs && this.b == font.b && this.i == font.i && this.u == font.u && this.s == font.s &&
			g_oColorManager.isEqual(this.c, font.c) && this.va == font.va && this.skip == font.skip &&
			this.repeat == font.repeat;
		if (bRes) {
			var schemeThis = this.getScheme();
			var schemeOther = font.getScheme();
			if (Asc.EFontScheme.fontschemeNone == schemeThis && Asc.EFontScheme.fontschemeNone == schemeOther) {
				bRes = this.fn == font.fn;
			} else if (Asc.EFontScheme.fontschemeNone != schemeThis &&
				Asc.EFontScheme.fontschemeNone != schemeOther) {
				bRes = schemeThis == schemeOther;
			} else {
				bRes = false;
			}
		}
		return bRes;
	};
	Font.prototype.isEqual2 = function (font) {
		return font && this.getName() === font.getName() && this.getSize() === font.getSize() && this.getBold() === font.getBold() && this.getItalic() === font.getItalic();
	};
	Font.prototype.isNormalXfColor = function () {
		return this.c && this.c.isEqual(g_StyleCache.normalXf.font.c);
	};
	Font.prototype.clone = function () {
		var font = new Font();
		font.assign(this);
		return font;
	};
	Font.prototype.intersect = function (oFont, oDefVal) {
		if (this.fn != oFont.fn) {
			this.fn = oDefVal.fn;
		}
		if (this.scheme != oFont.scheme) {
			this.scheme = oDefVal.scheme;
		}
		if (this.fs != oFont.fs) {
			this.fs = oDefVal.fs;
		}
		if (this.b != oFont.b) {
			this.b = oDefVal.b;
		}
		if (this.i != oFont.i) {
			this.i = oDefVal.i;
		}
		if (this.u != oFont.u) {
			this.u = oDefVal.u;
		}
		if (this.s != oFont.s) {
			this.s = oDefVal.s;
		}
		if (false == g_oColorManager.isEqual(this.c, oFont.c)) {
			this.c = oDefVal.c;
		}
		if (this.va != oFont.va) {
			this.va = oDefVal.va;
		}
		if (this.skip != oFont.skip) {
			this.skip = oDefVal.skip;
		}
		if (this.repeat != oFont.repeat) {
			this.repeat = oDefVal.repeat;
		}
	};
	Font.prototype.subtractEqual = function (oFont, oDefVal) {
		if (this.fn === oFont.fn) {
			this.fn = oDefVal.fn;
		}
		if (this.scheme === oFont.scheme) {
			this.scheme = oDefVal.scheme;
		}
		if (this.fs === oFont.fs) {
			this.fs = oDefVal.fs;
		}
		if (this.b === oFont.b) {
			this.b = oDefVal.b;
		}
		if (this.i === oFont.i) {
			this.i = oDefVal.i;
		}
		if (this.u === oFont.u) {
			this.u = oDefVal.u;
		}
		if (this.s === oFont.s) {
			this.s = oDefVal.s;
		}
		if (g_oColorManager.isEqual(this.c, oFont.c)) {
			this.c = oDefVal.c;
		}
		if (this.va === oFont.va) {
			this.va = oDefVal.va;
		}
		if (this.skip === oFont.skip) {
			this.skip = oDefVal.skip;
		}
		if (this.repeat === oFont.repeat) {
			this.repeat = oDefVal.repeat;
		}
	};
	Font.prototype.getName = function () {
		return this.fn || g_oDefaultFormat.Font.fn;
	};
	Font.prototype.setName = function (val) {
		return this.fn = val;
	};
	Font.prototype.getScheme = function () {
		return this.scheme || Asc.EFontScheme.fontschemeNone;
	};
	Font.prototype.setScheme = function(val) {
		return (null != val && Asc.EFontScheme.fontschemeNone != val) ? this.scheme = val : this.scheme = null;
	};
	Font.prototype.getSize = function () {
		return this.fs || g_oDefaultFormat.Font.fs;
	};
	Font.prototype.setSize = function(val) {
		return this.fs = val;
	};
	Font.prototype.getBold = function () {
		return !!this.b;
	};
	Font.prototype.setBold = function(val) {
		return val ? this.b = true : this.b = null;
	};
	Font.prototype.getItalic = function () {
		return !!this.i;
	};
	Font.prototype.setItalic = function(val) {
		return val ? this.i = true : this.i = null;
	};
	Font.prototype.getUnderline = function () {
		return null != this.u ? this.u : Asc.EUnderline.underlineNone;
	};
	Font.prototype.setUnderline = function(val) {
		return (null != val && Asc.EUnderline.underlineNone != val) ? this.u = val : this.u = null;
	};
	Font.prototype.getStrikeout = function () {
		return !!this.s;
	};
	Font.prototype.setStrikeout = function(val) {
		return val ? this.s = true : this.s = null;
	};
	Font.prototype.getColor = function () {
		return this.c || g_oDefaultFormat.ColorAuto;
	};
	Font.prototype.getColorNotDefault = function () {
		return this.c;
	};
	Font.prototype.setColor = function(val) {
		return this.c = val;
	};
	Font.prototype.getVerticalAlign = function () {
		return null != this.va ? this.va : AscCommon.vertalign_Baseline;
	};
	Font.prototype.setVerticalAlign = function(val) {
		return (null != val && AscCommon.vertalign_Baseline != val) ? this.va = val : this.va = null;
	};
	Font.prototype.getSkip = function () {
		return !!this.skip;
	};
	Font.prototype.setSkip = function(val) {
		return val ? this.skip = true : this.skip = null;
	};
	Font.prototype.getRepeat = function () {
		return !!this.repeat;
	};
	Font.prototype.setRepeat = function (val) {
		return val ? this.repeat = true : this.repeat = null;
	};
	Font.prototype.getType = function () {
		return UndoRedoDataTypes.StyleFont;
	};
	Font.prototype.getProperties = function () {
		return this.Properties;
	};
	Font.prototype.getProperty = function (nType) {
		switch (nType) {
			case this.Properties.fn:
				return this.fn;
			case this.Properties.scheme:
				return this.scheme;
			case this.Properties.fs:
				return this.fs;
			case this.Properties.b:
				return this.b;
			case this.Properties.i:
				return this.i;
			case this.Properties.u:
				return this.u;
			case this.Properties.s:
				return this.s;
			case this.Properties.c:
				return this.c;
			case this.Properties.va:
				return this.va;
		}
	};
	Font.prototype.setProperty = function (nType, value) {
		switch (nType) {
			case this.Properties.fn:
				this.fn = value;
				break;
			case this.Properties.scheme:
				this.scheme = value;
				break;
			case this.Properties.fs:
				this.fs = value;
				break;
			case this.Properties.b:
				this.b = value;
				break;
			case this.Properties.i:
				this.i = value;
				break;
			case this.Properties.u:
				this.u = value;
				break;
			case this.Properties.s:
				this.s = value;
				break;
			case this.Properties.c:
				this.c = value;
				break;
			case this.Properties.va:
				this.va = value;
				break;
		}
	};
	Font.prototype.onStartNode = function(elem, attr, uq) {
		var newContext = this;
		if("b" === elem){
			this.b = AscCommon.getBoolFromXml(AscCommon.readValAttr(attr));
		} else if("color" === elem){
			this.c = AscCommon.getColorFromXml(attr);
		} else if("i" === elem){
			this.i = AscCommon.getBoolFromXml(AscCommon.readValAttr(attr));
		} else if("name" === elem){
			this.fn = AscCommon.readValAttr(attr);
		} else if("scheme" === elem){
			this.scheme = AscCommon.readValAttr(attr);
		} else if("strike" === elem){
			this.s = AscCommon.getBoolFromXml(AscCommon.readValAttr(attr));
		} else if("sz" === elem){
			this.fs = AscCommon.getNumFromXml(AscCommon.readValAttr(attr));
		} else if("u" === elem){
			switch (AscCommon.readValAttr(attr)) {
				case "single":
					this.u = Asc.EUnderline.underlineSingle;
					break;
				case "double":
					this.u = Asc.EUnderline.underlineDouble;
					break;
				case "singleAccounting":
					this.u = Asc.EUnderline.underlineSingleAccounting;
					break;
				case "doubleAccounting":
					this.u = Asc.EUnderline.underlineDoubleAccounting;
					break;
				case "none":
					this.u = Asc.EUnderline.underlineNone;
					break;
			}
		} else if("vertAlign" === elem){
			switch (AscCommon.readValAttr(attr)) {
				case "baseline":
					this.va = AscCommon.vertalign_Baseline;
					break;
				case "superscript":
					this.va = AscCommon.vertalign_SuperScript;
					break;
				case "subscript":
					this.va = AscCommon.vertalign_SubScript;
					break;
			}
		} else {
			newContext = null;
		}
		return newContext;
	};
	Font.prototype.checkSchemeFont = function(theme) {
		if (null != this.scheme && theme) {
			var fontScheme = theme.themeElements.fontScheme;
			var sFontName = null;
			switch (this.scheme) {
				case Asc.EFontScheme.fontschemeMinor:sFontName = fontScheme.minorFont.latin;break;
				case Asc.EFontScheme.fontschemeMajor:sFontName = fontScheme.majorFont.latin;break;
			}
			if (null != sFontName && "" != sFontName) {
				this.fn = sFontName;
			}
		}
	};
	Font.prototype.fromXLSB = function(stream) {
		var dyHeight = stream.GetUShortLE();
		if(dyHeight >= 0x0014) {
			this.fs = dyHeight / 20;
		}
		var grbit = stream.GetUShortLE();
		if(0 !== (grbit & 0x2))
		{
			this.i = true;
		}
		if(0 !== (grbit & 0x8))
		{
			this.s = true;
		}
		var bls = stream.GetUShortLE();
		if(0x02BC == bls) {
			this.b = true;
		}
		var sss = stream.GetUShortLE();
		if(sss > 0)
		{
			switch(sss)
			{
				case 0x0001:
					this.va = AscCommon.vertalign_SuperScript;
					break;
				case 0x0002:
					this.va = AscCommon.vertalign_SubScript;
					break;
			}
		}
		var uls = stream.GetUChar();
		if(uls > 0)
		{
			switch(uls)
			{
				case 0x01:
					this.u =  Asc.EUnderline.underlineSingle;
					break;
				case 0x02:
					this.u =  Asc.EUnderline.underlineDouble;
					break;
				case 0x21:
					this.u =  Asc.EUnderline.underlineSingleAccounting;
					break;
				case 0x22:
					this.u =  Asc.EUnderline.underlineDoubleAccounting;
					break;
			}
		}
		stream.Skip2(3);

		var xColorType = stream.GetUChar();
		var index = stream.GetUChar();
		var nTintAndShade = stream.GetShortLE();
		var rgba = stream.GetULongLE();
		var isActualRgb = 0 !== (xColorType & 0x1);
		xColorType &= 0xFE;
		var tint = null;
		if(0 != nTintAndShade)
		{
			tint = nTintAndShade / 0x7FFF;
		}
		var theme = null;
		if(0x6 == xColorType)
		{
			theme = 1;
			switch(index)
			{
				case 0x01:
					theme = 0;
					break;
				case 0x00:
					theme = 1;
					break;
				case 0x03:
					theme = 2;
					break;
				case 0x02:
					theme = 3;
					break;
				default:
					theme = index;
					break;
			}
			this.c = AscCommonExcel.g_oColorManager.getThemeColor(theme, tint);
		} else if(isActualRgb){
			this.c = AscCommonExcel.createRgbColor((rgba & 0xFF), (rgba & 0xFF00)>>8, (rgba & 0xFF0000)>>16);
		}

		var bFontScheme = stream.GetUChar();
		if(bFontScheme > 0)
		{
			switch(bFontScheme)
			{
				case 0x01:
					this.scheme = Asc.EFontScheme.fontschemeMajor;
					break;
				case 0x02:
					this.scheme = Asc.EFontScheme.fontschemeMinor;
					break;
			}
		}
		this.fn = stream.GetString();
	};
	window["AscCommonExcel"].Font = Font;

	var c_oAscPatternType = {
		DarkDown :  0,
		DarkGray :  1,
		DarkGrid :  2,
		DarkHorizontal :  3,
		DarkTrellis :  4,
		DarkUp :  5,
		DarkVertical :  6,
		Gray0625 :  7,
		Gray125 :  8,
		LightDown :  9,
		LightGray : 10,
		LightGrid : 11,
		LightHorizontal : 12,
		LightTrellis : 13,
		LightUp : 14,
		LightVertical : 15,
		MediumGray : 16,
		None : 17,
		Solid : 18
	};

	function hatchFromExcelToWord(val) {
		switch (val) {
			case c_oAscPatternType.DarkDown:
				return 'dkDnDiag';
			case c_oAscPatternType.DarkGray:
				return 'pct70';
			case c_oAscPatternType.DarkGrid:
				return 'smCheck';
			case c_oAscPatternType.DarkHorizontal:
				return 'dkHorz';
			case c_oAscPatternType.DarkTrellis:
				return 'trellis';
			case c_oAscPatternType.DarkUp:
				return 'dkUpDiag';
			case c_oAscPatternType.DarkVertical:
				return 'dkVert';
			case c_oAscPatternType.Gray0625:
				return 'pct10';
			case c_oAscPatternType.Gray125:
				return 'pct20';
			case c_oAscPatternType.LightDown:
				return 'ltDnDiag';
			case c_oAscPatternType.LightGray:
				return 'pct25';
			case c_oAscPatternType.LightGrid:
				return 'smGrid';
			case c_oAscPatternType.LightHorizontal:
				return 'ltHorz';
			case c_oAscPatternType.LightTrellis:
				return 'pct30';
			case c_oAscPatternType.LightUp:
				return 'ltUpDiag';
			case c_oAscPatternType.LightVertical:
				return 'ltVert';
			case c_oAscPatternType.MediumGray:
			default:
				return 'pct50';
		}
	}

	function FromXml_ST_PatternType(val) {
		var res = -1;
		if ("none" === val) {
			res = c_oAscPatternType.None;
		} else if ("solid" === val) {
			res = c_oAscPatternType.Solid;
		} else if ("mediumGray" === val) {
			res = c_oAscPatternType.MediumGray;
		} else if ("darkGray" === val) {
			res = c_oAscPatternType.DarkGray;
		} else if ("lightGray" === val) {
			res = c_oAscPatternType.LightGray;
		} else if ("darkHorizontal" === val) {
			res = c_oAscPatternType.DarkHorizontal;
		} else if ("darkVertical" === val) {
			res = c_oAscPatternType.DarkVertical;
		} else if ("darkDown" === val) {
			res = c_oAscPatternType.DarkDown;
		} else if ("darkUp" === val) {
			res = c_oAscPatternType.DarkUp;
		} else if ("darkGrid" === val) {
			res = c_oAscPatternType.DarkGrid;
		} else if ("darkTrellis" === val) {
			res = c_oAscPatternType.DarkTrellis;
		} else if ("lightHorizontal" === val) {
			res = c_oAscPatternType.LightHorizontal;
		} else if ("lightVertical" === val) {
			res = c_oAscPatternType.LightVertical;
		} else if ("lightDown" === val) {
			res = c_oAscPatternType.LightDown;
		} else if ("lightUp" === val) {
			res = c_oAscPatternType.LightUp;
		} else if ("lightGrid" === val) {
			res = c_oAscPatternType.LightGrid;
		} else if ("lightTrellis" === val) {
			res = c_oAscPatternType.LightTrellis;
		} else if ("gray125" === val) {
			res = c_oAscPatternType.Gray125;
		} else if ("gray0625" === val) {
			res = c_oAscPatternType.Gray0625;
		}
		return res;
	}
	function ToXml_ST_PatternType(val) {
		switch (val) {
			case c_oAscPatternType.None:
				return 'none';
			case c_oAscPatternType.Solid:
				return 'solid';
			case c_oAscPatternType.MediumGray:
				return 'mediumGray';
			case c_oAscPatternType.DarkGray:
				return 'darkGray';
			case c_oAscPatternType.LightGray:
				return 'lightGray';
			case c_oAscPatternType.DarkHorizontal:
				return 'darkHorizontal';
			case c_oAscPatternType.DarkVertical:
				return 'darkVertical';
			case c_oAscPatternType.DarkDown:
				return 'darkDown';
			case c_oAscPatternType.DarkUp:
				return 'darkUp';
			case c_oAscPatternType.DarkGrid:
				return 'darkGrid';
			case c_oAscPatternType.DarkTrellis:
				return 'darkTrellis';
			case c_oAscPatternType.LightHorizontal:
				return 'lightHorizontal';
			case c_oAscPatternType.LightVertical:
				return 'lightVertical';
			case c_oAscPatternType.LightDown:
				return 'lightDown';
			case c_oAscPatternType.LightUp:
				return 'lightUp';
			case c_oAscPatternType.LightGrid:
				return 'lightGrid';
			case c_oAscPatternType.LightTrellis:
				return 'lightTrellis';
			case c_oAscPatternType.Gray125:
				return 'gray125';
			case c_oAscPatternType.Gray0625:
				return 'gray0625';
		}
	}

	function FromXml_ST_DataValidationOperator(val) {
		var res = null;
		switch (val) {
			case "between":
				res = Asc.EDataValidationOperator.Between;
				break;
			case "notBetween":
				res = Asc.EDataValidationOperator.NotBetween;
				break;
			case "equal":
				res = Asc.EDataValidationOperator.Equal;
				break;
			case "notEqual":
				res = Asc.EDataValidationOperator.NotEqual;
				break;
			case "lessThan":
				res = Asc.EDataValidationOperator.LessThan;
				break;
			case "lessThanOrEqual":
				res = Asc.EDataValidationOperator.LessThanOrEqual;
				break;
			case "greaterThan":
				res = Asc.EDataValidationOperator.GreaterThan;
				break;
			case "greaterThanOrEq":
				res = Asc.EDataValidationOperator.GreaterThanOrEqual;
		}
		return res;
	}

	function ToXml_ST_DataValidationOperator(val) {
		var res = null;
		switch (val) {
			case Asc.EDataValidationOperator.Between:
				res = "between";
				break;
			case Asc.EDataValidationOperator.NotBetween:
				res = "notBetween";
				break;
			case Asc.EDataValidationOperator.Equal:
				res = "equal";
				break;
			case Asc.EDataValidationOperator.NotEqual:
				res = "notEqual";
				break;
			case Asc.EDataValidationOperator.LessThan:
				res = "lessThan";
				break;
			case Asc.EDataValidationOperator.LessThanOrEqual:
				res = "lessThanOrEqual";
				break;
			case Asc.EDataValidationOperator.GreaterThan:
				res = "greaterThan";
				break;
			case Asc.EDataValidationOperator.GreaterThanOrEqual:
				res = "greaterThanOrEq";
		}
		return res;
	}

	function ToXml_CFOperatorType(nType) {
		var sType = undefined;
		switch (nType) {
			case AscCommonExcel.ECfOperator.Operator_beginsWith:
				sType = "beginsWith";
				break;
			case AscCommonExcel.ECfOperator.Operator_between:
				sType = "between";
				break;
			case AscCommonExcel.ECfOperator.Operator_containsText:
				sType = "containsText";
				break;
			case AscCommonExcel.ECfOperator.Operator_endsWith:
				sType = "endsWith";
				break;
			case AscCommonExcel.ECfOperator.Operator_equal:
				sType = "equal";
				break;
			case AscCommonExcel.ECfOperator.Operator_greaterThan:
				sType = "greaterThan";
				break;
			case AscCommonExcel.ECfOperator.Operator_greaterThanOrEqual:
				sType = "greaterThanOrEqual";
				break;
			case AscCommonExcel.ECfOperator.Operator_lessThan:
				sType = "lessThan";
				break;
			case AscCommonExcel.ECfOperator.Operator_lessThanOrEqual:
				sType = "lessThanOrEqual";
				break;
			case AscCommonExcel.ECfOperator.Operator_notBetween:
				sType = "notBetween";
				break;
			case AscCommonExcel.ECfOperator.Operator_notContains:
				sType = "notContains";
				break;
			case AscCommonExcel.ECfOperator.Operator_notEqual:
				sType = "notEqual";
				break;
		}

		return sType;
	}

	function FromXml_CFOperatorType(sType) {
		var nType = undefined;
		switch (sType) {
			case "beginsWith":
				nType = AscCommonExcel.ECfOperator.Operator_beginsWith;
				break;
			case "between":
				nType = AscCommonExcel.ECfOperator.Operator_between;
				break;
			case "containsText":
				nType = AscCommonExcel.ECfOperator.Operator_containsText;
				break;
			case "endsWith":
				nType = AscCommonExcel.ECfOperator.Operator_endsWith;
				break;
			case "equal":
				nType = AscCommonExcel.ECfOperator.Operator_equal;
				break;
			case "greaterThan":
				nType = AscCommonExcel.ECfOperator.Operator_greaterThan;
				break;
			case "greaterThanOrEqual":
				nType = AscCommonExcel.ECfOperator.Operator_greaterThanOrEqual;
				break;
			case "lessThan":
				nType = AscCommonExcel.ECfOperator.Operator_lessThan;
				break;
			case "lessThanOrEqual":
				nType = AscCommonExcel.ECfOperator.Operator_lessThanOrEqual;
				break;
			case "notBetween":
				nType = AscCommonExcel.ECfOperator.Operator_notBetween;
				break;
			case "notContains":
				nType = AscCommonExcel.ECfOperator.Operator_notContains;
				break;
			case "notEqual":
				nType = AscCommonExcel.ECfOperator.Operator_notEqual;
				break;
		}

		return nType;
	}

	function ToXml_ST_TimePeriod(nType) {
		var sType = undefined;
		switch (nType) {
			case AscCommonExcel.ST_TimePeriod.last7Days:
				sType = "last7Days";
				break;
			case AscCommonExcel.ST_TimePeriod.lastMonth:
				sType = "lastMonth";
				break;
			case AscCommonExcel.ST_TimePeriod.lastWeek:
				sType = "lastWeek";
				break;
			case AscCommonExcel.ST_TimePeriod.nextMonth:
				sType = "nextMonth";
				break;
			case AscCommonExcel.ST_TimePeriod.nextWeek:
				sType = "nextWeek";
				break;
			case AscCommonExcel.ST_TimePeriod.thisMonth:
				sType = "thisMonth";
				break;
			case AscCommonExcel.ST_TimePeriod.thisWeek:
				sType = "thisWeek";
				break;
			case AscCommonExcel.ST_TimePeriod.today:
				sType = "today";
				break;
			case AscCommonExcel.ST_TimePeriod.tomorrow:
				sType = "tomorrow";
				break;
			case AscCommonExcel.ST_TimePeriod.yesterday:
				sType = "yesterday";
				break;
		}

		return sType;
	}

	function FromXml_ST_TimePeriod(sType) {
		var nType = undefined;
		switch (sType) {
			case "last7Days":
				nType = AscCommonExcel.ST_TimePeriod.last7Days;
				break;
			case "lastMonth":
				nType = AscCommonExcel.ST_TimePeriod.lastMonth;
				break;
			case "lastWeek":
				nType = AscCommonExcel.ST_TimePeriod.lastWeek;
				break;
			case "nextMonth":
				nType = AscCommonExcel.ST_TimePeriod.nextMonth;
				break;
			case "nextWeek":
				nType = AscCommonExcel.ST_TimePeriod.nextWeek;
				break;
			case "thisMonth":
				nType = AscCommonExcel.ST_TimePeriod.thisMonth;
				break;
			case "thisWeek":
				nType = AscCommonExcel.ST_TimePeriod.thisWeek;
				break;
			case "today":
				nType = AscCommonExcel.ST_TimePeriod.today;
				break;
			case "tomorrow":
				nType = AscCommonExcel.ST_TimePeriod.tomorrow;
				break;
			case "yesterday":
				nType = AscCommonExcel.ST_TimePeriod.yesterday;
				break;
		}

		return nType;
	}

	function FromXml_ST_FilterOperator(val) {
		var res = -1;
		if ("equal" === val) {
			res = Asc.c_oAscCustomAutoFilter.equals;
		} else if ("lessThan" === val) {
			res = Asc.c_oAscCustomAutoFilter.isLessThan;
		} else if ("lessThanOrEqual" === val) {
			res = Asc.c_oAscCustomAutoFilter.isLessThanOrEqualTo;
		} else if ("notEqual" === val) {
			res = Asc.c_oAscCustomAutoFilter.doesNotEqual;
		} else if ("greaterThanOrEqual" === val) {
			res = Asc.c_oAscCustomAutoFilter.isGreaterThanOrEqualTo;
		} else if ("greaterThan" === val) {
			res = Asc.c_oAscCustomAutoFilter.isGreaterThan;
		}
		return res;
	}

	function ToXml_ST_FilterOperator(val) {
		var res = "";
		if (Asc.c_oAscCustomAutoFilter.equals === val) {
			res = "equal";
		} else if (Asc.c_oAscCustomAutoFilter.isLessThan === val) {
			res = "lessThan";
		} else if (Asc.c_oAscCustomAutoFilter.isLessThanOrEqualTo === val) {
			res = "lessThanOrEqual";
		} else if (Asc.c_oAscCustomAutoFilter.doesNotEqual === val) {
			res = "notEqual";
		} else if (Asc.c_oAscCustomAutoFilter.isGreaterThanOrEqualTo === val) {
			res = "greaterThanOrEqual";
		} else if (Asc.c_oAscCustomAutoFilter.isGreaterThan === val) {
			res = "greaterThan";
		}
		return res;
	}

	function FromXml_ST_DynamicFilterType(val) {
		var res = -1;
		if ("null" === val) {
			res = Asc.c_oAscDynamicAutoFilter.nullType;
		} else if ("aboveAverage" === val) {
			res = Asc.c_oAscDynamicAutoFilter.aboveAverage;
		} else if ("belowAverage" === val) {
			res = Asc.c_oAscDynamicAutoFilter.belowAverage;
		} else if ("tomorrow" === val) {
			res = Asc.c_oAscDynamicAutoFilter.tomorrow;
		} else if ("today" === val) {
			res = Asc.c_oAscDynamicAutoFilter.today;
		} else if ("yesterday" === val) {
			res = Asc.c_oAscDynamicAutoFilter.yesterday;
		} else if ("nextWeek" === val) {
			res = Asc.c_oAscDynamicAutoFilter.nextWeek;
		} else if ("thisWeek" === val) {
			res = Asc.c_oAscDynamicAutoFilter.thisWeek;
		} else if ("lastWeek" === val) {
			res = Asc.c_oAscDynamicAutoFilter.lastWeek;
		} else if ("nextMonth" === val) {
			res = Asc.c_oAscDynamicAutoFilter.nextMonth;
		} else if ("thisMonth" === val) {
			res = Asc.c_oAscDynamicAutoFilter.thisMonth;
		} else if ("lastMonth" === val) {
			res = Asc.c_oAscDynamicAutoFilter.lastMonth;
		} else if ("nextQuarter" === val) {
			res = Asc.c_oAscDynamicAutoFilter.nextQuarter;
		} else if ("thisQuarter" === val) {
			res = Asc.c_oAscDynamicAutoFilter.thisQuarter;
		} else if ("lastQuarter" === val) {
			res = Asc.c_oAscDynamicAutoFilter.lastQuarter;
		} else if ("nextYear" === val) {
			res = Asc.c_oAscDynamicAutoFilter.nextYear;
		} else if ("thisYear" === val) {
			res = Asc.c_oAscDynamicAutoFilter.thisYear;
		} else if ("lastYear" === val) {
			res = Asc.c_oAscDynamicAutoFilter.lastYear;
		} else if ("yearToDate" === val) {
			res = Asc.c_oAscDynamicAutoFilter.yearToDate;
		} else if ("Q1" === val) {
			res = Asc.c_oAscDynamicAutoFilter.q1;
		} else if ("Q2" === val) {
			res = Asc.c_oAscDynamicAutoFilter.q2;
		} else if ("Q3" === val) {
			res = Asc.c_oAscDynamicAutoFilter.q3;
		} else if ("Q4" === val) {
			res = Asc.c_oAscDynamicAutoFilter.q4;
		} else if ("M1" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m1;
		} else if ("M2" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m2;
		} else if ("M3" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m3;
		} else if ("M4" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m4;
		} else if ("M5" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m5;
		} else if ("M6" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m6;
		} else if ("M7" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m7;
		} else if ("M8" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m8;
		} else if ("M9" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m9;
		} else if ("M10" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m10;
		} else if ("M11" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m11;
		} else if ("M12" === val) {
			res = Asc.c_oAscDynamicAutoFilter.m12;
		}
		return res;
	}

	function ToXml_ST_DynamicFilterType(val) {
		var res = "";
		if (Asc.c_oAscDynamicAutoFilter.nullType === val) {
			res = "null";
		} else if (Asc.c_oAscDynamicAutoFilter.aboveAverage === val) {
			res = "aboveAverage";
		} else if (Asc.c_oAscDynamicAutoFilter.belowAverage === val) {
			res = "belowAverage";
		} else if (Asc.c_oAscDynamicAutoFilter.tomorrow === val) {
			res = "tomorrow";
		} else if (Asc.c_oAscDynamicAutoFilter.today === val) {
			res = "today";
		} else if (Asc.c_oAscDynamicAutoFilter.yesterday === val) {
			res = "yesterday";
		} else if (Asc.c_oAscDynamicAutoFilter.nextWeek === val) {
			res = "nextWeek";
		} else if (Asc.c_oAscDynamicAutoFilter.thisWeek === val) {
			res = "thisWeek";
		} else if (Asc.c_oAscDynamicAutoFilter.lastWeek === val) {
			res = "lastWeek";
		} else if (Asc.c_oAscDynamicAutoFilter.nextMonth === val) {
			res = "nextMonth";
		} else if (Asc.c_oAscDynamicAutoFilter.thisMonth === val) {
			res = "thisMonth";
		} else if (Asc.c_oAscDynamicAutoFilter.lastMonth === val) {
			res = "lastMonth";
		} else if (Asc.c_oAscDynamicAutoFilter.nextQuarter === val) {
			res = "nextQuarter";
		} else if (Asc.c_oAscDynamicAutoFilter.thisQuarter === val) {
			res = "thisQuarter";
		} else if (Asc.c_oAscDynamicAutoFilter.lastQuarter === val) {
			res = "lastQuarter";
		} else if (Asc.c_oAscDynamicAutoFilter.nextYear === val) {
			res = "nextYear";
		} else if (Asc.c_oAscDynamicAutoFilter.thisYear === val) {
			res = "thisYear";
		} else if (Asc.c_oAscDynamicAutoFilter.lastYear === val) {
			res = "lastYear";
		} else if (Asc.c_oAscDynamicAutoFilter.yearToDate === val) {
			res = "yearToDate";
		} else if (Asc.c_oAscDynamicAutoFilter.q1 === val) {
			res = "Q1";
		} else if (Asc.c_oAscDynamicAutoFilter.q2 === val) {
			res = "Q2";
		} else if (Asc.c_oAscDynamicAutoFilter.q3 === val) {
			res = "Q3";
		} else if (Asc.c_oAscDynamicAutoFilter.q4 === val) {
			res = "Q4";
		} else if (Asc.c_oAscDynamicAutoFilter.m1 === val) {
			res = "M1";
		} else if (Asc.c_oAscDynamicAutoFilter.m2 === val) {
			res = "M2";
		} else if (Asc.c_oAscDynamicAutoFilter.m3 === val) {
			res = "M3";
		} else if (Asc.c_oAscDynamicAutoFilter.m4 === val) {
			res = "M4";
		} else if (Asc.c_oAscDynamicAutoFilter.m5 === val) {
			res = "M5";
		} else if (Asc.c_oAscDynamicAutoFilter.m6 === val) {
			res = "M6";
		} else if (Asc.c_oAscDynamicAutoFilter.m7 === val) {
			res = "M7";
		} else if (Asc.c_oAscDynamicAutoFilter.m8 === val) {
			res = "M8";
		} else if (Asc.c_oAscDynamicAutoFilter.m9 === val) {
			res = "M9";
		} else if (Asc.c_oAscDynamicAutoFilter.m10 === val) {
			res = "M10";
		} else if (Asc.c_oAscDynamicAutoFilter.m11 === val) {
			res = "M11";
		} else if (Asc.c_oAscDynamicAutoFilter.m12 === val) {
			res = "M12";
		}
		return res;
	}

	function FromXml_ST_DateTimeGrouping(val) {
		var res = -1;
		if ("year" === val) {
			res = Asc.EDateTimeGroup.datetimegroupYear;
		} else if ("month" === val) {
			res = Asc.EDateTimeGroup.datetimegroupMonth;
		} else if ("day" === val) {
			res = Asc.EDateTimeGroup.datetimegroupDay;
		} else if ("hour" === val) {
			res = Asc.EDateTimeGroup.datetimegroupHour;
		} else if ("minute" === val) {
			res = Asc.EDateTimeGroup.datetimegroupMinute;
		} else if ("second" === val) {
			res = Asc.EDateTimeGroup.datetimegroupSecond;
		}
		return res;
	}

	function ToXml_ST_DateTimeGrouping(val) {
		var res = "";
		if (Asc.EDateTimeGroup.datetimegroupYear === val) {
			res = "year";
		} else if (Asc.EDateTimeGroup.datetimegroupMonth === val) {
			res = "month";
		} else if (Asc.EDateTimeGroup.datetimegroupDay === val) {
			res = "day";
		} else if (Asc.EDateTimeGroup.datetimegroupHour === val) {
			res = "hour";
		} else if (Asc.EDateTimeGroup.datetimegroupMinute === val) {
			res = "minute";
		} else if (Asc.EDateTimeGroup.datetimegroupSecond === val) {
			res = "second";
		}
		return res;
	}

	function ToXml_ST_HorizontalAlignment(val, default_null) {
		var res = default_null ? null : -1;
		switch (val) {
			case -1:
				res = "general";
				break;
			case AscCommon.align_Left:
				res = "left";
				break;
			case AscCommon.align_Center:
				res = "center";
				break;
			case AscCommon.align_Right:
				res = "right";
				break;
			case AscCommon.align_Justify:
				res = "justify";
				break;
		}
		return res;
	}

	function FromXml_ST_HorizontalAlignment(val, default_null) {
		var res = default_null ? null: -1;//general == null
		if ("left" === val) {
			res = AscCommon.align_Left;
		} else if ("center" === val) {
			res = AscCommon.align_Center;
		} else if ("right" === val) {
			res = AscCommon.align_Right;
		} else if ("fill" === val) {
			res = AscCommon.align_Justify;
		} else if ("justify" === val) {
			res = AscCommon.align_Justify;
		} else if ("centerContinuous" === val) {
			res = AscCommon.align_Center;
		} else if ("distributed" === val) {
			res = AscCommon.align_Justify;
		}
		return res;
	}

	function ToXml_ST_VerticalAlignment(val, default_null) {
		var res = default_null ? null : -1;
		switch (val) {
			case Asc.c_oAscVAlign.Top:
				res = "top";
				break;
			case Asc.c_oAscVAlign.Center:
				res = "center";
				break;
			case Asc.c_oAscVAlign.Bottom:
				res = "bottom";
				break;
			case Asc.c_oAscVAlign.Just:
				res = "justify";
				break;
			case Asc.c_oAscVAlign.Dist:
				res = "distributed";
				break;
		}
		return res;
	}

	function FromXml_ST_VerticalAlignment(val) {
		var res = -1;
		if ("top" === val) {
			res = Asc.c_oAscVAlign.Top;
		} else if ("center" === val) {
			res = Asc.c_oAscVAlign.Center;
		} else if ("bottom" === val) {
			res = Asc.c_oAscVAlign.Bottom;
		} else if ("justify" === val) {
			res = Asc.c_oAscVAlign.Just;
		} else if ("distributed" === val) {
			res = Asc.c_oAscVAlign.Dist;
		}
		return res;
	}

	function ToXml_ST_CfvoType(nType) {
		var sType = "";
		switch (nType) {
			case AscCommonExcel.ECfvoType.Formula:
				sType = "formula";
				break;
			case AscCommonExcel.ECfvoType.Maximum:
				sType = "max";
				break;
			case AscCommonExcel.ECfvoType.Minimum:
				sType = "min";
				break;
			case AscCommonExcel.ECfvoType.Number:
				sType = "num";
				break;
			case AscCommonExcel.ECfvoType.Percent:
				sType = "percent";
				break;
			case AscCommonExcel.ECfvoType.Percentile:
				sType = "percentile";
				break;
			case AscCommonExcel.ECfvoType.AutoMin:
				sType = "autoMin";
				break;
			case AscCommonExcel.ECfvoType.AutoMax:
				sType = "autoMax";
				break;
		}

		return sType;
	}

	function FromXml_ST_CfvoType(sType) {
		var nType = -1;
		switch (sType) {
			case "formula":
				nType = AscCommonExcel.ECfvoType.Formula;
				break;
			case "max":
				nType = AscCommonExcel.ECfvoType.Maximum;
				break;
			case "min":
				nType = AscCommonExcel.ECfvoType.Minimum;
				break;
			case "num":
				nType = AscCommonExcel.ECfvoType.Number;
				break;
			case "percent":
				nType = AscCommonExcel.ECfvoType.Percent;
				break;
			case "percentile":
				nType = AscCommonExcel.ECfvoType.Percentile;
				break;
			case "autoMin":
				nType = AscCommonExcel.ECfvoType.AutoMin;
				break;
			case "autoMax":
				nType = AscCommonExcel.ECfvoType.AutoMax;
				break;
		}

		return nType;
	}

	function FromXml_ST_IconSetType(val) {
		//в пивотах есть функция FromXml_ST_IconSetType, но там корвенртирцем в другие константы. пока оставляю так, нужно сделать общие
		var res = undefined;
		switch (val) {
			case "3Arrows":
				res = Asc.EIconSetType.Arrows3;
				break;
			case "3ArrowsGray":
				res = Asc.EIconSetType.Arrows3Gray;
				break;
			case "3Flags":
				res = Asc.EIconSetType.Flags3;
				break;
			case "3TrafficLights1":
				res = Asc.EIconSetType.Traffic3Lights1;
				break;
			case "3TrafficLights2":
				res = Asc.EIconSetType.Traffic3Lights2;
				break;
			case "3Signs":
				res = Asc.EIconSetType.Signs3;
				break;
			case "3Symbols":
				res = Asc.EIconSetType.Symbols3;
				break;
			case "3Symbols2":
				res = Asc.EIconSetType.Symbols3_2;
				break;
			case "4Arrows":
				res = Asc.EIconSetType.Arrows4;
				break;
			case "4ArrowsGray":
				res = Asc.EIconSetType.Arrows4Gray;
				break;
			case "4TrafficLights":
				res = Asc.EIconSetType.Traffic4Lights;
				break;
			case "5Arrows":
				res = Asc.EIconSetType.Arrows5;
				break;
			case "5ArrowsGray":
				res = Asc.EIconSetType.Arrows5Gray;
				break;
			case "5Rating":
				res = Asc.EIconSetType.Rating5;
				break;
			case "5Quarters":
				res = Asc.EIconSetType.Quarters5;
				break;
			case "4RedToBlack":
				res = Asc.EIconSetType.RedToBlack4;
				break;
			case "4Rating":
				res = Asc.EIconSetType.Rating4;
				break;
			case "3Stars":
				res = Asc.EIconSetType.Stars3;
				break;
			case "3Triangles":
				res = Asc.EIconSetType.Triangles3;
				break;
			case "5Boxes":
				res = Asc.EIconSetType.Boxes5;
				break;
			case "NoIcons":
				res = Asc.EIconSetType.NoIcons;
				break;
		}
		return res;
	}

	function ToXml_ST_IconSetType(val) {
		//в пивотах есть функция ToXml_ST_IconSetType, но там корвенртирцем в другие константы. пока оставляю так, нужно сделать общие
		var res = undefined;
		switch (val) {
			case Asc.EIconSetType.Arrows3:
				res = "3Arrows";
				break;
			case Asc.EIconSetType.Arrows3Gray:
				res = "3ArrowsGray";
				break;
			case Asc.EIconSetType.Flags3:
				res = "3Flags";
				break;
			case Asc.EIconSetType.Traffic3Lights1:
				res = "3TrafficLights1";
				break;
			case Asc.EIconSetType.Traffic3Lights2:
				res = "3TrafficLights2";
				break;
			case Asc.EIconSetType.Signs3:
				res = "3Signs";
				break;
			case Asc.EIconSetType.Symbols3:
				res = "3Symbols";
				break;
			case Asc.EIconSetType.Symbols3_2:
				res = "3Symbols2";
				break;
			case Asc.EIconSetType.Arrows4:
				res = "4Arrows";
				break;
			case Asc.EIconSetType.Arrows4Gray:
				res = "4ArrowsGray";
				break;
			case Asc.EIconSetType.Traffic4Lights:
				res = "4TrafficLights";
				break;
			case Asc.EIconSetType.Arrows5:
				res = "5Arrows";
				break;
			case Asc.EIconSetType.Arrows5Gray:
				res = "5ArrowsGray";
				break;
			case Asc.EIconSetType.Rating5:
				res = "5Rating";
				break;
			case Asc.EIconSetType.Quarters5:
				res = "5Quarters";
				break;
			case Asc.EIconSetType.RedToBlack4:
				res = "4RedToBlack";
				break;
			case Asc.EIconSetType.Rating4:
				res = "4Rating";
				break;
			case Asc.EIconSetType.Stars3:
				res = "3Stars";
				break;
			case Asc.EIconSetType.Triangles3:
				res = "3Triangles";
				break;
			case Asc.EIconSetType.Boxes5:
				res = "5Boxes";
				break;
			case Asc.EIconSetType.NoIcons:
				res = "NoIcons";
				break;
		}
		return res;
	}

	function FromXml_ST_CfType(val) {
		var res = null;
		switch (val) {
			case "aboveAverage":
				res = Asc.ECfType.aboveAverage;
				break;
			case "beginsWith":
				res = Asc.ECfType.beginsWith;
				break;
			case "cellIs":
				res = Asc.ECfType.cellIs;
				break;
			case "colorScale":
				res = Asc.ECfType.colorScale;
				break;
			case "containsBlanks":
				res = Asc.ECfType.containsBlanks;
				break;
			case "containsErrors":
				res = Asc.ECfType.containsErrors;
				break;
			case "containsText":
				res = Asc.ECfType.containsText;
				break;
			case "dataBar":
				res = Asc.ECfType.dataBar;
				break;
			case "duplicateValues":
				res = Asc.ECfType.duplicateValues;
				break;
			case "expression":
				res = Asc.ECfType.expression;
				break;
			case "notContainsBlanks":
				res = Asc.ECfType.notContainsBlanks;
				break;
			case "notContainsErrors":
				res = Asc.ECfType.notContainsErrors;
				break;
			case "notContainsText":
				res = Asc.ECfType.notContainsText;
				break;
			case "timePeriod":
				res = Asc.ECfType.timePeriod;
				break;
			case "top10":
				res = Asc.ECfType.top10;
				break;
			case "uniqueValues":
				res = Asc.ECfType.uniqueValues;
				break;
			case "endsWith":
				res = Asc.ECfType.endsWith;
				break;
			case "iconSet":
				res = Asc.ECfType.iconSet;
				break;
		}
		return res;
	}

	function ToXml_ST_CfType(val) {
		var res = null;
		switch (val) {
			case Asc.ECfType.aboveAverage:
				res = "aboveAverage";
				break;
			case Asc.ECfType.beginsWith:
				res = "beginsWith";
				break;
			case Asc.ECfType.cellIs:
				res = "cellIs";
				break;
			case Asc.ECfType.colorScale:
				res = "colorScale";
				break;
			case Asc.ECfType.containsBlanks:
				res = "containsBlanks";
				break;
			case Asc.ECfType.containsErrors:
				res = "containsErrors";
				break;
			case Asc.ECfType.containsText:
				res = "containsText";
				break;
			case Asc.ECfType.dataBar:
				res = "dataBar";
				break;
			case Asc.ECfType.duplicateValues:
				res = "duplicateValues";
				break;
			case Asc.ECfType.expression:
				res = "expression";
				break;
			case Asc.ECfType.notContainsBlanks:
				res = "notContainsBlanks";
				break;
			case Asc.ECfType.notContainsErrors:
				res = "notContainsErrors";
				break;
			case Asc.ECfType.notContainsText:
				res = "notContainsText";
				break;
			case Asc.ECfType.timePeriod:
				res = "timePeriod";
				break;
			case Asc.ECfType.top10:
				res = "top10";
				break;
			case Asc.ECfType.uniqueValues:
				res = "uniqueValues";
				break;
			case Asc.ECfType.endsWith:
				res = "endsWith";
				break;
			case Asc.ECfType.iconSet:
				res = "iconSet";
				break;
		}
		return res;
	}

	function ToXml_ST_DataValidationErrorStyle(nType) {
		var sType = undefined;
		switch (nType) {
			case Asc.EDataValidationErrorStyle.Stop:
				sType = "stop";
				break;
			case Asc.EDataValidationErrorStyle.Warning:
				sType = "warning";
				break;
			case Asc.EDataValidationErrorStyle.Information:
				sType = "information";
				break;
		}

		return sType;
	}

	function FromXml_ST_DataValidationErrorStyle(sType) {
		var nType = undefined;
		switch (sType) {
			case "stop":
				nType = Asc.EDataValidationErrorStyle.Stop;
				break;
			case "warning":
				nType = Asc.EDataValidationErrorStyle.Warning;
				break;
			case "information":
				nType = Asc.EDataValidationErrorStyle.Information;
				break;
		}

		return nType;
	}

	function FromXml_ST_DataValidationType(val) {
		var res = undefined;
		switch (val) {
			case "none":
				res = Asc.EDataValidationType.None;
				break;
			case "whole":
				res = Asc.EDataValidationType.Whole;
				break;
			case "decimal":
				res = Asc.EDataValidationType.Decimal;
				break;
			case "list":
				res = Asc.EDataValidationType.List;
				break;
			case "date":
				res = Asc.EDataValidationType.Date;
				break;
			case "time":
				res = Asc.EDataValidationType.Time;
				break;
			case "textLength":
				res = Asc.EDataValidationType.TextLength;
				break;
			case "custom":
				res = Asc.EDataValidationType.Custom;
				break;
		}
		return res;
	}

	function ToXml_ST_DataValidationType(val) {
		var res = undefined;
		switch (val) {
			case Asc.EDataValidationType.None:
				res = "none";
				break;
			case Asc.EDataValidationType.Whole:
				res = "whole";
				break;
			case Asc.EDataValidationType.Decimal:
				res = "decimal";
				break;
			case Asc.EDataValidationType.List:
				res = "list";
				break;
			case Asc.EDataValidationType.Date:
				res = "date";
				break;
			case Asc.EDataValidationType.Time:
				res = "time";
				break;
			case Asc.EDataValidationType.TextLength:
				res = "textLength";
				break;
			case Asc.EDataValidationType.Custom:
				res = "custom";
				break;
		}
		return res;
	}

	function FromXml_ST_DataValidationImeMode(val) {
		var res = undefined;
		switch (val) {
			case "noControl":
				res = Asc.EDataValidationImeMode.NoControl;
				break;
			case "off":
				res = Asc.EDataValidationImeMode.Off;
				break;
			case "on":
				res = Asc.EDataValidationImeMode.On;
				break;
			case "disabled":
				res = Asc.EDataValidationImeMode.Disabled;
				break;
			case "hiragana":
				res = Asc.EDataValidationImeMode.Hiragana;
				break;
			case "fullKatakana":
				res = Asc.EDataValidationImeMode.FullKatakana;
				break;
			case "halfKatakana":
				res = Asc.EDataValidationImeMode.HalfKatakana;
				break;
			case "fullAlpha":
				res = Asc.EDataValidationImeMode.FullAlpha;
				break;
			case "halfAlpha":
				res = Asc.EDataValidationImeMode.HalfAlpha;
				break;
			case "fullHangul":
				res = Asc.EDataValidationImeMode.FullHangul;
				break;
			case "halfHangul":
				res = Asc.EDataValidationImeMode.HalfHangul;
				break;
		}
		return res;
	}

	function ToXml_ST_DataValidationImeMode(val) {
		var res = undefined;
		switch (val) {
			case Asc.EDataValidationImeMode.NoControl:
				res = "noControl";
				break;
			case Asc.EDataValidationImeMode.Off:
				res = "off";
				break;
			case Asc.EDataValidationImeMode.On:
				res = "on";
				break;
			case Asc.EDataValidationImeMode.Disabled:
				res = "disabled";
				break;
			case Asc.EDataValidationImeMode.Hiragana:
				res = "hiragana";
				break;
			case Asc.EDataValidationImeMode.FullKatakana:
				res = "fullKatakana";
				break;
			case Asc.EDataValidationImeMode.HalfKatakana:
				res = "halfKatakana";
				break;
			case Asc.EDataValidationImeMode.FullAlpha:
				res = "fullAlpha";
				break;
			case Asc.EDataValidationImeMode.HalfAlpha:
				res = "halfAlpha";
				break;
			case Asc.EDataValidationImeMode.FullHangul:
				res = "fullHangul";
				break;
			case Asc.EDataValidationImeMode.HalfHangul:
				res = "halfHangul";
				break;
		}
		return res;
	}

	function FromXml_ST_GradientType(val, default_null) {
		var res = default_null ? null : -1;
		if ("linear" === val) {
			res = Asc.c_oAscFillGradType.GRAD_LINEAR;
		} else if ("path" === val) {
			res = Asc.c_oAscFillGradType.GRAD_PATH;
		}
		return res;
	}

	function ToXml_ST_GradientType(val, default_null) {
		var res = default_null ? null : -1;
		if (Asc.c_oAscFillGradType.GRAD_LINEAR === val) {
			res = "linear";
		} else if (Asc.c_oAscFillGradType.GRAD_PATH === val) {
			res = "path";
		}
		return res;
	}

	function FromXml_ST_FontScheme(val) {
		var res = null;
		switch (val) {
			case  "major":
				res = Asc.EFontScheme.fontschemeMajor;
				break;
			case  "minor":
				res = Asc.EFontScheme.fontschemeMinor;
				break;
			case  Asc.EFontScheme.fontschemeNone:
				break;
		}
		return res;
	}

	function ToXml_ST_FontScheme(val) {
		var res = null;
		switch (val) {
			case  Asc.EFontScheme.fontschemeMajor:
				res = "major";
				break;
			case  Asc.EFontScheme.fontschemeMinor:
				res = "minor";
				break;
			case  Asc.EFontScheme.fontschemeNone:
				break;
		}
		return res;
	}

	function FromXml_ST_PageOrder(val) {
		var res = null;
		switch (val) {
			case  "downThenOver":
				res = 0;
				break;
			case  "overThenDown":
				res = 1;
				break;
		}
		return res;
	}

	function ToXml_ST_PageOrder(val) {
		var res = null;
		switch (val) {
			case  0:
				res = "downThenOver";
				break;
			case  1:
				res = "overThenDown";
				break;
		}
		return res;
	}

	function FromXml_ST_PageOrientation(val) {
		var res = null;
		switch (val) {
			case "portrait":
				res = Asc.c_oAscPageOrientation.PagePortrait;
				break;
			case "landscape":
				res = Asc.c_oAscPageOrientation.PageLandscape;
				break;
		}
		return res;
	}

	function ToXml_ST_PageOrientation(val) {
		var res = null;
		switch (val) {
			case Asc.c_oAscPageOrientation.PagePortrait:
				res = "portrait";
				break;
			case Asc.c_oAscPageOrientation.PageLandscape:
				res = "landscape";
				break;
		}
		return res;
	}

	function FromXml_ST_CellComments(val) {
		var res = 0;
		switch (val) {
			case "none":
				res = AscCommonExcel.ST_CellComments.none;
				break;
			case "asDisplayed":
				res = AscCommonExcel.ST_CellComments.asDisplayed;
				break;
			case "atEnd":
				res = AscCommonExcel.ST_CellComments.atEnd;
				break;
		}
		return res;
	}

	function ToXml_ST_CellComments(val) {
		var res = null;
		switch (val) {
			case AscCommonExcel.ST_CellComments.none:
				res = "none";
				break;
			case AscCommonExcel.ST_CellComments.asDisplayed:
				res = "asDisplayed";
				break;
			case AscCommonExcel.ST_CellComments.atEnd:
				res = "atEnd";
				break;
		}
		return res;
	}

	function FromXml_ST_PrintError(val) {
		var res = 0;
		switch (val) {
			case "displayed":
				res = AscCommonExcel.ST_PrintError.displayed;
				break;
			case "blank":
				res = AscCommonExcel.ST_PrintError.blank;
				break;
			case "dash":
				res = AscCommonExcel.ST_PrintError.dash;
				break;
			case "NA":
				res = AscCommonExcel.ST_PrintError.NA;
				break;
		}
		return res;
	}

	function ToXml_ST_PrintError(val) {
		var res = null;
		switch (val) {
			case AscCommonExcel.ST_PrintError.displayed:
				res = "displayed";
				break;
			case AscCommonExcel.ST_PrintError.blank:
				res = "blank";
				break;
			case AscCommonExcel.ST_PrintError.dash:
				res = "dash";
				break;
			case AscCommonExcel.ST_PrintError.NA:
				res = "NA";
				break;
		}
		return res;
	}

	function FromXml_ST_SortBy(val) {
		var res = 0;
		switch (val) {
			case "value":
				res = Asc.ESortBy.sortbyValue;
				break;
			case "cellColor":
				res = Asc.ESortBy.sortbyCellColor;
				break;
			case "fontColor":
				res = Asc.ESortBy.sortbyFontColor;
				break;
			case "icon":
				res = Asc.ESortBy.sortbyIcon;
				break;
		}
		return res;
	}

	function ToXml_ST_SortBy(val) {
		var res = null;
		switch (val) {
			case Asc.ESortBy.sortbyValue:
				res = "value";
				break;
			case Asc.ESortBy.sortbyCellColor:
				res = "cellColor";
				break;
			case Asc.ESortBy.sortbyFontColor:
				res = "fontColor";
				break;
			case Asc.ESortBy.sortbyIcon:
				res = "icon";
				break;
		}
		return res;
	}

	function FromXml_ST_TabularSlicerCacheSortOrder(val) {
		var res = null;
		switch (val) {
			case "ascending":
				res = Asc.ST_tabularSlicerCacheSortOrder.Ascending;
				break;
			case "descending":
				res = Asc.ST_tabularSlicerCacheSortOrder.Descending;
				break;
		}
		return res;
	}

	function ToXml_ST_TabularSlicerCacheSortOrder(val) {
		var res = null;
		switch (val) {
			case Asc.ST_tabularSlicerCacheSortOrder.Ascending:
				res = "ascending";
				break;
			case Asc.ST_tabularSlicerCacheSortOrder.Descending:
				res = "descending";
				break;
		}
		return res;
	}

	function FromXml_ST_SlicerCacheCrossFilter(val) {
		var res = null;
		switch (val) {
			case "showItemsWithDataAtTop":
				res = Asc.ST_slicerCacheCrossFilter.ShowItemsWithDataAtTop;
				break;
			case "showItemsWithNoData":
				res = Asc.ST_slicerCacheCrossFilter.ShowItemsWithNoData;
				break;
			default:
				res = Asc.ST_slicerCacheCrossFilter.None;
				break;
		}
		return res;
	}

	function ToXml_ST_SlicerCacheCrossFilter(val) {
		var res = null;
		switch (val) {
			case Asc.ST_slicerCacheCrossFilter.ShowItemsWithDataAtTop:
				res = "showItemsWithDataAtTop";
				break;
			case Asc.ST_slicerCacheCrossFilter.ShowItemsWithNoData:
				res = "showItemsWithNoData";
				break;
			default:
				res = "none";
				break;
		}
		return res;
	}

	function FromXml_ST_SortMethod(val) {
		var res = null;
		switch (val) {
			case "stroke":
				res = AscCommonExcel.ESortMethod.sortmethodStroke;
				break;
			case "pinYin":
				res = AscCommonExcel.ESortMethod.sortmethodPinYin;
				break;
		}
		return res;
	}

	function ToXml_ST_SortMethod(val) {
		var res = null;
		switch (val) {
			case AscCommonExcel.ESortMethod.sortmethodStroke:
				res = "stroke";
				break;
			case AscCommonExcel.ESortMethod.sortmethodPinYin:
				res = "pinYin";
				break;
		}
		return res;
	}

	function FromXML_ST_DispBlanksAs(val) {
		var res = null;
		switch (val) {
			case "span":
				res = Asc.c_oAscEDispBlanksAs.Span;
				break;
			case "gap":
				res = Asc.c_oAscEDispBlanksAs.Gap;
				break;
			case "zero":
				res = Asc.c_oAscEDispBlanksAs.Zero;
				break;
		}
		return res;
	}

	function ToXML_ST_DispBlanksAs(val) {
		var res = null;
		switch (val) {
			case Asc.c_oAscEDispBlanksAs.Span:
				res = "span";
				break;
			case Asc.c_oAscEDispBlanksAs.Gap:
				res = "gap";
				break;
			case Asc.c_oAscEDispBlanksAs.Zero:
				res = "zero";
				break;
		}
		return res;
	}

	function FromXml_ST_SparklineType(val) {
		var res = null;
		switch (val) {
			case "line":
				res = Asc.c_oAscSparklineType.Line;
				break;
			case "column":
				res = Asc.c_oAscSparklineType.Column;
				break;
			case "stacked":
				res = Asc.c_oAscSparklineType.Stacked;
				break;
		}
		return res;
	}

	function ToXml_ST_SparklineType(val) {
		var res = null;
		switch (val) {
			case Asc.c_oAscSparklineType.Line:
				res = "line";
				break;
			case Asc.c_oAscSparklineType.Column:
				res = "column";
				break;
			case Asc.c_oAscSparklineType.Stacked:
				res = "stacked";
				break;
		}
		return res;
	}

	function FromXml_ST_UnderlineValues(val) {
		var res = null;
		switch (val) {
			case "single":
				res = Asc.EUnderline.underlineSingle;
				break;
			case "double":
				res = Asc.EUnderline.underlineDouble;
				break;
			case "singleAccounting":
				res = Asc.EUnderline.underlineSingleAccounting;
				break;
			case "doubleAccounting":
				res = Asc.EUnderline.underlineDoubleAccounting;
				break;
			case "none":
				res = Asc.EUnderline.underlineNone;
				break;
		}
		return res;
	}

	function ToXml_ST_UnderlineValues(val) {
		var res = null;
		switch (val) {
			case Asc.EUnderline.underlineSingle:
				res = "single";
				break;
			case Asc.EUnderline.underlineDouble:
				res = "double";
				break;
			case Asc.EUnderline.underlineSingleAccounting:
				res = "singleAccounting";
				break;
			case Asc.EUnderline.underlineDoubleAccounting:
				res = "doubleAccounting";
				break;
			case Asc.EUnderline.underlineNone:
				res = "none";
				break;
		}
		return res;
	}

	function FromXml_ST_SparklineAxisMinMax(val) {
		var res = null;
		switch (val) {
			case "individual":
				res = Asc.c_oAscSparklineAxisMinMax.Individual;
				break;
			case "group":
				res = Asc.c_oAscSparklineAxisMinMax.Group;
				break;
			case "custom":
				res = Asc.c_oAscSparklineAxisMinMax.Custom;
				break;
		}
		return res;
	}

	function ToXml_ST_SparklineAxisMinMax(val) {
		var res = null;
		switch (val) {
			case Asc.c_oAscSparklineAxisMinMax.Individual:
				res = "individual";
				break;
			case Asc.c_oAscSparklineAxisMinMax.Group:
				res = "group";
				break;
			case Asc.c_oAscSparklineAxisMinMax.Custom:
				res = "custom";
				break;
		}
		return res;
	}

	function FromXml_ST_CellFormulaType(val) {
		var res = null;
		switch (val) {
			case "array":
				res = window["Asc"].ECellFormulaType.cellformulatypeArray;
				break;
			case "shared":
				res = window["Asc"].ECellFormulaType.cellformulatypeShared;
				break;
			case "dataTable":
				res = window["Asc"].ECellFormulaType.cellformulatypeDataTable;
				break;
		}
		return res;
	}

	function ToXml_ST_CellFormulaType(val) {
		var res = null;
		switch (val) {
			case window["Asc"].ECellFormulaType.cellformulatypeArray:
				res = "array";
				break;
			case window["Asc"].ECellFormulaType.cellformulatypeShared:
				res = "shared";
				break;
			case window["Asc"].ECellFormulaType.cellformulatypeDataTable:
				res = "dataTable";
				break;
		}
		return res;
	}

	function FromXml_ST_TableStyleType(val) {
		var res = null;
		switch (val) {
			case "wholeTable":
				res = Asc.ETableStyleType.tablestyletypeWholeTable;
				break;
			case "headerRow":
				res = Asc.ETableStyleType.tablestyletypeHeaderRow;
				break;
			case "totalRow":
				res = Asc.ETableStyleType.tablestyletypeTotalRow;
				break;
			case "firstColumn":
				res = Asc.ETableStyleType.tablestyletypeFirstColumn;
				break;
			case "lastColumn":
				res = Asc.ETableStyleType.tablestyletypeLastColumn;
				break;
			case "firstRowStripe":
				res = Asc.ETableStyleType.tablestyletypeFirstRowStripe;
				break;
			case "secondRowStripe":
				res = Asc.ETableStyleType.tablestyletypeSecondRowStripe;
				break;
			case "firstColumnStripe":
				res = Asc.ETableStyleType.tablestyletypeFirstColumnStripe;
				break;
			case "secondColumnStripe":
				res = Asc.ETableStyleType.tablestyletypeSecondColumnStripe;
				break;
			case "firstHeaderCell":
				res = Asc.ETableStyleType.tablestyletypeFirstHeaderCell;
				break;
			case "lastHeaderCell":
				res = Asc.ETableStyleType.tablestyletypeLastHeaderCell;
				break;
			case "firstTotalCell":
				res = Asc.ETableStyleType.tablestyletypeFirstTotalCell;
				break;
			case "lastTotalCell":
				res = Asc.ETableStyleType.tablestyletypeLastTotalCell;
				break;
			case "firstSubtotalColumn":
				res = Asc.ETableStyleType.tablestyletypeFirstSubtotalColumn;
				break;
			case "secondSubtotalColumn":
				res = Asc.ETableStyleType.tablestyletypeSecondSubtotalColumn;
				break;
			case "thirdSubtotalColumn":
				res = Asc.ETableStyleType.tablestyletypeThirdSubtotalColumn;
				break;
			case "firstSubtotalRow":
				res = Asc.ETableStyleType.tablestyletypeFirstSubtotalRow;
				break;
			case "secondSubtotalRow":
				res = Asc.ETableStyleType.tablestyletypeSecondSubtotalRow;
				break;
			case "thirdSubtotalRow":
				res = Asc.ETableStyleType.tablestyletypeThirdSubtotalRow;
				break;
			case "blankRow":
				res = Asc.ETableStyleType.tablestyletypeBlankRow;
				break;
			case "firstColumnSubheading":
				res = Asc.ETableStyleType.tablestyletypeFirstColumnSubheading;
				break;
			case "secondColumnSubheading":
				res = Asc.ETableStyleType.tablestyletypeSecondColumnSubheading;
				break;
			case "thirdColumnSubheading":
				res = Asc.ETableStyleType.tablestyletypeThirdColumnSubheading;
				break;
			case "firstRowSubheading":
				res = Asc.ETableStyleType.tablestyletypeFirstRowSubheading;
				break;
			case "secondRowSubheading":
				res = Asc.ETableStyleType.tablestyletypeSecondRowSubheading;
				break;
			case "thirdRowSubheading":
				res = Asc.ETableStyleType.tablestyletypeThirdRowSubheading;
				break;
			case "pageFieldLabels":
				res = Asc.ETableStyleType.tablestyletypePageFieldLabels;
				break;
			case "pageFieldValues":
				res = Asc.ETableStyleType.tablestyletypePageFieldValues;
				break;

		}
		return res;
	}

	function ToXml_ST_TableStyleType(val) {
		var res = null;
		switch (val) {
			case Asc.ETableStyleType.tablestyletypeWholeTable:
				res = "wholeTable";
				break;
			case Asc.ETableStyleType.tablestyletypeHeaderRow:
				res = "headerRow";
				break;
			case Asc.ETableStyleType.tablestyletypeTotalRow:
				res = "totalRow";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstColumn:
				res = "firstColumn";
				break;
			case Asc.ETableStyleType.tablestyletypeLastColumn:
				res = "lastColumn";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstRowStripe:
				res = "firstRowStripe";
				break;
			case Asc.ETableStyleType.tablestyletypeSecondRowStripe:
				res = "secondRowStripe";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstColumnStripe:
				res = "firstColumnStripe";
				break;
			case Asc.ETableStyleType.tablestyletypeSecondColumnStripe:
				res = "secondColumnStripe";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstHeaderCell:
				res = "firstHeaderCell";
				break;
			case Asc.ETableStyleType.tablestyletypeLastHeaderCell:
				res = "lastHeaderCell";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstTotalCell:
				res = "firstTotalCell";
				break;
			case Asc.ETableStyleType.tablestyletypeLastTotalCell:
				res = "lastTotalCell";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstSubtotalColumn:
				res = "firstSubtotalColumn";
				break;
			case Asc.ETableStyleType.tablestyletypeSecondSubtotalColumn:
				res = "secondSubtotalColumn";
				break;
			case Asc.ETableStyleType.tablestyletypeThirdSubtotalColumn:
				res = "thirdSubtotalColumn";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstSubtotalRow:
				res = "firstSubtotalRow";
				break;
			case Asc.ETableStyleType.tablestyletypeSecondSubtotalRow:
				res = "secondSubtotalRow";
				break;
			case Asc.ETableStyleType.tablestyletypeThirdSubtotalRow:
				res = "thirdSubtotalRow";
				break;
			case Asc.ETableStyleType.tablestyletypeBlankRow:
				res = "blankRow";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstColumnSubheading:
				res = "firstColumnSubheading";
				break;
			case Asc.ETableStyleType.tablestyletypeSecondColumnSubheading:
				res = "secondColumnSubheading";
				break;
			case Asc.ETableStyleType.tablestyletypeThirdColumnSubheading:
				res = "thirdColumnSubheading";
				break;
			case Asc.ETableStyleType.tablestyletypeFirstRowSubheading:
				res = "firstRowSubheading";
				break;
			case Asc.ETableStyleType.tablestyletypeSecondRowSubheading:
				res = "secondRowSubheading";
				break;
			case Asc.ETableStyleType.tablestyletypeThirdRowSubheading:
				res = "thirdRowSubheading";
				break;
			case Asc.ETableStyleType.tablestyletypePageFieldLabels:
				res = "pageFieldLabels";
				break;
			case Asc.ETableStyleType.tablestyletypePageFieldValues:
				res = "pageFieldValues";
				break;

		}
		return res;
	}

	function FromXml_ST_DataBarDirection(val) {
		var res = undefined;
		switch (val) {
			case "context":
				res = AscCommonExcel.EDataBarDirection.context;
				break;
			case "leftToRight":
				res = AscCommonExcel.EDataBarDirection.leftToRight;
				break;
			case "rightToLeft":
				res = AscCommonExcel.EDataBarDirection.rightToLeft;
				break;
		}
		return res;
	}

	function ToXml_ST_DataBarDirection(val) {
		var res = undefined;
		switch (val) {
			case AscCommonExcel.EDataBarDirection.context:
				res = "context";
				break;
			case AscCommonExcel.EDataBarDirection.leftToRight:
				res = "leftToRight";
				break;
			case AscCommonExcel.EDataBarDirection.rightToLeft:
				res = "rightToLeft";
				break;
		}
		return res;
	}

	function FromXml_ST_DataBarAxisPosition(val) {
		var res = undefined;
		switch (val) {
			case "automatic":
				res = AscCommonExcel.EDataBarAxisPosition.context;
				break;
			case "middle":
				res = AscCommonExcel.EDataBarAxisPosition.middle;
				break;
			case "none":
				res = AscCommonExcel.EDataBarAxisPosition.none;
				break;
		}
		return res;
	}

	function ToXml_ST_DataBarAxisPosition(val) {
		var res = undefined;
		switch (val) {
			case AscCommonExcel.EDataBarAxisPosition.automatic:
				res = "automatic";
				break;
			case AscCommonExcel.EDataBarAxisPosition.middle:
				res = "middle";
				break;
			case AscCommonExcel.EDataBarAxisPosition.none:
				res = "none";
				break;
		}
		return res;
	}

	function FromXml_ST_CellValueType(val) {
		var res = undefined;
		switch (val) {
			case "s":
				res = CellValueType.String;
				break;
			case "str":
				res = CellValueType.String;
				break;
			case "n":
				res = CellValueType.Number;
				break;
			case "e":
				res = CellValueType.Error;
				break;
			case "b":
				res =  CellValueType.Bool;
				break;
			case "inlineStr":
				res = CellValueType.String;
				break;
			case "d":
				res = CellValueType.String;
				break;
		}
		return res;
	}

	function ToXml_ST_CellValueType(val) {
		var res = undefined;
		switch (val) {
			case CellValueType.String:
				res = "s";
				break;
			/*case CellValueType.String:
				res = "str";
				break;*/
			case CellValueType.Number:
				res = "n";
				break;
			case CellValueType.Error:
				res = "e";
				break;
			case CellValueType.Bool:
				res =   "b";
				break;
			/*case "inlineStr":
				res = CellValueType.String;
				break;
			case "d":
				res = CellValueType.String;
				break;*/
		}
		return res;
	}

	function FromXml_ST_SheetViewType(val) {
		var res = null;
		switch (val) {
			case "normal":
				res = window["Asc"].c_oAscESheetViewType.normal;
				break;
			case "pageBreakPreview":
				res = window["Asc"].c_oAscESheetViewType.pageBreakPreview;
				break;
			case "pageLayout":
				res = window["Asc"].c_oAscESheetViewType.pageLayout;
				break;
		}
		return res;
	}

	function ToXml_ST_SheetViewType(val) {
		var res = null;
		switch (val) {
			case window["Asc"].c_oAscESheetViewType.normal:
				res = "normal";
				break;
			case window["Asc"].c_oAscESheetViewType.pageBreakPreview:
				res = "pageBreakPreview";
				break;
			case window["Asc"].c_oAscESheetViewType.pageLayout:
				res = "pageLayout";
				break;
		}
		return res;
	}

	function GradientFill() {
		//Attributes
		this.type = Asc.c_oAscFillGradType.GRAD_LINEAR;
		this.degree = 0;
		this.left = 0;
		this.right = 0;
		this.top = 0;
		this.bottom = 0;
		//Members
		this.stop = [];

		this._hash = null;
	}
	GradientFill.prototype.Properties = {
		type: 0,
		degree: 1,
		left: 2,
		right: 3,
		top: 4,
		bottom: 5,
		stop: 6
	};
	GradientFill.prototype.getType = function () {
		return UndoRedoDataTypes.StyleGradientFill;
	};
	GradientFill.prototype.getProperties = function () {
		return this.Properties;
	};
	GradientFill.prototype.getProperty = function (nType) {
		switch (nType) {
			case this.Properties.type:
				return this.type;
			case this.Properties.degree:
				return this.degree;
			case this.Properties.left:
				return this.left;
			case this.Properties.right:
				return this.right;
			case this.Properties.top:
				return this.top;
			case this.Properties.bottom:
				return this.bottom;
			case this.Properties.stop:
				return this.stop;
		}
	};
	GradientFill.prototype.setProperty = function (nType, value) {
		switch (nType) {
			case this.Properties.type:
				this.type = value;
				break;
			case this.Properties.degree:
				this.degree = value;
				break;
			case this.Properties.left:
				this.left = value;
				break;
			case this.Properties.right:
				this.right = value;
				break;
			case this.Properties.top:
				this.top = value;
				break;
			case this.Properties.bottom:
				this.bottom = value;
				break;
			case this.Properties.stop:
				this.stop = value;
				break;
		}
	};
	GradientFill.prototype.getHash = function() {
		if (!this._hash) {
			this._hash = this.type + ';' + this.degree + ';' + this.left + ';' + this.right + ';' + this.top + ';' +
				this.bottom + ';' + this.stop.length;
			for (var i = 0; i < this.stop.length; ++i) {
				this._hash += ';' + this.stop[i].getHash();
			}
		}
		return this._hash;
	};
	GradientFill.prototype.isEqual = function(gradientFill) {
		var res = this.type === gradientFill.type && this.degree === gradientFill.degree &&
			this.left === gradientFill.left && this.right === gradientFill.right && this.top === gradientFill.top &&
			this.bottom === gradientFill.bottom && this.stop.length == gradientFill.stop.length;
		if (res) {
			for (var i = 0; i < this.stop.length; ++i) {
				res = res && this.stop[i].isEqual(gradientFill.stop[i]);
			}
		}
		return res;
	};
	GradientFill.prototype.clone = function() {
		var res = new GradientFill();
		res.type = this.type;
		res.degree = this.degree;
		res.left = this.left;
		res.right = this.right;
		res.top = this.top;
		res.bottom = this.bottom;
		for (var i = 0; i < this.stop.length; ++i) {
			res.stop[i] = this.stop[i].clone();
		}
		return res;
	};
	GradientFill.prototype.notEmpty = function() {
		return true;
	};
	GradientFill.prototype.readAttributes = function(attr, uq) {
		if (attr()) {
			var vals = attr();
			var val;
			val = vals["type"];
			if (undefined !== val) {
				val = AscCommonExcel.FromXml_ST_GradientType(val);
				if (-1 !== val) {
					this.type = val;
				}
			}
			val = vals["degree"];
			if (undefined !== val) {
				this.degree = val - 0;
			}
			val = vals["left"];
			if (undefined !== val) {
				this.left = val - 0;
			}
			val = vals["right"];
			if (undefined !== val) {
				this.right = val - 0;
			}
			val = vals["top"];
			if (undefined !== val) {
				this.top = val - 0;
			}
			val = vals["bottom"];
			if (undefined !== val) {
				this.bottom = val - 0;
			}
		}
	};
	GradientFill.prototype.onStartNode = function(elem, attr, uq) {
		var newContext = this;
		if ("stop" === elem) {
			newContext = new GradientStop();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.stop.push(newContext);
		}
		else {
			newContext = null;
		}
		return newContext;
	};
	GradientFill.prototype.asc_getType = function () {
		return this.type;
	};
	GradientFill.prototype.asc_setType = function (value) {
		this.type = value;
	};
	GradientFill.prototype.asc_getDegree = function () {
		return this.degree;
	};
	GradientFill.prototype.asc_setDegree = function (value) {
		this.degree = value;
	};
	GradientFill.prototype.asc_getLeft = function () {
		return this.left;
	};
	GradientFill.prototype.asc_setLeft = function (value) {
		this.left = value;
	};
	GradientFill.prototype.asc_getRight = function () {
		return this.right;
	};
	GradientFill.prototype.asc_setRight = function (value) {
		this.right = value;
	};
	GradientFill.prototype.asc_getTop = function () {
		return this.top;
	};
	GradientFill.prototype.asc_setTop = function (value) {
		this.top = value;
	};
	GradientFill.prototype.asc_getBottom = function () {
		return this.bottom;
	};
	GradientFill.prototype.asc_setBottom = function (value) {
		this.bottom = value;
	};
	GradientFill.prototype.asc_getGradientStops = function () {
		var res = [];
		for (var i = 0; i < this.stop.length; ++i) {
			res[i] = this.stop[i].clone();
		}
		return res;
	};
	GradientFill.prototype.asc_putGradientStops = function (value) {
		this.stop = value;
	};
	function GradientStop() {
		//Attributes
		this.position = null;
		//Members
		this.color = null;

		this._hash = null;
	}
	GradientStop.prototype.Properties = {
		position: 0,
		color: 1
	};
	GradientStop.prototype.getType = function () {
		return UndoRedoDataTypes.StyleGradientFillStop;
	};
	GradientStop.prototype.getProperties = function () {
		return this.Properties;
	};
	GradientStop.prototype.getProperty = function (nType) {
		switch (nType) {
			case this.Properties.position:
				return this.position;
				break;
			case this.Properties.color:
				return this.color;
				break;
		}
	};
	GradientStop.prototype.setProperty = function (nType, value) {
		switch (nType) {
			case this.Properties.position:
				this.position = value;
				break;
			case this.Properties.color:
				this.color = value;
				break;
		}
	};
	GradientStop.prototype.getHash = function() {
		if (!this._hash) {
			var color = this.color ? this.color.getHash() : '';
			this._hash = this.position + ';' +color;
		}
		return this._hash;
	};
	GradientStop.prototype.isEqual = function(gradientStop) {
		return this.position === gradientStop.position && g_oColorManager.isEqual(this.color, gradientStop.color);
	};
	GradientStop.prototype.clone = function() {
		var res = new GradientStop();
		res.position = this.position;
		res.color = this.color;
		return res;
	};
	GradientStop.prototype.readAttributes = function(attr, uq) {
		if (attr()) {
			var vals = attr();
			var val;
			val = vals["position"];
			if (undefined !== val) {
				this.position = val - 0;
			}
		}
	};
	GradientStop.prototype.onStartNode = function(elem, attr, uq) {
		var newContext = this;
		if ("color" === elem) {
			this.color = AscCommon.getColorFromXml(attr);
		}
		else {
			newContext = null;
		}
		return newContext;
	};
	GradientStop.prototype.asc_getPosition = function () {
		return this.position;
	};
	GradientStop.prototype.asc_setPosition = function (value) {
		this.position = value;
	};
	GradientStop.prototype.asc_getColor = function () {
		return Asc.colorObjToAscColor(this.color);
	};
	GradientStop.prototype.asc_setColor = function (value) {
		this.color = CorrectAscColor(value);
	};
	function PatternFill() {
		//Attributes
		this.patternType = c_oAscPatternType.None;
		//Members
		this.fgColor = null;
		this.bgColor = null;

		this._hash = null;
	}
	PatternFill.prototype.Properties = {
		patternType: 0,
		fgColor: 1,
		bgColor: 2
	};
	PatternFill.prototype.getType = function () {
		return UndoRedoDataTypes.StylePatternFill;
	};
	PatternFill.prototype.getProperties = function () {
		return this.Properties;
	};
	PatternFill.prototype.getProperty = function (nType) {
		switch (nType) {
			case this.Properties.patternType:
				return this.patternType;
			case this.Properties.fgColor:
				return this.fgColor;
			case this.Properties.bgColor:
				return this.bgColor;
		}
	};
	PatternFill.prototype.setProperty = function (nType, value) {
		switch (nType) {
			case this.Properties.patternType:
				this.patternType = value;
				break;
			case this.Properties.fgColor:
				this.fgColor = value;
				break;
			case this.Properties.bgColor:
				this.bgColor = value;
				break;
		}
	};
	PatternFill.prototype.getHatchOffset = function () {
		return AscCommon.global_hatch_offsets[hatchFromExcelToWord(this.patternType)];
	};
	PatternFill.prototype.fromParams = function(type, color) {
		this.patternType = type;
		this.fgColor = color;
		this.bgColor = color;
	};
	PatternFill.prototype.fromColor = function(color) {
		this.fromParams(c_oAscPatternType.Solid, color);
	};
	PatternFill.prototype.getHash = function() {
		if (!this._hash) {
			this._hash = this.patternType + ';';
			if(this.fgColor){
				this._hash += this.fgColor.getHash();
			}
			this._hash += ';';
			if(this.bgColor){
				this._hash += this.bgColor.getHash();
			}
		}
		return this._hash;
	};
	PatternFill.prototype.isEqual = function(patternFill) {
		return this.patternType === patternFill.patternType &&
			g_oColorManager.isEqual(this.fgColor, patternFill.fgColor) &&
			g_oColorManager.isEqual(this.bgColor, patternFill.bgColor);
	};
	PatternFill.prototype.clone = function() {
		var res = new PatternFill();
		res.patternType = this.patternType;
		res.fgColor = this.fgColor;
		res.bgColor = this.bgColor;
		return res;
	};
	PatternFill.prototype.notEmpty = function() {
		return c_oAscPatternType.None !== this.patternType;
	};
	PatternFill.prototype.fixForDxf = function () {
		if ((c_oAscPatternType.None === this.patternType || c_oAscPatternType.Solid === this.patternType) && null !== this.bgColor) {
			this.patternType = c_oAscPatternType.Solid;
			var tmp = this.fgColor;
			this.fgColor = this.bgColor;
			this.bgColor = tmp || this.bgColor;
		}
	};
	PatternFill.prototype.readAttributes = function(attr, uq) {
		if (attr()) {
			var vals = attr();
			var val;
			val = vals["patternType"];
			if (undefined !== val) {
				val = FromXml_ST_PatternType(val);
				if (-1 !== val) {
					this.patternType = val;
				}
			}
		}
	};
	PatternFill.prototype.onStartNode = function(elem, attr, uq) {
		var newContext = this;
		if ("fgColor" === elem) {
			this.fgColor = AscCommon.getColorFromXml(attr);
		}
		else if ("bgColor" === elem) {
			this.bgColor = AscCommon.getColorFromXml(attr);
		}
		else {
			newContext = null;
		}
		return newContext;
	};
	PatternFill.prototype.asc_getType = function () {
		return c_oAscPatternType.Solid === this.patternType ? -1 : this.getHatchOffset();
	};
	PatternFill.prototype.asc_setType = function (value) {
		switch (value) {
			case -1:
				this.patternType = c_oAscPatternType.Solid;
				break;
			case 8:
				this.patternType = c_oAscPatternType.DarkDown;
				break;
			case 9:
				this.patternType = c_oAscPatternType.DarkHorizontal;
				break;
			case 10:
				this.patternType = c_oAscPatternType.DarkUp;
				break;
			case 11:
				this.patternType = c_oAscPatternType.DarkVertical;
				break;
			case 20:
				this.patternType = c_oAscPatternType.LightDown;
				break;
			case 21:
				this.patternType = c_oAscPatternType.LightHorizontal;
				break;
			case 22:
				this.patternType = c_oAscPatternType.LightUp;
				break;
			case 23:
				this.patternType = c_oAscPatternType.LightVertical;
				break;
			case 27:
				this.patternType = c_oAscPatternType.Gray0625;
				break;
			case 28:
				this.patternType = c_oAscPatternType.Gray125;
				break;
			case 29:
				this.patternType = c_oAscPatternType.LightGray;
				break;
			case 30:
				this.patternType = c_oAscPatternType.LightTrellis;
				break;
			case 33:
				this.patternType = c_oAscPatternType.MediumGray;
				break;
			case 35:
				this.patternType = c_oAscPatternType.DarkGray;
				break;
			case 41:
				this.patternType = c_oAscPatternType.DarkGrid;
				break;
			case 43:
				this.patternType = c_oAscPatternType.LightGrid;
				break;
			case 46:
			default:
				this.patternType = c_oAscPatternType.DarkTrellis;
				break;
		}
	};
	PatternFill.prototype.asc_getFgColor = function () {
		return Asc.colorObjToAscColor(this.fgColor);
	};
	PatternFill.prototype.asc_setFgColor = function (value) {
		this.fgColor = CorrectAscColor(value);
	};
	PatternFill.prototype.asc_getBgColor = function () {
		return Asc.colorObjToAscColor(this.bgColor);
	};
	PatternFill.prototype.asc_setBgColor = function (value) {
		this.bgColor = CorrectAscColor(value);
	};

	/** @constructor */
	function Fill() {
		this.patternFill = null;
		this.gradientFill = null;

		this._hash = null;
		this._index;
	}

	Fill.prototype.Properties = {
		patternFill: 0,
		gradientFill: 1
	};
	Fill.prototype.hasFill = function () {
		return ((this.patternFill && c_oAscPatternType.None !== this.patternFill.patternType) || this.gradientFill);
	};
	Fill.prototype.getSolidFill = function () {
		return (this.patternFill && c_oAscPatternType.Solid === this.patternFill.patternType) ? (this.patternFill.fgColor || createRgbColor(255, 255, 255)) : null;
	};
	Fill.prototype.bg = function () {
		var res = null;
		if (this.patternFill && c_oAscPatternType.None !== this.patternFill.patternType) {
			res = this.patternFill.fgColor || AscCommonExcel.g_oColorManager.getThemeColor(g_nColorTextDefault, 0);
		}
		return res;
	};
	Fill.prototype.fixForDxf = function () {
		if (this.patternFill) {
			this.patternFill.fixForDxf();
		}
	};
	Fill.prototype.fromColor = function (color) {
		this.patternFill = null;
		this.gradientFill = null;
		if (color) {
			this.patternFill = new PatternFill();
			this.patternFill.fromColor(color);
		}
		this.checkEmptyContent();
	};
	Fill.prototype.fromPatternParams = function (type, color) {
		this.patternFill = null;
		this.gradientFill = null;
		if (null !== type) {
			this.patternFill = new PatternFill();
			this.patternFill.fromParams(type, color);
		}
	};
	Fill.prototype.getHash = function () {
		if (!this._hash) {
			this._hash = (this.patternFill ? this.patternFill.getHash() : '') + '|';
			this._hash += (this.gradientFill ? this.gradientFill.getHash() : '');
		}
		return this._hash;
	};
	Fill.prototype.getIndexNumber = function () {
		return this._index;
	};
	Fill.prototype.setIndexNumber = function (val) {
		return this._index = val;
	};
	Fill.prototype.isEqual = function (fill) {
		if (this.patternFill && fill.patternFill) {
			return this.patternFill.isEqual(fill.patternFill);
		} else if (this.gradientFill && fill.gradientFill) {
			return this.gradientFill.isEqual(fill.gradientFill);
		}
		return false;
	};
	Fill.prototype.clone = function () {
		var res = new Fill();
		res.patternFill = this.patternFill ? this.patternFill.clone() : null;
		res.gradientFill = this.gradientFill ? this.gradientFill.clone() : null;
		return res;
	};
	Fill.prototype.getType = function () {
		return UndoRedoDataTypes.StyleFill;
	};
	Fill.prototype.getProperties = function () {
		return this.Properties;
	};
	Fill.prototype.getProperty = function (nType) {
		switch (nType) {
			case this.Properties.patternFill:
				return this.patternFill;
			case this.Properties.gradientFill:
				return this.gradientFill;
		}
	};
	Fill.prototype.setProperty = function (nType, value) {
		switch (nType) {
			case this.Properties.patternFill:
				this.patternFill = value;
				break;
			case this.Properties.gradientFill:
				this.gradientFill = value;
				break;
		}
	};
	Fill.prototype.notEmpty = function() {
		return (this.patternFill && this.patternFill.notEmpty()) || (this.gradientFill && this.gradientFill.notEmpty());
	};
	Fill.prototype.onStartNode = function (elem, attr, uq) {
		var newContext = this;
		if ("gradientFill" === elem) {
			newContext = new GradientFill();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.gradientFill = newContext;
		} else if ("patternFill" === elem) {
			newContext = new PatternFill();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.patternFill = newContext;
		} else {
			newContext = null;
		}
		return newContext;
	};
	Fill.prototype.onEndNode = function (prevContext, elem) {
		if ("patternFill" === elem && AscCommon.openXml.SaxParserDataTransfer.priorityBg) {
			prevContext.fixForDxf();
		}
	};
	Fill.prototype.asc_getPatternFill = function () {
		return this.patternFill && this.patternFill.notEmpty() ? this.patternFill : null;
	};
	Fill.prototype.asc_setPatternFill = function (value) {
		this.patternFill = value;
		this.gradientFill = null;
	};
	Fill.prototype.asc_getGradientFill = function () {
		return this.gradientFill;
	};
	Fill.prototype.asc_setGradientFill = function (value) {
		this.patternFill = null;
		this.gradientFill = value;
	};
	Fill.prototype.checkEmptyContent = function () {
		if (!this.patternFill && !this.gradientFill) {
			this.fromPatternParams(AscCommonExcel.c_oAscPatternType.None, null);
		}
	};

	function FromXml_ST_BorderStyle(val) {
		var res = -1;
		if ("none" === val) {
			res = c_oAscBorderStyles.None;
		} else if ("thin" === val) {
			res = c_oAscBorderStyles.Thin;
		} else if ("medium" === val) {
			res = c_oAscBorderStyles.Medium;
		} else if ("dashed" === val) {
			res = c_oAscBorderStyles.Dashed;
		} else if ("dotted" === val) {
			res = c_oAscBorderStyles.Dotted;
		} else if ("thick" === val) {
			res = c_oAscBorderStyles.Thick;
		} else if ("double" === val) {
			res = c_oAscBorderStyles.Double;
		} else if ("hair" === val) {
			res = c_oAscBorderStyles.Hair;
		} else if ("mediumDashed" === val) {
			res = c_oAscBorderStyles.MediumDashed;
		} else if ("dashDot" === val) {
			res = c_oAscBorderStyles.DashDot;
		} else if ("mediumDashDot" === val) {
			res = c_oAscBorderStyles.MediumDashDot;
		} else if ("dashDotDot" === val) {
			res = c_oAscBorderStyles.DashDotDot;
		} else if ("mediumDashDotDot" === val) {
			res = c_oAscBorderStyles.MediumDashDotDot;
		} else if ("slantDashDot" === val) {
			res = c_oAscBorderStyles.SlantDashDot;
		}
		return res;
	}

	function ToXml_ST_BorderStyle(val) {
		var res = null;
		switch (val) {
			case  Asc.c_oAscBorderStyles.DashDot:
				res = "dashDot";
				break;
			case  Asc.c_oAscBorderStyles.DashDotDot:
				res = "dashDotDot";
				break;
			case  Asc.c_oAscBorderStyles.Dashed:
				res = "dashed";
				break;
			case  Asc.c_oAscBorderStyles.Dotted:
				res = "dotted";
				break;
			case  Asc.c_oAscBorderStyles.Double:
				res = "double";
				break;
			case  Asc.c_oAscBorderStyles.Hair:
				res = "hair";
				break;
			case  Asc.c_oAscBorderStyles.Medium:
				res = "medium";
				break;
			case  Asc.c_oAscBorderStyles.MediumDashDot:
				res = "mediumDashDot";
				break;
			case  Asc.c_oAscBorderStyles.MediumDashDotDot:
				res = "mediumDashDotDot";
				break;
			case  Asc.c_oAscBorderStyles.MediumDashed:
				res = "mediumDashed";
				break;
			case  Asc.c_oAscBorderStyles.None:
				//res = "none";
				break;
			case  Asc.c_oAscBorderStyles.SlantDashDot:
				res = "slantDashDot";
				break;
			case  Asc.c_oAscBorderStyles.Thick:
				res = "thick";
				break;
			case  Asc.c_oAscBorderStyles.Thin:
				res = "thin";
				break;
		}
		return res;
	}



	var g_oBorderPropProperties = {
		s: 0, c: 1
	};

	function BorderProp() {
		this.s = c_oAscBorderStyles.None;
		this.w = c_oAscBorderWidth.None;
		this.c = g_oColorManager.getThemeColor(1);
	}
	BorderProp.prototype.Properties = g_oBorderPropProperties;
	BorderProp.prototype.getHash = function() {
		if (!this._hash) {
			var color = this.c ? this.c.getHash() : '';
			this._hash = this.s + ';' + this.w + ';' + color;
		}
		return this._hash;
	};
	BorderProp.prototype.setStyle = function (style) {
		this.s = style;
		switch (this.s) {
			case c_oAscBorderStyles.Thin:
			case c_oAscBorderStyles.DashDot:
			case c_oAscBorderStyles.DashDotDot:
			case c_oAscBorderStyles.Dashed:
			case c_oAscBorderStyles.Dotted:
			case c_oAscBorderStyles.Hair:
				this.w = c_oAscBorderWidth.Thin;
				break;
			case c_oAscBorderStyles.Medium:
			case c_oAscBorderStyles.MediumDashDot:
			case c_oAscBorderStyles.MediumDashDotDot:
			case c_oAscBorderStyles.MediumDashed:
			case c_oAscBorderStyles.SlantDashDot:
				this.w = c_oAscBorderWidth.Medium;
				break;
			case c_oAscBorderStyles.Thick:
			case c_oAscBorderStyles.Double:
				this.w = c_oAscBorderWidth.Thick;
				break;
			default:
				this.w = c_oAscBorderWidth.None;
				break;
		}
	};
	BorderProp.prototype.getDashSegments = function () {
		var res;
		switch (this.s) {
			case c_oAscBorderStyles.Hair:
				res = [1, 1];
				break;
			case c_oAscBorderStyles.Dotted:
				res = [2, 2];
				break;
			case c_oAscBorderStyles.DashDotDot:
			case c_oAscBorderStyles.MediumDashDotDot:
				res = [3, 3, 3, 3, 9, 3];
				break;
			case c_oAscBorderStyles.DashDot:
			case c_oAscBorderStyles.MediumDashDot:
			case c_oAscBorderStyles.SlantDashDot:
				res = [3, 3, 9, 3];
				break;
			case c_oAscBorderStyles.Dashed:
				res = [3, 1];
				break;
			case c_oAscBorderStyles.MediumDashed:
				res = [9, 3];
				break;
			case c_oAscBorderStyles.Thin:
			case c_oAscBorderStyles.Medium:
			case c_oAscBorderStyles.Thick:
			case c_oAscBorderStyles.Double:
			default:
				res = [];
				break;
		}
		return res;
	};
	BorderProp.prototype.getRgbOrNull = function () {
		var nRes = null;
		if (null != this.c) {
			nRes = this.c.getRgb();
		}
		return nRes;
	};
	BorderProp.prototype.getColorOrDefault = function () {
		return this.c || g_oDefaultFormat.ColorAuto;
	};
	BorderProp.prototype.isEmpty = function () {
		return c_oAscBorderStyles.None === this.s;
	};
	BorderProp.prototype.isEqual = function (val, byRgb) {
		return val && this.s === val.s && g_oColorManager.isEqual(this.c, val.c, byRgb);
	};
	BorderProp.prototype.isEqual2 = function (val1, val2, byRgb) {
		return ((!val1 || val1.isEmpty()) && (!val2 || val2.isEmpty())) || (val1 && val2 && val1.isEqual(val2, byRgb))
	};
	BorderProp.prototype.clone = function () {
		var res = new BorderProp();
		res.merge(this);
		return res;
	};
	BorderProp.prototype.merge = function (oBorderProp) {
		if (oBorderProp && !oBorderProp.isEmpty()) {
			this.s = oBorderProp.s;
			this.w = oBorderProp.w;
			this.c = oBorderProp.c;
		}
	};
	BorderProp.prototype.getType = function () {
		return UndoRedoDataTypes.StyleBorderProp;
	};
	BorderProp.prototype.getProperties = function () {
		return this.Properties;
	};
	BorderProp.prototype.getProperty = function (nType) {
		switch (nType) {
			case this.Properties.s:
				return this.s;
				break;
			case this.Properties.c:
				return this.c;
				break;
		}
	};
	BorderProp.prototype.setProperty = function (nType, value) {
		switch (nType) {
			case this.Properties.s:
				this.setStyle(value);
				break;
			case this.Properties.c:
				this.c = value;
				break;
		}
	};
	BorderProp.prototype.readAttributes = function(attr, uq) {
		if(attr()){
			var vals = attr();
			var val;
			val = vals["style"];
			if(undefined !== val){
				val = FromXml_ST_BorderStyle(val);
				if(-1 !== val){
					this.setStyle(val);
				}
			}
		}
	};
	BorderProp.prototype.onStartNode = function(elem, attr, uq) {
		var newContext = this;
		if("color" === elem){
			this.c = AscCommon.getColorFromXml(attr);
		}
		else {
			newContext = null;
		}
		return newContext;
	};

	/** @constructor */
	function Border() {
		this.l = null;
		this.t = null;
		this.r = null;
		this.b = null;
		this.d = null;
		this.ih = null;
		this.iv = null;
		this.dd = false;
		this.du = false;

		this._hash = null;
		this._index;
	}

	Border.prototype.Properties = {
		l: 0,
		t: 1,
		r: 2,
		b: 3,
		d: 4,
		ih: 5,
		iv: 6,
		dd: 7,
		du: 8
	};
	Border.prototype.getHash = function() {
		if (!this._hash) {
			this._hash = (this.l ? this.l.getHash() : '') + '|';
			this._hash += (this.t ? this.t.getHash() : '') + '|';
			this._hash += (this.r ? this.r.getHash() : '') + '|';
			this._hash += (this.b ? this.b.getHash() : '') + '|';
			this._hash += (this.d ? this.d.getHash() : '') + '|';
			this._hash += (this.ih ? this.ih.getHash() : '') + '|';
			this._hash += (this.iv ? this.iv.getHash() : '') + '|';
			this._hash += this.dd + '|';
			this._hash += this.du;
		}
		return this._hash;
	};
	Border.prototype.getIndexNumber = function() {
		return this._index;
	};
	Border.prototype.setIndexNumber = function(val) {
		return this._index = val;
	};
	Border.prototype.initDefault = function() {
		this.l = new BorderProp();
		this.t = new BorderProp();
		this.r = new BorderProp();
		this.b = new BorderProp();
		this.d = new BorderProp();
	};
	Border.prototype.getL = function() {
		return this.l || g_oDefaultFormat.Border.l;
	};
	Border.prototype.getT = function() {
		return this.t || g_oDefaultFormat.Border.t;
	};
	Border.prototype.getR = function() {
		return this.r || g_oDefaultFormat.Border.r;
	};
	Border.prototype.getB = function() {
		return this.b || g_oDefaultFormat.Border.b;
	};
	Border.prototype.getD = function() {
		return this.d || g_oDefaultFormat.Border.d;
	};
	Border.prototype.getIH = function() {
		return this.ih || g_oDefaultFormat.Border.ih;
	};
	Border.prototype.getIV = function() {
		return this.iv || g_oDefaultFormat.Border.iv;
	};
	Border.prototype._mergeProperty = function (first, second) {
		if (first && !first.isEmpty())
			return first.clone();
		else if (second)
			return second.clone();
		else
			return null;
	};
	Border.prototype.merge = function (border, isTable) {
		var oRes = new Border();
		if (isTable) {
			oRes.l = this._mergeProperty(this.l, border.l);
			oRes.t = this._mergeProperty(this.t, border.t);
			oRes.r = this._mergeProperty(this.r, border.r);
			oRes.b = this._mergeProperty(this.b, border.b);
			oRes.ih = this._mergeProperty(this.ih, border.ih);
			oRes.iv = this._mergeProperty(this.iv, border.iv);
			oRes.d = this._mergeProperty(this.d, border.d);
			oRes.dd = this.dd || border.dd;
			oRes.du = this.du || border.du;
		} else {
			//todo merge with default
			oRes.l = this.l ? this.l.clone() : null;
			oRes.t = this.t ? this.t.clone() : null;
			oRes.r = this.r ? this.r.clone() : null;
			oRes.b = this.b ? this.b.clone() : null;
			oRes.ih = this.ih ? this.ih.clone() : null;
			oRes.iv = this.iv ? this.iv.clone() : null;
			oRes.d = this._mergeProperty(this.d, border.d);
			oRes.dd = this.dd || border.dd;
			oRes.du = this.du || border.du;
		}
		return oRes;
	};
	Border.prototype.intersect = function (border, byRgb) {
		if (!BorderProp.prototype.isEqual2(this.l, border.l, byRgb)) {
			this.l = new BorderProp();
		}
		if (!BorderProp.prototype.isEqual2(this.t, border.t, byRgb)) {
			this.t = new BorderProp();
		}
		if (!BorderProp.prototype.isEqual2(this.r, border.r, byRgb)) {
			this.r = new BorderProp();
		}
		if (!BorderProp.prototype.isEqual2(this.b, border.b, byRgb)) {
			this.b = new BorderProp();
		}
		if (!BorderProp.prototype.isEqual2(this.d, border.d, byRgb)) {
			this.d = new BorderProp();
			this.dd = false;
			this.du = false;
		}
		if (!BorderProp.prototype.isEqual2(this.ih, border.ih, byRgb)) {
			this.ih = null;
		}
		if (!BorderProp.prototype.isEqual2(this.iv, border.iv, byRgb)) {
			this.iv = null;
		}
		if (this.dd !== border.dd) {
			this.dd = false;
		}
		if (this.du !== border.du) {
			this.du = false;
		}
	};
	Border.prototype.isEqual = function(val) {
		return BorderProp.prototype.isEqual2(this.l, val.l) && BorderProp.prototype.isEqual2(this.t, val.t)
			&& BorderProp.prototype.isEqual2(this.r, val.r) && BorderProp.prototype.isEqual2(this.b, val.b)
			&& BorderProp.prototype.isEqual2(this.d, val.d) && this.dd === val.dd && this.du === val.du
			&& BorderProp.prototype.isEqual2(this.ih, val.ih) && BorderProp.prototype.isEqual2(this.iv, val.iv);
	};
	Border.prototype.clone = function () {
		var res = new Border();
		res.l = this.l ? this.l.clone() : null;
		res.t = this.t ? this.t.clone() : null;
		res.r = this.r ? this.r.clone() : null;
		res.b = this.b ? this.b.clone() : null;
		res.d = this.d ? this.d.clone() : null;
		res.ih = this.ih ? this.ih.clone() : null;
		res.iv = this.iv ? this.iv.clone() : null;
		res.dd = this.dd;
		res.du = this.du;
		return res;
	};
	Border.prototype.clean = function () {
		this.l = null;
		this.t = null;
		this.r = null;
		this.b = null;
		this.d = null;
		this.ih = null;
		this.iv = null;
		this.dd = false;
		this.du = false;
	};
	Border.prototype.mergeInner = function (border) {
		if (border) {
			if (border.l) {
				if (this.l) {
					this.l.merge(border.l);
				} else {
					this.l = border.l.clone();
				}
			}
			if (border.t) {
				if (this.t) {
					this.t.merge(border.t);
				} else {
					this.t = border.t.clone();
				}
			}
			if (border.r) {
				if (this.r) {
					this.r.merge(border.r);
				} else {
					this.r = border.r.clone();
				}
			}
			if (border.b) {
				if (this.b) {
					this.b.merge(border.b);
				} else {
					this.b = border.b.clone();
				}
			}
			if (border.d) {
				if (this.d) {
					this.d.merge(border.d);
				} else {
					this.d = border.d.clone();
				}
			}
			if (border.ih) {
				if (this.ih) {
					this.ih.merge(border.ih);
				} else {
					this.ih = border.ih.clone();
				}
			}
			if (border.iv) {
				if (this.iv) {
					this.iv.merge(border.iv);
				} else {
					this.iv = border.iv.clone();
				}
			}
			if (null != border.dd) {
				this.dd = this.dd || border.dd;
			}
			if (null != border.du) {
				this.du = this.du || border.du;
			}
		}
	};
	Border.prototype.getType = function () {
		return UndoRedoDataTypes.StyleBorder;
	};
	Border.prototype.getProperties = function () {
		return this.Properties;
	};
	Border.prototype.getProperty = function (nType) {
		switch (nType) {
			case this.Properties.l:
				return this.l;
				break;
			case this.Properties.t:
				return this.t;
				break;
			case this.Properties.r:
				return this.r;
				break;
			case this.Properties.b:
				return this.b;
				break;
			case this.Properties.d:
				return this.d;
				break;
			case this.Properties.ih:
				return this.ih;
				break;
			case this.Properties.iv:
				return this.iv;
				break;
			case this.Properties.dd:
				return this.dd;
				break;
			case this.Properties.du:
				return this.du;
				break;
		}
	};
	Border.prototype.setProperty = function (nType, value) {
		switch (nType) {
			case this.Properties.l:
				this.l = value;
				break;
			case this.Properties.t:
				this.t = value;
				break;
			case this.Properties.r:
				this.r = value;
				break;
			case this.Properties.b:
				this.b = value;
				break;
			case this.Properties.d:
				this.d = value;
				break;
			case this.Properties.ih:
				this.ih = value;
				break;
			case this.Properties.iv:
				this.iv = value;
				break;
			case this.Properties.dd:
				this.dd = value;
				break;
			case this.Properties.du:
				this.du = value;
				break;
		}
	};
	Border.prototype.notEmpty = function () {
		return (this.l && !this.l.isEmpty()) || (this.r && !this.r.isEmpty()) ||
			(this.t && !this.t.isEmpty()) || (this.b && !this.b.isEmpty()) ||
			(this.d && !this.d.isEmpty() && (this.dd || this.du));
	};
	Border.prototype.readAttributes = function(attr, uq) {
		if(attr()){
			var vals = attr();
			var val;
			val = vals["diagonalUp"];
			if(undefined !== val){
				this.du = AscCommon.getBoolFromXml(val);
			}
			val = vals["diagonalDown"];
			if(undefined !== val){
				this.dd = AscCommon.getBoolFromXml(val);
			}
		}
	};
	Border.prototype.onStartNode = function(elem, attr, uq) {
		var newContext = this;
		if("start" === elem || "left" === elem){
			newContext = new BorderProp();
			if(newContext.readAttributes){
				newContext.readAttributes(attr, uq);
			}
			this.l = newContext;
		}
		else if("end" === elem || "right" === elem){
			newContext = new BorderProp();
			if(newContext.readAttributes){
				newContext.readAttributes(attr, uq);
			}
			this.r = newContext;
		}
		else if("top" === elem){
			newContext = new BorderProp();
			if(newContext.readAttributes){
				newContext.readAttributes(attr, uq);
			}
			this.t = newContext;
		}
		else if("bottom" === elem){
			newContext = new BorderProp();
			if(newContext.readAttributes){
				newContext.readAttributes(attr, uq);
			}
			this.b = newContext;
		}
		else if("diagonal" === elem){
			newContext = new BorderProp();
			if(newContext.readAttributes){
				newContext.readAttributes(attr, uq);
			}
			this.d = newContext;
		}
		else if("vertical" === elem){
			newContext = new BorderProp();
			if(newContext.readAttributes){
				newContext.readAttributes(attr, uq);
			}
			this.iv = newContext;
		}
		else if("horizontal" === elem){
			newContext = new BorderProp();
			if(newContext.readAttributes){
				newContext.readAttributes(attr, uq);
			}
			this.ih = newContext;
		}
		else {
			newContext = null;
		}
		return newContext;
	};

	var g_oNumProperties = {
		f: 0,
		id: 1
	};

	/** @constructor */
	function Num(val) {
		if (null == val)
			val = g_oDefaultFormat.NumAbs;
		this.f = val.f;
		this.id = val.id;

		this._hash;
		this._index;
	}

	Num.prototype.Properties = g_oNumProperties;
	Num.prototype.getHash = function () {
		if (!this._hash) {
			this._hash = this.f + '|' + this.id;
		}
		return this._hash;
	};
	Num.prototype.getIndexNumber = function () {
		return this._index;
	};
	Num.prototype.setIndexNumber = function (val) {
		this._index = val;
	};
	Num.prototype.initFromParams = function (id, format, oNumFmts) {
		var res = oNumFmts && oNumFmts[id];
		if (res) {
			return res;
		}
		res = new Num();
		if (format) {
			res.f = format;
		} else {
			res.f = AscCommonExcel.aStandartNumFormats[id];
		}
		if (!res.f) {
			res.f = "General";
		}
		if (AscCommon.canGetFormatByStandardId(id)) {
			res.id = id;
		}
		var numFormat = AscCommon.oNumFormatCache.get(res.f);
		numFormat.checkCultureInfoFontPicker();
		res = g_StyleCache.addNum(res);
		if (oNumFmts) {
			oNumFmts[res.id] = res;
		}
		return res;
	};
	Num.prototype.setFormat = function (f, opt_id) {
		this.f = f;
		this.id = opt_id;
	};
	Num.prototype.getFormat = function () {
		return (null != this.id) ? (AscCommon.getFormatByStandardId(this.id) || this.f) : this.f;
	};
	Num.prototype.getNumFormat = function () {
		return AscCommon.oNumFormatCache.get(this.getFormat());
	};
	Num.prototype._mergeProperty = function (first, second, def) {
		if (def != first)
			return first;
		else
			return second;
	};
	Num.prototype.merge = function (num) {
		var oRes = new Num();
		oRes.f = this._mergeProperty(this.f, num.f, g_oDefaultFormat.Num.f);
		oRes.id = this._mergeProperty(this.id, num.id, g_oDefaultFormat.Num.id);
		return oRes;
	};
	Num.prototype.isEqual = function (val) {
		if (null != this.id && null != val.id) {
			return this.id == val.id;
		} else if (null != this.id || null != val.id) {
			return false;
		} else {
			return this.f == val.f;
		}
	};
	Num.prototype.clone = function () {
		return new Num(this);
	};
	Num.prototype.getType = function () {
		return UndoRedoDataTypes.StyleNum;
	};
	Num.prototype.getProperties = function () {
		return this.Properties;
	};
	Num.prototype.getProperty = function (nType) {
		switch (nType) {
			case this.Properties.f:
				return this.f;
				break;
			case this.Properties.id:
				return this.id;
				break;
		}
	};
	Num.prototype.setProperty = function (nType, value) {
		switch (nType) {
			case this.Properties.f:
				this.f = value;
				break;
			case this.Properties.id:
				this.id = value;
				break;
		}
	};
	Num.prototype.readAttributes = function (attr, uq) {
		if (attr()) {
			var vals = attr();
			var val;
			val = vals["numFmtId"];
			var sFormat = null;
			var id;
			if (undefined !== val) {
				id = val - 0;
			}
			val = vals["formatCode"];
			if (undefined !== val) {
				sFormat = AscCommon.unleakString(uq(val));
			}
			this.f = null != sFormat ? sFormat : (AscCommonExcel.aStandartNumFormats[id] || "General");
			if (AscCommon.canGetFormatByStandardId(id)) {
				this.id = id;
			}
		}
	};

    var g_oCellXfsProperties = {
        border: 0,
        fill: 1,
        font: 2,
        num: 3,
        align: 4,
        QuotePrefix: 5,
        XfId: 6,
        PivotButton: 7,
		applyProtection: 8,
		hidden: 9,
		locked: 10
    };

    /** @constructor */
    function CellXfs() {
        this.border = null;
        this.fill = null;
        this.font = null;
        this.num = null;
        this.align = null;
        this.QuotePrefix = null;
        this.PivotButton = null;
        this.XfId = null;

        this.applyProtection = null;
        this.locked = null;
        this.hidden = null;

        //inner
        this._hash;
        this._index;
        this.operationCache = {};
    }

    CellXfs.prototype.Properties = g_oCellXfsProperties;
    CellXfs.prototype.getHash = function () {
        if (!this._hash) {
            this._hash = (this.border ? this.border.getIndexNumber() : '') + '|';
            this._hash += (this.fill ? this.fill.getIndexNumber() : '') + '|';
            this._hash += (this.font ? this.font.getIndexNumber() : '') + '|';
            this._hash += (this.num ? this.num.getIndexNumber() : '') + '|';
            this._hash += (this.align ? this.align.getIndexNumber() : '') + '|';
            this._hash += this.QuotePrefix + '|';
            this._hash += this.PivotButton + '|';
            this._hash += this.XfId + '|';
            this._hash += this.applyProtection + '|';
            this._hash += this.locked + '|';
            this._hash += this.hidden + '|';
        }
        return this._hash;
    };
    CellXfs.prototype.getIndexNumber = function () {
        return this._index;
    };
    CellXfs.prototype.setIndexNumber = function (val) {
        this._index = val;
    };
    CellXfs.prototype._mergeProperty = function (addFunc, first, second, isTable, isTableColor) {
        var res = null;
        if (null != first || null != second) {
            if (null == first)
                res = second;
            else if (null == second)
                res = first;
            else {
                if (null != first.merge) {
                    res = addFunc.call(g_StyleCache, first.merge(second, isTable, isTableColor));
                } else {
                    res = first;
                }
            }
        }
        return res;
    };
	CellXfs.prototype.isNormalFont = function () {
		return g_StyleCache.firstXf === this || g_StyleCache.normalXf.font === this.font;
	};
	CellXfs.prototype.isNormalFill = function () {
		return g_StyleCache.firstXf === this || g_StyleCache.normalXf.fill === this.fill;
	};
    CellXfs.prototype.merge = function (xfs, isTable, isTableBorders) {
        var xfIndexNumber = xfs.getIndexNumber();
        if (undefined === xfIndexNumber) {
            xfs = g_StyleCache.addXf(xfs);
            xfIndexNumber = xfs.getIndexNumber();
        }
        var cache = this.getOperationCache("merge", xfIndexNumber);
        if (!cache) {
            cache = new CellXfs();
			//todo test isTable insted of isTableBorders
			if ((isTable || isTableBorders) && (g_StyleCache.firstXf === xfs || g_StyleCache.normalXf.border === xfs.border)) {
				cache.border = this._mergeProperty(g_StyleCache.addBorder, this.border, xfs.border, (isTable || isTableBorders));
			} else {
				cache.border = this._mergeProperty(g_StyleCache.addBorder, xfs.border, this.border, (isTable || isTableBorders));
			}
			if (isTable && (g_StyleCache.firstXf === xfs || g_StyleCache.normalXf.fill === xfs.fill)) {
                if (g_StyleCache.normalXf.fill === xfs.fill) {
                    cache.fill = this._mergeProperty(g_StyleCache.addFill, this.fill, g_oDefaultFormat.Fill);
                } else {
                    cache.fill = this._mergeProperty(g_StyleCache.addFill, this.fill, xfs.fill);
                }
            } else {
                cache.fill = this._mergeProperty(g_StyleCache.addFill, xfs.fill, this.fill);
            }
			var isTableColor = true;
			if (isTable && xfs.isNormalFont()) {
				if (g_StyleCache.normalXf.font === xfs.font) {
					cache.font = this._mergeProperty(g_StyleCache.addFont, g_oDefaultFormat.Font, this.font, isTable, isTableColor);
				} else {
					cache.font = this._mergeProperty(g_StyleCache.addFont, xfs.font, this.font, isTable, isTableColor);
				}
			} else {
				isTableColor = isTable && xfs.font && xfs.font.isNormalXfColor();
				cache.font = this._mergeProperty(g_StyleCache.addFont, xfs.font, this.font, isTable, isTableColor);
			}
            cache.num = this._mergeProperty(g_StyleCache.addNum, xfs.num, this.num);
            cache.align = this._mergeProperty(g_StyleCache.addAlign, xfs.align, this.align);
            cache.QuotePrefix = this._mergeProperty(null, xfs.QuotePrefix, this.QuotePrefix);
            cache.PivotButton = this._mergeProperty(null, xfs.PivotButton, this.PivotButton);
            cache.XfId = this._mergeProperty(null, xfs.XfId, this.XfId);
            cache.applyProtection = this._mergeProperty(null, xfs.applyProtection, this.applyProtection);
            cache.locked = this._mergeProperty(null, xfs.locked, this.locked);
            cache.hidden = this._mergeProperty(null, xfs.hidden, this.hidden);
            cache = g_StyleCache.addXf(cache);
            this.setOperationCache("merge", xfIndexNumber, cache);
        }
        return cache;
    };
    CellXfs.prototype.clone = function () {
        var res = new CellXfs();
		if (this.border) {
			res.border = this.border.clone();
		}
		if (this.fill) {
			res.fill = this.fill.clone();
		}
		if (this.font) {
			res.font = this.font.clone();
		}
        res.num = this.num;
        res.align = this.align;
        res.QuotePrefix = this.QuotePrefix;
        res.PivotButton = this.PivotButton;
        res.XfId = this.XfId;
        res.applyProtection = this.applyProtection;
        res.locked = this.locked;
        res.hidden = this.hidden;

        return res;
    };
    CellXfs.prototype.isEqual = function (xfs) {
		return this.font === xfs.font && this.fill === xfs.fill && this.border === xfs.border && this.num === xfs.num &&
			this.align === xfs.align && this.QuotePrefix === xfs.QuotePrefix && this.PivotButton === xfs.PivotButton &&
			this.XfId === xfs.XfId && this.applyProtection === xfs.applyProtection && this.locked === xfs.locked &&
			this.hidden === xfs.hidden;
    };
    CellXfs.prototype.getType = function () {
        return UndoRedoDataTypes.StyleXfs;
    };
    CellXfs.prototype.getProperties = function () {
        return this.Properties;
    };
    CellXfs.prototype.getProperty = function (nType) {
        switch (nType) {
            case this.Properties.border:
                return this.border;
            case this.Properties.fill:
                return this.fill;
            case this.Properties.font:
                return this.font;
            case this.Properties.num:
                return this.num;
            case this.Properties.align:
                return this.align;
            case this.Properties.QuotePrefix:
                return this.QuotePrefix;
            case this.Properties.PivotButton:
                return this.PivotButton;
            case this.Properties.XfId:
                return this.XfId;
            case this.Properties.applyProtection:
                return this.applyProtection;
            case this.Properties.locked:
                return this.locked;
            case this.Properties.hidden:
                return this.hidden;
        }
    };
    CellXfs.prototype.setProperty = function (nType, value) {
        switch (nType) {
            case this.Properties.border:
                this.border = value;
                break;
            case this.Properties.fill:
                this.fill = value;
                break;
            case this.Properties.font:
                this.font = value;
                break;
            case this.Properties.num:
                this.num = value;
                break;
            case this.Properties.align:
                this.align = value;
                break;
            case this.Properties.QuotePrefix:
                this.QuotePrefix = value;
                break;
            case this.Properties.PivotButton:
                this.PivotButton = value;
                break;
            case this.Properties.XfId:
                this.XfId = value;
                break;
            case this.Properties.applyProtection:
                this.applyProtection = value;
                break;
            case this.Properties.locked:
                this.locked = value;
                break;
            case this.Properties.hidden:
                this.hidden = value;
                break;
        }
    };
	/*CellXfs.prototype.Write_ToBinary2 = function (writer) {
		var t = this;
		var oBinaryStylesTableWriter = new AscCommonExcel.BinaryStylesTableWriter(writer);
		oBinaryStylesTableWriter.bs.WriteItem(0, function(){oBinaryStylesTableWriter.WriteDxf(t);});
	};
	CellXfs.prototype.Read_FromBinary2 = function (reader) {
		var api_sheet = Asc['editor'];
		var wb = api_sheet.wbModel;
		var bsr = new AscCommonExcel.Binary_StylesTableReader(reader, wb);
		var bcr = new AscCommon.Binary_CommonReader(reader);
		reader.GetUChar();
		var oThis = this;
		var length = reader.GetULongLE();
		bcr.Read1(length, function (t, l) {
			return bsr.ReadDxf(t, l, oThis);
		});
		return this;
	};*/
    CellXfs.prototype.getBorder = function () {
        return this.border;
    };
    CellXfs.prototype.setBorder = function (val) {
        this.border = val;
    };
    CellXfs.prototype.getFill = function () {
        return this.fill;
    };
	CellXfs.prototype.getFill2 = function () {
		// ToDo check this! Rename to getFill
		return this.fill || g_oDefaultFormat.Fill;
	};
    CellXfs.prototype.setFill = function (val) {
        this.fill = val;
    };
    CellXfs.prototype.getFont = function () {
        return this.font;
    };
	CellXfs.prototype.getFont2 = function () {
		// ToDo check this! Rename to getFont
		return this.font || g_oDefaultFormat.Font;
	};
    CellXfs.prototype.setFont = function (val) {
        this.font = val;
    };
    CellXfs.prototype.getNum = function () {
        return this.num;
    };
	CellXfs.prototype.getNum2 = function () {
		// ToDo check this! Rename to getNum
		return this.num || g_oDefaultFormat.Num;
	};
    CellXfs.prototype.setNum = function (val) {
        this.num = val;
    };
    CellXfs.prototype.getAlign = function () {
        return this.align;
    };
    CellXfs.prototype.getAlign2 = function () {
        // ToDo check this! Rename to getAlign
        return this.align || g_oDefaultFormat.AlignAbs;
    };
    CellXfs.prototype.setAlign = function (val) {
        this.align = val;
    };
    CellXfs.prototype.getQuotePrefix = function () {
        return this.QuotePrefix;
    };
    CellXfs.prototype.setQuotePrefix = function (val) {
        this.QuotePrefix = val;
    };
	CellXfs.prototype.getApplyProtection = function () {
		return this.applyProtection;
	};
	CellXfs.prototype.setApplyProtection = function (val) {
		this.applyProtection = val;
	};
	CellXfs.prototype.getLocked = function () {
		return this.locked;
	};
	CellXfs.prototype.setLocked = function (val) {
		this.locked = val;
	};
	CellXfs.prototype.getHidden = function () {
		return this.hidden;
	};
	CellXfs.prototype.setHidden = function (val) {
		this.hidden = val;
	};
    CellXfs.prototype.getPivotButton = function () {
        return this.PivotButton;
    };
    CellXfs.prototype.setPivotButton = function (val) {
        this.PivotButton = val;
    };
    CellXfs.prototype.getXfId = function () {
        return this.XfId;
    };
    CellXfs.prototype.setXfId = function (val) {
        this.XfId = val;
    };
    CellXfs.prototype.getOperationCache = function (operation, val) {
        var res = undefined;
        operation = this.operationCache[operation];
        if (operation) {
            res = operation[val];
        }
        return res;
    };
    CellXfs.prototype.setOperationCache = function (operation, val, xfs) {
        var valCache = this.operationCache[operation];
        if (!valCache) {
            valCache = {};
            this.operationCache[operation] = valCache;
        }
        valCache[val] = xfs;
    };

	CellXfs.prototype.asc_getFillColor = function () {
		return Asc.colorObjToAscColor(this.getFill2().bg());
	};
	CellXfs.prototype.asc_getFill = function () {
		return this.getFill2().clone();
	};

    CellXfs.prototype.asc_getFontName = function () {
    	var name = this.getFont2().getName();
		var _name = AscFonts.g_fontApplication ? AscFonts.g_fontApplication.NameToInterface[name] : null;
		return _name ? _name : name;
	};
    CellXfs.prototype.asc_getFontSize = function () {
        return this.getFont2().getSize();
    };
    CellXfs.prototype.asc_getFontColor = function () {
        return Asc.colorObjToAscColor(this.getFont2().getColorNotDefault());
    };
    CellXfs.prototype.asc_getFontBold = function () {
        return this.getFont2().getBold();
    };
    CellXfs.prototype.asc_getFontItalic = function () {
        return this.getFont2().getItalic();
    };
    CellXfs.prototype.asc_getFontUnderline = function () {
        // ToDo убрать, когда будет реализовано двойное подчеркивание
        return (Asc.EUnderline.underlineNone !== this.getFont2().getUnderline());
    };
    CellXfs.prototype.asc_getFontStrikeout = function () {
        return this.getFont2().getStrikeout();
    };
    CellXfs.prototype.asc_getFontSubscript = function () {
        return (AscCommon.vertalign_SubScript === this.getFont2().getVerticalAlign());
    };
    CellXfs.prototype.asc_getFontSuperscript = function () {
        return (AscCommon.vertalign_SuperScript === this.getFont2().getVerticalAlign());
    };

	CellXfs.prototype.asc_getNumFormat = function () {
		return this.getNum2().getFormat();
	};
	CellXfs.prototype.asc_getNumFormatInfo = function () {
		return this.getNum2().getNumFormat().getTypeInfo();
	};

	CellXfs.prototype.asc_getHorAlign = function () {
		return this.getAlign2().getAlignHorizontal();
	};
	CellXfs.prototype.asc_getVertAlign = function () {
		return this.getAlign2().getAlignVertical();
	};
	CellXfs.prototype.asc_getAngle = function () {
		return this.getAlign2().getAngle();
	};
	CellXfs.prototype.asc_getIndent = function () {
		return this.getAlign2().getIndent();
	};
	CellXfs.prototype.asc_getWrapText = function () {
		var align = this.getAlign2();
		return align.getWrap() || align.hor === AscCommon.align_Distributed;
	};
	CellXfs.prototype.asc_getShrinkToFit = function () {
		return this.getAlign2().getShrinkToFit();
	};
	CellXfs.prototype.asc_getPreview = function (api, text, width, height) {
		return AscCommonExcel.generateXfsStyle(width, height, api.wb, this, text);
	};
	CellXfs.prototype.asc_getPreview2 = function (api, id, text) {
		if (this.num) {
			var oNumFormat = AscCommon.oNumFormatCache.get(this.num.getFormat());
			if (false == oNumFormat.isGeneralFormat()) {
				var aText = oNumFormat.format(previewConditionalFormattingNum);
				text = AscCommonExcel.getStringFromMultiText(aText);
			}
		}

		var oldAlign = this.align;
		var oldAlignHor;
		if (oldAlign) {
			oldAlignHor = this.align.hor;
		} else {
			this.align = new Align();
		}
		this.align.hor = AscCommon.align_Center;

		var res = AscCommonExcel.generateXfsStyle2(id, api.wb, this, text);
		if (oldAlign) {
			this.align.hor = oldAlignHor;
		} else {
			this.align = null;
		}

		return res;
	};
	CellXfs.prototype.asc_getLocked = function () {
		return this.locked === null || this.locked === true;
	};
	CellXfs.prototype.asc_getHidden = function () {
		return this.hidden;
	};

	CellXfs.prototype.asc_setFillColor = function (val) {
		//TODO не применяю цвета темы?
		var fill = null;
		if (val) {
			fill = new AscCommonExcel.Fill();
			fill.fromColor(AscCommonExcel.createRgbColor(val.asc_getR(),val.asc_getG(),val.asc_getB()));
		}
		return this.setFill(fill);
	};
	CellXfs.prototype.asc_setFill = function (val) {
		return this.setFill(val);
	};
	CellXfs.prototype.asc_setFontName = function (val) {
		if (!this.font) {
			this.font = new AscCommonExcel.Font();
		}
		this.getFont().setName(val);
	};
	CellXfs.prototype.asc_setFontSize = function (val) {
		if (!this.font) {
			this.font = new AscCommonExcel.Font();
		}
		this.getFont().setSize(val);
	};
	CellXfs.prototype.asc_setFontColor = function (val) {
		if (!this.font) {
			this.font = new AscCommonExcel.Font();
		}
		var color = AscCommonExcel.CorrectAscColor(val);
		this.getFont().setColor(color);
	};
	CellXfs.prototype.asc_setFontBold = function (val) {
		if (!this.font) {
			this.font = new AscCommonExcel.Font();
		}
		this.getFont().setBold(val);
	};
	CellXfs.prototype.asc_setFontItalic = function (val) {
		if (!this.font) {
			this.font = new AscCommonExcel.Font();
		}
		this.getFont().setItalic(val);
	};
	CellXfs.prototype.asc_setFontUnderline = function (val) {
		if (!this.font) {
			this.font = new AscCommonExcel.Font();
		}
		this.getFont().setUnderline(val ? Asc.EUnderline.underlineSingle : Asc.EUnderline.underlineNone);
	};
	CellXfs.prototype.asc_setFontStrikeout = function (val) {
		if (!this.font) {
			this.font = new AscCommonExcel.Font();
		}
		this.getFont().setStrikeout(val);
	};
	CellXfs.prototype.asc_setFontSubscript = function (val) {
		if (!this.font) {
			this.font = new AscCommonExcel.Font();
		}
		this.getFont().setVerticalAlign(val ? AscCommon.vertalign_SubScript : AscCommon.vertalign_Baseline);
	};
	CellXfs.prototype.asc_setFontSuperscript = function (val) {
		if (!this.font) {
			this.font = new AscCommonExcel.Font();
		}
		this.getFont().setVerticalAlign(val ? AscCommon.vertalign_SuperScript : AscCommon.vertalign_Baseline);
	};
	CellXfs.prototype.asc_setBorder = function (val) {
		if (val.length < 1) {
			this.border = null;
		}

		//TODO duplicate
		function makeBorder(b) {
			var border = new AscCommonExcel.BorderProp();
			if (b === false) {
				border.setStyle(c_oAscBorderStyles.None);
			} else if (b) {
				if (b.style !== null && b.style !== undefined) {
					border.setStyle(b.style);
				}
				if (b.color !== null && b.color !== undefined) {
					if (b.color instanceof Asc.asc_CColor) {
						border.c = AscCommonExcel.CorrectAscColor(b.color);
					}
				}
			}
			return border;
		}
		
		var res = new AscCommonExcel.Border();
		res.initDefault();
		var c_oAscBorderOptions = Asc.c_oAscBorderOptions;
		// Diagonal
		res.d = makeBorder(val[c_oAscBorderOptions.DiagD] || val[c_oAscBorderOptions.DiagU]);
		res.dd = !!val[c_oAscBorderOptions.DiagD];
		res.du = !!val[c_oAscBorderOptions.DiagU];
		// Vertical
		res.l = makeBorder(val[c_oAscBorderOptions.Left]);
		res.iv = makeBorder(val[c_oAscBorderOptions.InnerV]);
		res.r = makeBorder(val[c_oAscBorderOptions.Right]);
		// Horizontal
		res.t = makeBorder(val[c_oAscBorderOptions.Top]);
		res.ih = makeBorder(val[c_oAscBorderOptions.InnerH]);
		res.b = makeBorder(val[c_oAscBorderOptions.Bottom]);
		
		this.border = res;
	};

	CellXfs.prototype.asc_setNumFormatInfo = function (val) {
		this.num = new AscCommonExcel.Num({f:val});
	};

	var g_oAlignProperties = {
		hor: 0,
		indent: 1,
		RelativeIndent: 2,
		shrink: 3,
		angle: 4,
		ver: 5,
		wrap: 6
	};

	/** @constructor */
	function Align(val) {
		if (null == val) {
			val = g_oDefaultFormat.AlignAbs;
		}
		this.hor = val.hor;
		this.indent = val.indent;
		this.RelativeIndent = val.RelativeIndent;
		this.shrink = val.shrink;
		this.angle = val.angle;
		this.ver = val.ver;
		this.wrap = val.wrap;

		this._hash;
		this._index;
	}

	Align.prototype.Properties = g_oAlignProperties;
	Align.prototype.getHash = function () {
		if (!this._hash) {
			this._hash = this.hor + '|' + this.indent + '|' + this.RelativeIndent + '|' + this.shrink + '|' +
				this.angle + '|' + this.ver + '|' + this.wrap;
		}
		return this._hash;
	};
	Align.prototype.getIndexNumber = function () {
		return this._index;
	};
	Align.prototype.setIndexNumber = function (val) {
		this._index = val;
	};
	Align.prototype._mergeProperty = function (first, second, def) {
		if (def != first) {
			return first;
		} else {
			return second;
		}
	};
	Align.prototype.merge = function (align) {
		var defaultAlign = g_oDefaultFormat.Align;
		var oRes = new Align();
		oRes.hor = this._mergeProperty(this.hor, align.hor, defaultAlign.hor);
		oRes.indent = this._mergeProperty(this.indent, align.indent, defaultAlign.indent);
		oRes.RelativeIndent = this._mergeProperty(this.RelativeIndent, align.RelativeIndent, defaultAlign.RelativeIndent);
		oRes.shrink = this._mergeProperty(this.shrink, align.shrink, defaultAlign.shrink);
		oRes.angle = this._mergeProperty(this.angle, align.angle, defaultAlign.angle);
		oRes.ver = this._mergeProperty(this.ver, align.ver, defaultAlign.ver);
		oRes.wrap = this._mergeProperty(this.wrap, align.wrap, defaultAlign.wrap);
		return oRes;
	};
	Align.prototype.getDif = function (val) {
		var oRes = new Align(this);
		var bEmpty = true;
		if (this.hor == val.hor) {
			oRes.hor = null;
		} else {
			bEmpty = false;
		}
		if (this.indent == val.indent) {
			oRes.indent = null;
		} else {
			bEmpty = false;
		}
		if (this.RelativeIndent == val.RelativeIndent) {
			oRes.RelativeIndent = null;
		} else {
			bEmpty = false;
		}
		if (this.shrink == val.shrink) {
			oRes.shrink = null;
		} else {
			bEmpty = false;
		}
		if (this.angle == val.angle) {
			oRes.angle = null;
		} else {
			bEmpty = false;
		}
		if (this.ver == val.ver) {
			oRes.ver = null;
		} else {
			bEmpty = false;
		}
		if (this.wrap == val.wrap) {
			oRes.wrap = null;
		} else {
			bEmpty = false;
		}
		if (bEmpty) {
			oRes = null;
		}
		return oRes;
	};
	Align.prototype.isEqual = function (val) {
		return this.hor == val.hor && this.indent == val.indent && this.RelativeIndent == val.RelativeIndent && this.shrink == val.shrink &&
			this.angle == val.angle && this.ver == val.ver && this.wrap == val.wrap;
	};
	Align.prototype.clone = function () {
		return new Align(this);
	};
	Align.prototype.getType = function () {
		return UndoRedoDataTypes.StyleAlign;
	};
	Align.prototype.getProperties = function () {
		return this.Properties;
	};
	Align.prototype.getProperty = function (nType) {
		switch (nType) {
			case this.Properties.hor:
				return this.hor;
				break;
			case this.Properties.indent:
				return this.indent;
				break;
			case this.Properties.RelativeIndent:
				return this.RelativeIndent;
				break;
			case this.Properties.shrink:
				return this.shrink;
				break;
			case this.Properties.angle:
				return this.angle;
				break;
			case this.Properties.ver:
				return this.ver;
				break;
			case this.Properties.wrap:
				return this.wrap;
				break;
		}
	};
	Align.prototype.setProperty = function (nType, value) {
		switch (nType) {
			case this.Properties.hor:
				this.hor = value;
				break;
			case this.Properties.indent:
				this.indent = value;
				break;
			case this.Properties.RelativeIndent:
				this.RelativeIndent = value;
				break;
			case this.Properties.shrink:
				this.shrink = value;
				break;
			case this.Properties.angle:
				this.angle = value;
				break;
			case this.Properties.ver:
				this.ver = value;
				break;
			case this.Properties.wrap:
				this.wrap = value;
				break;
		}
	};
	Align.prototype.getAngle = function () {
		var nRes = 0;
		if (0 <= this.angle && this.angle <= 180) {
			nRes = this.angle <= 90 ? this.angle : 90 - this.angle;
		} else if (this.angle === AscCommonExcel.g_nVerticalTextAngle) {
			nRes = this.angle;
		}
		return nRes;
	};
	Align.prototype.setAngle = function (val) {
		this.angle = null !== val ? AscCommonExcel.angleInterfaceToFormat(val) : val;
	};
	Align.prototype.getWrap = function () {
		// Для justify wrap всегда true
		return (AscCommon.align_Justify === this.hor || Asc.c_oAscVAlign.Just === this.ver || Asc.c_oAscVAlign.Dist === this.ver) ? true : this.wrap;
	};
	Align.prototype.setWrap = function (val) {
		this.wrap = val;
	};
	Align.prototype.getShrinkToFit = function () {
		return this.shrink;
	};
	Align.prototype.setShrinkToFit = function (val) {
		this.shrink = val;
	};
	Align.prototype.getAlignHorizontal = function () {
		return this.hor;
	};
	Align.prototype.setAlignHorizontal = function (val) {
		this.hor = val;
	};
	Align.prototype.getAlignVertical = function () {
		return this.ver;
	};
	Align.prototype.setAlignVertical = function (val) {
		this.ver = val;
	};
	Align.prototype.getIndent = function () {
		return this.indent;
	};
	Align.prototype.setIndent = function (val) {
		this.indent = val;
	};
	Align.prototype.readAttributes = function (attr, uq) {
		if (attr()) {
			var vals = attr();
			var val;
			val = vals["horizontal"];
			if (undefined !== val) {
				val = AscCommonExcel.FromXml_ST_HorizontalAlignment(val);
				if (-1 !== val) {
					this.hor = val;
				}
			}
			val = vals["vertical"];
			if (undefined !== val) {
				val = AscCommonExcel.FromXml_ST_VerticalAlignment(val);
				if (-1 !== val) {
					this.ver = val;
				}
			}
			val = vals["textRotation"];
			if (undefined !== val) {
				this.angle = val - 0;
			}
			val = vals["wrapText"];
			if (undefined !== val) {
				this.wrap = AscCommon.getBoolFromXml(val);
			}
			val = vals["indent"];
			if (undefined !== val) {
				this.indent = val - 0;
			}
			val = vals["relativeIndent"];
			if (undefined !== val) {
				this.RelativeIndent = val - 0;
			}
			val = vals["shrinkToFit"];
			if (undefined !== val) {
				this.shrink = AscCommon.getBoolFromXml(val);
			}
		}
	};

/** @constructor */
function CCellStyles() {
	this.CustomStyles = [];
	this.DefaultStyles = [];
	// ToDo нужно все компоновать в общий список стилей (для того, чтобы не было проблем с добавлением стилей и отсутствия имени стиля)
	this.AllStyles = {};
}
CCellStyles.prototype.generateFontMap = function (oFontMap) {
	this._generateFontMap(oFontMap, this.DefaultStyles);
	this._generateFontMap(oFontMap, this.CustomStyles);
};
CCellStyles.prototype._generateFontMap = function (oFontMap, aStyles) {
	var i, length, oStyle;
	for (i = 0, length = aStyles.length; i < length; ++i) {
		oStyle = aStyles[i];
		if (null != oStyle.xfs && null != oStyle.xfs.font)
			oFontMap[oStyle.xfs.font.getName()] = 1;
	}
};
/**
 * Возвращает колличество стилей без учета скрытых
 */
CCellStyles.prototype.getDefaultStylesCount = function () {
	var nCount = this.DefaultStyles.length;
	for (var i = 0, length = nCount; i < length; ++i) {
		if (this.DefaultStyles[i].Hidden)
			--nCount;
	}
	return nCount;
};
/**
 * Возвращает колличество стилей без учета скрытых и стандартных
 */
CCellStyles.prototype.getCustomStylesCount = function () {
	var nCount = this.CustomStyles.length;
	for (var i = 0, length = nCount; i < length; ++i) {
		if (this.CustomStyles[i].Hidden || null != this.CustomStyles[i].BuiltinId)
			--nCount;
	}
	return nCount;
};
CCellStyles.prototype.getStyleByXfId = function (oXfId) {
	for (var i = 0, length = this.CustomStyles.length; i < length; ++i) {
		if (oXfId === this.CustomStyles[i].XfId) {
			return this.CustomStyles[i];
		}
	}

	return null;
};
CCellStyles.prototype.getStyleNameByXfId = function (oXfId) {
	var styleName = null;
	if (null === oXfId)
		return styleName;

	var style = null;
	for (var i = 0, length = this.CustomStyles.length; i < length; ++i) {
		style = this.CustomStyles[i];
		if (oXfId === style.XfId) {
			if (null !== style.BuiltinId) {
				styleName = this.getDefaultStyleNameByBuiltinId(style.BuiltinId);
				if (null === styleName)
					styleName = style.Name;
				break;
			} else {
				styleName = style.Name;
				break;
			}
		}
	}

	return styleName;
};
CCellStyles.prototype.getDefaultStyleNameByBuiltinId = function (oBuiltinId) {
	var style = null;
	for (var i = 0, length = this.DefaultStyles.length; i < length; ++i) {
		style = this.DefaultStyles[i];
		if (style.BuiltinId === oBuiltinId)
			return style.Name;
	}
	return null;
};
CCellStyles.prototype.getCustomStyleByBuiltinId = function (oBuiltinId) {
	var style = null;
	for (var i = 0, length = this.CustomStyles.length; i < length; ++i) {
		style = this.CustomStyles[i];
		if (style.BuiltinId === oBuiltinId)
			return style;
	}
	return null;
};
CCellStyles.prototype._prepareCellStyle = function (name) {
	var defaultStyle = null;
	var style = null;
	var i, length;
	var maxXfId = -1;
	// Проверим, есть ли в default
	for (i = 0, length = this.DefaultStyles.length; i < length; ++i) {
		if (name === this.DefaultStyles[i].Name) {
			defaultStyle = this.DefaultStyles[i];
			break;
		}
	}
	// Если есть в default, ищем в custom по builtinId. Если нет, то по имени
	if (defaultStyle) {
		for (i = 0, length = this.CustomStyles.length; i < length; ++i) {
			if (defaultStyle.BuiltinId === this.CustomStyles[i].BuiltinId) {
				style = this.CustomStyles[i];
				break;
			}
			maxXfId = Math.max(maxXfId, this.CustomStyles[i].XfId);
		}
	} else {
		for (i = 0, length = this.CustomStyles.length; i < length; ++i) {
			if (name === this.CustomStyles[i].Name) {
				style = this.CustomStyles[i];
				break;
			}
			maxXfId = Math.max(maxXfId, this.CustomStyles[i].XfId);
		}
	}

	// Если нашли, то возвращаем XfId
	if (style)
		return style.XfId;

	if (defaultStyle) {
		//todo add to history. it allows save XfId in history
		this.CustomStyles[i] = defaultStyle.clone();
		this.CustomStyles[i].XfId = ++maxXfId;
		return this.CustomStyles[i].XfId;
	}
	return g_oDefaultFormat.XfId;
};
/** @constructor */
function CCellStyle() {
	this.BuiltinId = null;
	this.CustomBuiltin = null;
	this.Hidden = null;
	this.ILevel = null;
	this.Name = null;
	this.XfId = null;

	this.xfs = null;

	this.ApplyBorder = true;
	this.ApplyFill = true;
	this.ApplyFont = true;
	this.ApplyNumberFormat = true;
}
CCellStyle.prototype.clone = function () {
	var oNewStyle = new CCellStyle();
	oNewStyle.BuiltinId = this.BuiltinId;
	oNewStyle.CustomBuiltin = this.CustomBuiltin;
	oNewStyle.Hidden = this.Hidden;
	oNewStyle.ILevel = this.ILevel;
	oNewStyle.Name = this.Name;

	oNewStyle.ApplyBorder = this.ApplyBorder;
	oNewStyle.ApplyFill = this.ApplyFill;
	oNewStyle.ApplyFont = this.ApplyFont;
	oNewStyle.ApplyNumberFormat = this.ApplyNumberFormat;

	oNewStyle.xfs = this.xfs.clone();
	return oNewStyle;
};
CCellStyle.prototype.getFill = function () {
	if (null != this.xfs && null != this.xfs.fill)
		return this.xfs.fill;

	return g_oDefaultFormat.Fill;
};
CCellStyle.prototype.getFontColor = function () {
	if (null != this.xfs && null != this.xfs.font)
		return this.xfs.font.getColor();

	return g_oDefaultFormat.Font.c;
};
CCellStyle.prototype.getFont = function () {
	if (null != this.xfs && null != this.xfs.font)
		return this.xfs.font;
	return g_oDefaultFormat.Font;
};
CCellStyle.prototype.getBorder = function () {
	if (null != this.xfs && null != this.xfs.border)
		return this.xfs.border;
	return g_oDefaultFormat.Border;
};
CCellStyle.prototype.getNumFormatStr = function () {
	if(null != this.xfs && null != this.xfs.num)
		return this.xfs.num.getFormat();
	return g_oDefaultFormat.Num.getFormat();
};
/** @constructor */
function StyleManager(){
}
StyleManager.prototype =
{
	init: function(firstXf, firstFont, firstFill, secondFill, firstBorder, normalXf) {
		g_StyleCache.firstXf = firstXf;
		g_StyleCache.firstFont = firstFont;
		g_StyleCache.firstFill = firstFill;
		g_StyleCache.secondFill = secondFill;
		g_StyleCache.firstBorder = firstBorder;
		g_StyleCache.normalXf = normalXf;
		if(null != firstXf.font)
			g_oDefaultFormat.xfs.font = g_oDefaultFormat.Font = firstXf.font;
		if(null != firstXf.fill)
			g_oDefaultFormat.xfs.fill = g_oDefaultFormat.Fill = firstXf.fill.clone();
		if(null != firstXf.border)
			g_oDefaultFormat.xfs.border = g_oDefaultFormat.Border = firstXf.border.clone();
		if(null != firstXf.num)
			g_oDefaultFormat.xfs.num = g_oDefaultFormat.Num = firstXf.num.clone();
		if(null != firstXf.align)
			g_oDefaultFormat.xfs.align = g_oDefaultFormat.Align = firstXf.align.clone();
		if (null !== firstXf.XfId) {
			g_oDefaultFormat.XfId = firstXf.XfId;
		}
	},
	setCellStyle : function(oItemWithXfs, val)
	{
		return this._setProperty(oItemWithXfs, val, "XfId", CellXfs.prototype.getXfId, CellXfs.prototype.setXfId);
	},
	setNum : function(oItemWithXfs, val)
	{
		if (val && val.f) {
			var numFormat = AscCommon.oNumFormatCache.get(val.f);
			numFormat.checkCultureInfoFontPicker();
		}
		return this._setProperty(oItemWithXfs, val, "num", CellXfs.prototype.getNum, CellXfs.prototype.setNum, g_StyleCache.addNum);
	},
	setFont : function(oItemWithXfs, val)
	{
		return this._setProperty(oItemWithXfs, val, "font", CellXfs.prototype.getFont, CellXfs.prototype.setFont, g_StyleCache.addFont);
	},
	setFill : function(oItemWithXfs, val)
	{
		val && val.checkEmptyContent();
		return this._setProperty(oItemWithXfs, val, "fill", CellXfs.prototype.getFill, CellXfs.prototype.setFill, g_StyleCache.addFill);
	},
	setBorder : function(oItemWithXfs, val)
	{
		return this._setProperty(oItemWithXfs, val, "border", CellXfs.prototype.getBorder, CellXfs.prototype.setBorder, g_StyleCache.addBorder);
	},
	setQuotePrefix : function(oItemWithXfs, val)
	{
		return this._setProperty(oItemWithXfs, val, "quotePrefix", CellXfs.prototype.getQuotePrefix, CellXfs.prototype.setQuotePrefix);
	},
	setApplyProtection : function(oItemWithXfs, val)
	{
		return this._setProperty(oItemWithXfs, val, "applyProtection", CellXfs.prototype.getApplyProtection, CellXfs.prototype.setApplyProtection);
	},
	setLocked : function(oItemWithXfs, val)
	{
		return this._setProperty(oItemWithXfs, val, "locked", CellXfs.prototype.getLocked, CellXfs.prototype.setLocked);
	},
	setHiddenFormulas : function(oItemWithXfs, val)
	{
		return this._setProperty(oItemWithXfs, val, "hidden", CellXfs.prototype.getHidden, CellXfs.prototype.setHidden);
	},
	setPivotButton : function(oItemWithXfs, val)
	{
		return this._setProperty(oItemWithXfs, val, "pivotButton", CellXfs.prototype.getPivotButton, CellXfs.prototype.setPivotButton);
	},
	setFontname : function(oItemWithXfs, val)
	{
		return this._setFontProperty(oItemWithXfs, val, "name", Font.prototype.getName, function(val){
			this.setName(val);
			this.setScheme(null);
		}, "name");
	},
	setFontsize : function(oItemWithXfs, val)
	{
		return this._setFontProperty(oItemWithXfs, val, "size", Font.prototype.getSize, Font.prototype.setSize);
	},
	setFontcolor : function(oItemWithXfs, val)
	{
		return this._setFontProperty(oItemWithXfs, val, "color", Font.prototype.getColor, Font.prototype.setColor);
	},
	setBold : function(oItemWithXfs, val)
	{
		return this._setFontProperty(oItemWithXfs, val, "bold", Font.prototype.getBold, Font.prototype.setBold);
	},
	setItalic : function(oItemWithXfs, val)
	{
		return this._setFontProperty(oItemWithXfs, val, "italic", Font.prototype.getItalic, Font.prototype.setItalic);
	},
	setUnderline : function(oItemWithXfs, val)
	{
		return this._setFontProperty(oItemWithXfs, val, "underline", Font.prototype.getUnderline, Font.prototype.setUnderline);
	},
	setStrikeout : function(oItemWithXfs, val)
	{
		return this._setFontProperty(oItemWithXfs, val, "strikeout", Font.prototype.getStrikeout, Font.prototype.setStrikeout);
	},
	setFontAlign : function(oItemWithXfs, val)
	{
		return this._setFontProperty(oItemWithXfs, val, "fontAlign", Font.prototype.getVerticalAlign, Font.prototype.setVerticalAlign);
	},
	setAlignVertical : function(oItemWithXfs, val)
	{
		return this._setAlignProperty(oItemWithXfs, val, "alignVertical", Align.prototype.getAlignVertical, Align.prototype.setAlignVertical);
	},
	setAlignHorizontal : function(oItemWithXfs, val)
	{
		return this._setAlignProperty(oItemWithXfs, val, "alignHorizontal", Align.prototype.getAlignHorizontal, Align.prototype.setAlignHorizontal);
	},
	setShrinkToFit : function(oItemWithXfs, val)
	{
		return this._setAlignProperty(oItemWithXfs, val, "shrinkToFit", Align.prototype.getShrinkToFit, Align.prototype.setShrinkToFit);
	},
	setWrap : function(oItemWithXfs, val)
	{
		return this._setAlignProperty(oItemWithXfs, val, "wrap", Align.prototype.getWrap, Align.prototype.setWrap);
	},
    setAngle : function(oItemWithXfs, val)
    {
		return this._setAlignProperty(oItemWithXfs, val, "angle", function(){
			return AscCommonExcel.angleFormatToInterface2(this.angle);
		}, Align.prototype.setAngle);
    },
	setIndent : function(oItemWithXfs, val)
	{
		return this._setAlignProperty(oItemWithXfs, val, "indent", Align.prototype.getIndent, Align.prototype.setIndent);
	},
	_initXf: function(oItemWithXfs){
		var xfs = oItemWithXfs.xfs;
		if (!xfs) {
			if (oItemWithXfs.getDefaultXfs) {
				xfs = oItemWithXfs.getDefaultXfs();
			}
			if (!xfs) {
				xfs = g_StyleCache.firstXf;
			}
		}
		return xfs;
	},
	_initXfFont: function(xfs){
		xfs = xfs.clone();
		if(null == xfs.font){
			xfs.font = g_oDefaultFormat.Font;
		}
		xfs.font = xfs.font.clone();
		return xfs;
	},
	_initXfAlign: function(xfs){
		xfs = xfs.clone();
		if(null == xfs.align){
			xfs.align = g_oDefaultFormat.AlignAbs;
		}
		xfs.align = xfs.align.clone();
		return xfs;
	},
	_setProperty: function(oItemWithXfs, val, prop, getFunc, setFunc, addFunc) {
		var xfs = oItemWithXfs.xfs;
		var oRes = {newVal: null, oldVal: xfs ? getFunc.call(xfs) : null};
		xfs = this._initXf(oItemWithXfs);
		var hash = val && val.getHash ? val.getHash() : val;
		var xfsOperationCache = xfs;
		var newXf = xfs.getOperationCache(prop, hash);
		if (newXf) {
			oItemWithXfs.setStyleInternal(newXf);
			xfs = newXf;
		} else {
			xfs = xfs.clone();
			if (null != val) {
				if (addFunc) {
					setFunc.call(xfs, addFunc.call(g_StyleCache, val));
				} else {
					setFunc.call(xfs, val);
				}
			}
			else if (null != xfs) {
				setFunc.call(xfs, null);
			}

			xfs = g_StyleCache.addXf(xfs);
			xfsOperationCache.setOperationCache(prop, hash, xfs);
			oItemWithXfs.setStyleInternal(xfs);
		}
		oRes.newVal = xfs ? getFunc.call(xfs) : null;
		return oRes;
	},
	_setFontProperty : function(oItemWithXfs, val, prop, getFunc, setFunc)
	{
		var xfs = oItemWithXfs.xfs;
		var oRes = {newVal: val, oldVal: xfs && xfs.font ? getFunc.call(xfs.font): null};
		xfs = this._initXf(oItemWithXfs);
		var hash = val && val.getHash ? val.getHash() : val;
		var xfsOperationCache = xfs;
		var newXf = xfs.getOperationCache(prop, hash);
		if (newXf) {
			oItemWithXfs.setStyleInternal(newXf);
		} else {
			xfs = this._initXfFont(xfs);

			setFunc.call(xfs.font, val);
			xfs.font = g_StyleCache.addFont(xfs.font);

			xfs = g_StyleCache.addXf(xfs);
			xfsOperationCache.setOperationCache(prop, val, xfs);
			oItemWithXfs.setStyleInternal(xfs);
		}
		return oRes;
	},
	_setAlignProperty : function(oItemWithXfs, val, prop, getFunc, setFunc)
	{
		var xfs = oItemWithXfs.xfs;
		let oldAlign = xfs ? (xfs.align || g_oDefaultFormat.AlignAbs) : g_oDefaultFormat.Align;
		var oRes = {newVal: val, oldVal: getFunc.call(oldAlign)};
		xfs = this._initXf(oItemWithXfs);
		var xfsOperationCache = xfs;
		var newXf = xfs.getOperationCache(prop, val);
		if (newXf) {
			oItemWithXfs.setStyleInternal(newXf);
		} else {
			xfs = this._initXfAlign(xfs);

			setFunc.call(xfs.align, val);
			xfs.align = g_StyleCache.addAlign(xfs.align);

			xfs = g_StyleCache.addXf(xfs);
			xfsOperationCache.setOperationCache(prop, val, xfs);
			oItemWithXfs.setStyleInternal(xfs);
		}
		return oRes;
	}
};

	function StyleCache() {
		this.Clear();
	}
	StyleCache.prototype.Clear = function() {
		this.fonts = {count: 0, vals: {}};
		this.fills = {count: 0, vals: {}};
		this.borders = {count: 0, vals: {}};
		this.nums = {count: 0, vals: {}};
		this.aligns = {count: 0, vals: {}};
		this.xfs = {list: [], vals: {}};
		this.firstXf =  new CellXfs();
		this.firstFont = null;
		this.firstFill = null;
		this.secondFill = null;
		this.firstBorder = null;
		this.normalXf =  new CellXfs();
	};
	StyleCache.prototype.addFont = function(newFont) {
		return this._add(this.fonts, newFont);
	};
	StyleCache.prototype.addFill = function(newFill, forceAdd) {
		return this._add(this.fills, newFill, forceAdd);
	};
	StyleCache.prototype.addBorder = function(newBorder, forceAdd) {
		return this._add(this.borders, newBorder, forceAdd);
	};
	StyleCache.prototype.addNum = function(newNum) {
		return this._add(this.nums, newNum);
	};
	StyleCache.prototype.addAlign = function(newAlign) {
		return this._add(this.aligns, newAlign);
	};
	StyleCache.prototype.addXf = function(newXf, forceAdd) {
		if (newXf) {
			if(newXf.font){
				newXf.font = this.addFont(newXf.font);
			}
			if(newXf.fill){
				newXf.fill = this.addFill(newXf.fill);
			}
			if(newXf.border){
				newXf.border = this.addBorder(newXf.border);
			}
			if(newXf.num){
				newXf.num = this.addNum(newXf.num);
			}
			if(newXf.align){
				newXf.align = this.addAlign(newXf.align);
			}
		}
		return this._add(this.xfs, newXf, forceAdd);
	};
	StyleCache.prototype.getXf = function(index) {
		return 1 <= index && index <= this.xfs.list.length ? this.xfs.list[index - 1] : null;
	};
	StyleCache.prototype.getXfCount = function() {
		return this.xfs.list.length;
	};
	StyleCache.prototype.getNumFormatStrings = function() {
		var res = [];
		for(var fmt in this.nums.vals){
			res.push(this.nums.vals[fmt].getFormat());
		}
		return res;
	};
	StyleCache.prototype._add = function(container, newVal, forceAdd) {
		if (newVal && undefined === newVal.getIndexNumber()) {
			var hash = newVal.getHash();
			var res = container.vals[hash];
			if (!res || forceAdd) {
				if (container.list) {
					//index starts with 1
					newVal.setIndexNumber(container.list.push(newVal));
				} else {
					newVal.setIndexNumber(container.count++);
				}
				if (!res || forceAdd) {
					container.vals[hash] = newVal;
				}
				res = newVal;
			}
			return res;
		} else {
			return newVal;
		}
	};
	var g_StyleCache = new StyleCache();

	/** @constructor */
	function SheetMergedStyles() {
		this.stylesTablePivot = [];
		this.stylesConditional = {};
		this.stylesConditionalIterator = null;
	}

	SheetMergedStyles.prototype.setTablePivotStyle = function(range, xf, stripe) {
		this.stylesTablePivot.push({xf: xf, range: range, stripe: stripe, borders: undefined});
	};
	SheetMergedStyles.prototype.clearTablePivotStyle = function(range) {
		for (var i = this.stylesTablePivot.length - 1; i >= 0; --i) {
			var style = this.stylesTablePivot[i];
			if (style.range.isIntersect(range)) {
				this.stylesTablePivot.splice(i, 1);
			}
		}
	};
	SheetMergedStyles.prototype.setConditionalStyle = function(rule, ranges, formula) {
		this.stylesConditionalIterator = null;
		this.stylesConditional[rule.Get_Id()] = {ranges: ranges, formula: formula, rule: rule};
	};
	SheetMergedStyles.prototype.clearConditionalStyle = function(multiplyRange) {
		this.stylesConditionalIterator = null;
		for (var i in this.stylesConditional) {
			if (this.stylesConditional.hasOwnProperty(i)) {
				var style = this.stylesConditional[i];
				var mr = new AscCommonExcel.MultiplyRange(style.ranges);
				if (mr.isIntersect(multiplyRange)) {
					delete this.stylesConditional[i];
				}
			}
		}
	};
	SheetMergedStyles.prototype.getStyle = function(hiddenManager, row, col, opt_ws) {
		var res = {table: [], conditional: []};
		if (opt_ws) {
			opt_ws._updateConditionalFormatting();
		}
		if (!this.stylesConditionalIterator) {
			this.stylesConditionalIterator = new AscCommon.RangeTopBottomIterator();
			//todo lose stylesConditional sorting
			this.stylesConditionalIterator.init(Object.values(this.stylesConditional), function(elem) {
				return elem.ranges;
			});
		}
		var rules = this.stylesConditionalIterator.get(row, col);
		//todo sort inside RangeTopBottomIterator ?
		rules.sort(function(v1, v2) {
			return v2.rule.priority - v1.rule.priority;
		});
		for (var i = 0; i < rules.length; ++i) {
			var xf = rules[i].formula(row, col);
			if (xf) {
				res.conditional.push(xf);
			}
		}
		for (var i = 0; i < this.stylesTablePivot.length; ++i) {
			var style = this.stylesTablePivot[i];
			var borderIndex;
			var xf = style.xf;
			if (style.range.contains(col, row) && (borderIndex = this._getBorderIndex(hiddenManager, style.range, style.stripe, row, col, xf)) >= 0) {
				if (borderIndex > 0) {
					if (!style.borders) {
						style.borders = {};
					}
					var xfModified = style.borders[borderIndex];
					if (!xfModified) {
						xfModified = xf.clone();
						var borderModified = xfModified.border.clone();
						if (0 !== (1 & borderIndex)) {
							borderModified.l = borderModified.iv;
						}
						if (0 !== (2 & borderIndex)) {
							borderModified.t = borderModified.ih;
						}
						if (0 !== (4 & borderIndex)) {
							borderModified.r = borderModified.iv;
						}
						if (0 !== (8 & borderIndex)) {
							borderModified.b = borderModified.ih;
						}
						if (0 !== (16 & borderIndex)) {
							borderModified.du = false;
							borderModified.dd = false;
						}
						xfModified.border = g_StyleCache.addBorder(borderModified);
						xfModified = g_StyleCache.addXf(xfModified);
						style.borders[borderIndex] = xfModified;
					}
					xf = xfModified;
				}
				res.table.push(xf);
			}
		}
		return res;
	};
	SheetMergedStyles.prototype._getBorderIndex = function(hiddenManager, bbox, stripe, row, col, xf) {
		var borderIndex = 0;
		var hidden;
		if (stripe) {
			if (stripe.row) {
				hidden = hiddenManager.getHiddenRowsCount(bbox.r1, row);
				var rowIndex = (row - bbox.r1 - hidden) % (stripe.size + stripe.offset);
				if (rowIndex < stripe.size) {
					if (xf.border) {
						if (bbox.c1 !== col && (xf.border.l || xf.border.iv)) {
							borderIndex += 1;
						}
						if (0 != rowIndex && (xf.border.t || xf.border.ih)) {
							borderIndex += 2;
						}
						if (bbox.c2 !== col && (xf.border.r || xf.border.iv)) {
							borderIndex += 4;
						}
						if (stripe.size - 1 != rowIndex && (xf.border.b || xf.border.ih)) {
							borderIndex += 8;
						}
					}
				} else {
					borderIndex = -1;
				}
			} else {
				hidden = hiddenManager.getHiddenColsCount(bbox.c1, col);
				var colIndex = (col - bbox.c1 - hidden) % (stripe.size + stripe.offset);
				if (colIndex < stripe.size) {
					if (xf.border) {
						if (0 != colIndex && (xf.border.l || xf.border.iv)) {
							borderIndex += 1;
						}
						if (bbox.r1 !== row && (xf.border.t || xf.border.ih)) {
							borderIndex += 2;
						}
						if (stripe.size - 1 != colIndex && (xf.border.r || xf.border.iv)) {
							borderIndex += 4;
						}
						if (bbox.r2 !== row && (xf.border.b || xf.border.ih)) {
							borderIndex += 8;
						}
					}
				} else {
					borderIndex = -1;
				}
			}
		} else if (xf.border) {
			if (bbox.c1 !== col && (xf.border.l || xf.border.iv)) {
				borderIndex += 1;
			}
			if (bbox.r1 !== row && (xf.border.t || xf.border.ih)) {
				borderIndex += 2;
			}
			if (bbox.c2 !== col && (xf.border.r || xf.border.iv)) {
				borderIndex += 4;
			}
			if (bbox.r2 !== row && (xf.border.b || xf.border.ih)) {
				borderIndex += 8;
			}
		}
		if (xf.border && (xf.border.du || xf.border.dd)) {
			borderIndex += 16;
		}
		return borderIndex;
	};

	var g_oHyperlinkProperties = {
		Ref: 0,
		Location: 1,
		Hyperlink: 2,
		Tooltip: 3
	};
	/** @constructor */
	function Hyperlink() {
		this.Properties = g_oHyperlinkProperties;
		this.Ref = null;
		this.Hyperlink = null;
		this.Tooltip = null;
		// Составные части Location
		this.Location = null;
		this.LocationSheet = null;
		this.LocationRange = null;
		this.LocationRangeBbox = null;
		this.bUpdateLocation = false;

		this.bVisited = false;

		this.bHyperlinkFunction = null;
	}

	Hyperlink.prototype.clone = function (oNewWs) {
		var oNewHyp = new Hyperlink();
		if (null !== this.Ref)
			oNewHyp.Ref = this.Ref.clone(oNewWs);
		if (null !== this.getLocation())
			oNewHyp.setLocation(this.getLocation());
		if (null !== this.LocationSheet)
			oNewHyp.LocationSheet = this.LocationSheet;
		if (null !== this.LocationRange)
			oNewHyp.LocationRange = this.LocationRange;
		if (null !== this.LocationRangeBbox)
			oNewHyp.LocationRangeBbox = this.LocationRangeBbox.clone();
		if (null !== this.Hyperlink)
			oNewHyp.Hyperlink = this.Hyperlink;
		if (null !== this.Tooltip)
			oNewHyp.Tooltip = this.Tooltip;
		if (null !== this.bVisited)
			oNewHyp.bVisited = this.bVisited;
		return oNewHyp;
	};
	Hyperlink.prototype.isEqual = function (obj) {
		var bRes = (this.getLocation() == obj.getLocation() && this.Hyperlink == obj.Hyperlink && this.Tooltip == obj.Tooltip);
		if (bRes) {
			var oBBoxRef = this.Ref.getBBox0();
			var oBBoxObj = obj.Ref.getBBox0();
			bRes = (oBBoxRef.r1 == oBBoxObj.r1 && oBBoxRef.c1 == oBBoxObj.c1 && oBBoxRef.r2 == oBBoxObj.r2 && oBBoxRef.c2 == oBBoxObj.c2);
		}
		return bRes;
	};
	Hyperlink.prototype.isValid = function () {
		var isValidLength = !this.Hyperlink || (this.Hyperlink && AscCommonExcel.getFullHyperlinkLength(this.Hyperlink) <= Asc.c_nMaxHyperlinkLength);
		return null != this.Ref && (null != this.getLocation() || null != this.Hyperlink) && isValidLength;
	};
	Hyperlink.prototype.setLocationSheet = function (LocationSheet) {
		this.LocationSheet = LocationSheet;
		this.bUpdateLocation = true;
	};
	Hyperlink.prototype.setLocationRange = function (LocationRange) {
		this.LocationRange = LocationRange;
		this.LocationRangeBbox = null;
		this.bUpdateLocation = true;
	};
	Hyperlink.prototype.setLocation = function (Location) {
		this.bUpdateLocation = true;
		this.LocationSheet = this.LocationRange = this.LocationRangeBbox = null;

		if (null !== Location) {
			if (parserHelp.isName3D(Location, 0) || parserHelp.isName(Location, 0)) {
				this.LocationRange = Location;
			} else {
				var result = parserHelp.parse3DRef(Location);
				if (!result) {
					// Can be in all mods. Excel bug...
					AscCommonExcel.executeInR1C1Mode(!AscCommonExcel.g_R1C1Mode, function () {
						result = parserHelp.parse3DRef(Location);
					});
				}
				if (null !== result) {
					this.LocationSheet = result.sheet;
					this.LocationRange = result.range;
				}
			}
		}
		this._updateLocation();
	};
	Hyperlink.prototype.getLocation = function () {
		if (this.bUpdateLocation)
			this._updateLocation();
		return this.Location;
	};
	Hyperlink.prototype.getLocationRange = function () {
		return this.LocationRangeBbox && this.LocationRangeBbox.getName(AscCommonExcel.g_R1C1Mode ?
			AscCommonExcel.referenceType.A : AscCommonExcel.referenceType.R);
	};
	Hyperlink.prototype._updateLocation = function () {
		var t = this;
		this.Location = null;
		this.bUpdateLocation = false;
		if (null !== this.LocationSheet && null !== this.LocationRange) {
			this.LocationRangeBbox = AscCommonExcel.g_oRangeCache.getAscRange(this.LocationRange);
			if (!this.LocationRangeBbox) {
				// Can be in all mods. Excel bug...
				AscCommonExcel.executeInR1C1Mode(!AscCommonExcel.g_R1C1Mode, function () {
					t.LocationRangeBbox = AscCommonExcel.g_oRangeCache.getAscRange(t.LocationRange);
				});
			}
			if (this.LocationRangeBbox) {
				AscCommonExcel.executeInR1C1Mode(false, function () {
					t.LocationRange = t.LocationRangeBbox.getName(AscCommonExcel.referenceType.R);
				});
				this.Location = parserHelp.get3DRef(this.LocationSheet, this.LocationRange);
			}
		} else if (null !== this.LocationRange) {
			this.Location = this.LocationRange;
		}
	};
	Hyperlink.prototype.setVisited = function (bVisited) {
		this.bVisited = bVisited;
	};
	Hyperlink.prototype.getVisited = function () {
		return this.bVisited;
	};
	Hyperlink.prototype.getHyperlinkType = function () {
		return null !== this.Hyperlink ? Asc.c_oAscHyperlinkType.WebLink : Asc.c_oAscHyperlinkType.RangeLink;

		//for local file
		/*var res = null;
		if (null !== this.Hyperlink) {
			//либо гиперссылка, либо ссылка на локальный файл(отдельное поле не стал заводить, все будет в Hyperlink)
			if (XRegExp.exec(this.Hyperlink, new XRegExp("([a-zA-Z]:)?(\\\\[a-zA-Z0-9_.-]+)+\\\\?"))) {
				res = Asc.c_oAscHyperlinkType.FileLink;
			} else {
				res = Asc.c_oAscHyperlinkType.WebLink;
			}
		}
		return null !== res ? res : Asc.c_oAscHyperlinkType.RangeLink;*/
	};
	Hyperlink.prototype.getType = function () {
		return UndoRedoDataTypes.Hyperlink;
	};
	Hyperlink.prototype.getProperties = function () {
		return this.Properties;
	};
	Hyperlink.prototype.getProperty = function (nType) {
		switch (nType) {
			case this.Properties.Ref: return parserHelp.get3DRef(this.Ref.worksheet.getName(), this.Ref.getName());
			case this.Properties.Location: return this.getLocation();
			case this.Properties.Hyperlink: return this.Hyperlink;
			case this.Properties.Tooltip: return this.Tooltip;
		}
	};
	Hyperlink.prototype.setProperty = function (nType, value) {
		switch (nType) {
			case this.Properties.Ref:
				//todo обработать нули
				var oRefParsed = parserHelp.parse3DRef(value);
				if (null !== oRefParsed) {
					// Получаем sheet по имени
					var ws = window["Asc"]["editor"].wbModel.getWorksheetByName (oRefParsed.sheet);
					if (ws)
						this.Ref = ws.getRange2(oRefParsed.range);
				}
				break;
			case this.Properties.Location: this.setLocation(value);break;
			case this.Properties.Hyperlink: this.Hyperlink = value;break;
			case this.Properties.Tooltip: this.Tooltip = value;break;
		}
	};
	Hyperlink.prototype.applyCollaborative = function (nSheetId, collaborativeEditing) {
		var bbox = this.Ref.getBBox0();
		var OffsetFirst = new AscCommon.CellBase(0, 0);
		var OffsetLast = new AscCommon.CellBase(0, 0);
		OffsetFirst.row = collaborativeEditing.getLockMeRow2(nSheetId, bbox.r1) - bbox.r1;
		OffsetFirst.col = collaborativeEditing.getLockMeColumn2(nSheetId, bbox.c1) - bbox.c1;
		OffsetLast.row = collaborativeEditing.getLockMeRow2(nSheetId, bbox.r2) - bbox.r2;
		OffsetLast.col = collaborativeEditing.getLockMeColumn2(nSheetId, bbox.c2) - bbox.c2;
		this.Ref.setOffsetFirst(OffsetFirst);
		this.Ref.setOffsetLast(OffsetLast);
		return !OffsetFirst.isEmpty() || !OffsetLast.isEmpty();
	};
	Hyperlink.prototype.tryInitLocalLink = function (wb) {
		if (this.Hyperlink && this.Hyperlink[0] === "#") {
			//TODO r1c1 mode. excel ignore and goto "A1" in r1c1
			let sRef = this.Hyperlink.slice(1);
			let result = AscCommon.parserHelp.parse3DRef(sRef);
			let sheetModel, range;
			if (result)
			{
				sheetModel = wb.getWorksheetByName(result.sheet);
				if (sheetModel)
				{
					range = AscCommonExcel.g_oRangeCache.getAscRange(result.range);
				}
			}

			if (!range) {
				range = AscCommon.rx_defName.test(sRef);
			}
			if (!range) {
				range = parserHelp.isTable(sRef, 0, true);
			}
			if (!range) {
				range = AscCommonExcel.g_oRangeCache.getAscRange(sRef);
			}
			if (range) {
				let ws = sheetModel ? sheetModel : wb.getActiveWs();
				this.Ref = ws.getRange3(range.r1, range.c1, range.r2, range.c2);
				this.Location = AscCommon.parserHelp.getEscapeSheetName(ws.getName()) + "!" + sRef;
				this.LocationRange = sRef;
				this.LocationRangeBbox = this.Ref.bbox;
				this.Tooltip = this.Hyperlink;
				this.Hyperlink = null;
				if (sheetModel) {
					this.LocationSheet = sheetModel.sName;
				}
			}
		}
	};
	Hyperlink.prototype.setHyperlinkFunction = function (val) {
		this.bHyperlinkFunction = val;
	};
	Hyperlink.prototype.getHyperlinkFunction = function () {
		return this.bHyperlinkFunction;
	};

	/** @constructor */
	function SheetFormatPr() {
		this.nBaseColWidth = null;
		this.dDefaultColWidth = null;
		this.nOutlineLevelCol = 0;
		this.nOutlineLevelRow = 0;
		this.oAllRow = null;
	}
	SheetFormatPr.prototype.clone = function () {
		var oRes = new SheetFormatPr();
		oRes.nBaseColWidth = this.nBaseColWidth;
		oRes.dDefaultColWidth = this.dDefaultColWidth;
		oRes.nOutlineLevelCol = this.nOutlineLevelCol;
		oRes.nOutlineLevelRow = this.nOutlineLevelRow;
		if (null != this.oAllRow) {
			oRes.oAllRow = this.oAllRow.clone();
		}
		return oRes;
	};
	SheetFormatPr.prototype.correction = function () {
		if (null !== this.dDefaultColWidth && 0 > this.dDefaultColWidth) {
			this.dDefaultColWidth = null;
		}
	};
	/** @constructor */
	function Col(worksheet, index) {
		this.ws = worksheet;
		this.index = index;
		this.BestFit = null;
		this.hd = null;
		this.CustomWidth = null;
		this.width = null;
		this.xfs = null;
		this.outlineLevel = 0;
		this.collapsed = false;

		this.widthPx = null;
		this.charCount = null;
	}

	Col.prototype.fixOnOpening = function () {
		if (null == this.width) {
			this.width = 0;
			this.hd = true;
		} else if (this.width < 0) {
			this.width = 0;
		} else if (this.width > Asc.c_oAscMaxColumnWidth) {
			this.width = Asc.c_oAscMaxColumnWidth;
		}
		if(AscCommon.CurFileVersion < 2)
			this.CustomWidth = 1;
	};
	Col.prototype.moveHor = function (nDif) {
		this.index += nDif;
	};
	Col.prototype.isEqual = function (obj) {
		var bRes = this.BestFit == obj.BestFit && this.hd == obj.hd && this.width == obj.width &&
			this.CustomWidth == obj.CustomWidth && this.outlineLevel == obj.outlineLevel &&
			this.collapsed == obj.collapsed;
		if (bRes) {
			if (null != this.xfs && null != obj.xfs) {
				bRes = this.xfs.isEqual(obj.xfs);
			} else if (null != this.xfs ||
				null != obj.xfs) {
				bRes = false;
			}
		}
		return bRes;
	};
	Col.prototype.isEmpty = function () {
		return null == this.BestFit && null == this.hd && null == this.width && null == this.xfs &&
			null == this.CustomWidth && 0 === this.outlineLevel && false == this.collapsed;
	};
	Col.prototype.isUpdateScroll = function () {
		//TODO temporary added check on CustomWidth -> nColsCount common for scroll/draw. need separate nColsCount for draw and for scroll
		return null !== this.hd || null !== this.xfs || 0 !== this.outlineLevel || false !== this.collapsed || true === this.CustomWidth;
	};
	Col.prototype.clone = function (oNewWs) {
		if (!oNewWs) {
			oNewWs = this.ws;
		}
		var oNewCol = new Col(oNewWs, this.index);
		this.cloneTo(oNewCol);
		return oNewCol;
	};
	Col.prototype.cloneTo = function (oNewCol) {
		if (null != this.BestFit) {
			oNewCol.BestFit = this.BestFit;
		}
		if (null != this.hd) {
			oNewCol.hd = this.hd;
		}
		if (null != this.width) {
			oNewCol.width = this.width;
		}
		if (null != this.CustomWidth) {
			oNewCol.CustomWidth = this.CustomWidth;
		}
		if (null != this.xfs) {
			oNewCol.xfs = this.xfs;
		}
		oNewCol.outlineLevel = this.outlineLevel;
		oNewCol.collapsed = this.collapsed;

		if (null != this.widthPx) {
			oNewCol.widthPx = this.widthPx;
		}
		if (null != this.charCount) {
			oNewCol.charCount = this.charCount;
		}
	};
	Col.prototype.getWidthProp = function () {
		return new AscCommonExcel.UndoRedoData_ColProp(this);
	};
	Col.prototype.setWidthProp = function (prop) {
		if (null != prop) {
			if (null != prop.width) {
				this.width = prop.width;
			} else {
				this.width = null;
			}
			this.setHidden(prop.hd);
			if (null != prop.CustomWidth) {
				this.CustomWidth = prop.CustomWidth;
			} else {
				this.CustomWidth = null;
			}
			if (null != prop.BestFit) {
				this.BestFit = prop.BestFit;
			} else {
				this.BestFit = null;
			}
			if (null != prop.OutlineLevel) {
				this.outlineLevel = prop.OutlineLevel;
			}
		}
	};
	Col.prototype.getStyle = function () {
		return this.xfs;
	};
	Col.prototype._getUpdateRange = function () {
		if (AscCommonExcel.g_nAllColIndex == this.index) {
			return new Asc.Range(0, 0, gc_nMaxCol0, gc_nMaxRow0);
		} else {
			return new Asc.Range(this.index, 0, this.index, gc_nMaxRow0);
		}
	};
	Col.prototype.setStyle = function (xfs) {
		var oldVal = this.xfs;
		this.setStyleInternal(xfs);
		if (AscCommon.History.Is_On() && oldVal !== this.xfs) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_SetStyle, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oldVal, this.xfs));
		}
	};
	Col.prototype.setStyleInternal = function (xfs) {
		this.xfs = g_StyleCache.addXf(xfs);
	};
	Col.prototype.setCellStyle = function (val) {
		var oStyle;
		var newVal = this.ws.workbook.CellStyles._prepareCellStyle(val);
		var oRes = this.ws.workbook.oStyleManager.setCellStyle(this, newVal);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			var oldStyleName = this.ws.workbook.CellStyles.getStyleNameByXfId(oRes.oldVal);
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_SetCellStyle, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oldStyleName, val));

			// Выставляем стиль
			oStyle = this.ws.workbook.CellStyles.getStyleByXfId(oRes.newVal);
			if (oStyle.ApplyFont) {
				this.setFont(oStyle.getFont());
			}
			if (oStyle.ApplyFill) {
				this.setFill(oStyle.getFill());
			}
			if (oStyle.ApplyBorder) {
				this.setBorder(oStyle.getBorder());
			}
			if (oStyle.ApplyNumberFormat) {
				this.setNumFormat(oStyle.getNumFormatStr());
			}
		}
		return oStyle;
	};
	Col.prototype.setNumFormat = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setNum(this, new Num({f: val}));
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_Num, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setNum = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setNum(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_Num, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setFont = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setFont(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			var oldVal = null;
			if (null != oRes.oldVal) {
				oldVal = oRes.oldVal.clone();
			}
			var newVal = null;
			if (null != oRes.newVal) {
				newVal = oRes.newVal.clone();
			}
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_SetFont, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oldVal, newVal));
		}
	};
	Col.prototype.setFontname = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setFontname(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_Fontname, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setFontsize = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setFontsize(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_Fontsize, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setFontcolor = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setFontcolor(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_Fontcolor, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setBold = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setBold(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_Bold, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setItalic = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setItalic(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_Italic, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setUnderline = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setUnderline(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_Underline, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setStrikeout = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setStrikeout(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_Strikeout, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setFontAlign = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setFontAlign(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_FontAlign, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setAlignVertical = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setAlignVertical(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_AlignVertical, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setAlignHorizontal = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setAlignHorizontal(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_AlignHorizontal, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setFill = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setFill(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_Fill, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setBorder = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setBorder(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			var oldVal = null;
			if (null != oRes.oldVal) {
				oldVal = oRes.oldVal.clone();
			}
			var newVal = null;
			if (null != oRes.newVal) {
				newVal = oRes.newVal.clone();
			}
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_Border, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oldVal, newVal));
		}
	};
	Col.prototype.setShrinkToFit = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setShrinkToFit(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_ShrinkToFit, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setWrap = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setWrap(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_Wrap, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setAngle = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setAngle(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_Angle, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setIndent = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setIndent(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_Indent, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setApplyProtection = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setApplyProtection(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_ApplyProtection, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setLocked = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setLocked(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_Locked, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setHiddenFormulas = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setHiddenFormulas(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCol, AscCH.historyitem_RowCol_HiddenFormulas, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Col.prototype.setHidden = function (val) {
		if (this.index >= 0 && (!this.hd !== !val)) {
			this.ws.hiddenManager.addHidden(false, this.index);
		}
		this.hd = val;
	};
	Col.prototype.getHidden = function () {
		return true === this.hd;
	};
	Col.prototype.setIndex = function (val) {
		this.index = val;
	};
	Col.prototype.getIndex = function () {
		return this.index;
	};
	Col.prototype.setOutlineLevel = function (val, bDel, notAddHistory) {
		var oldVal = this.outlineLevel;
		if(null !== val) {
			this.outlineLevel = val;
		} else {
			if(!this.outlineLevel) {
				this.outlineLevel = 0;
			}
			this.outlineLevel = bDel ? this.outlineLevel - 1 : this.outlineLevel + 1;
		}
		if(this.outlineLevel < 0){
			this.outlineLevel = 0;
		} else if(this.outlineLevel > c_maxOutlineLevel) {
			this.outlineLevel = c_maxOutlineLevel;
		} else {
			//TODO ?
			//this._hasChanged = true;
		}

		if (!notAddHistory && AscCommon.History.Is_On() && oldVal != this.outlineLevel) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_GroupCol, this.ws.getId(), this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, false, oldVal, this.outlineLevel));
		}
	};
	Col.prototype.getOutlineLevel = function () {
		return this.outlineLevel;
	};
	Col.prototype.setCollapsed = function (val) {
		this.collapsed = val;
	};
	Col.prototype.getCollapsed = function () {
		return this.collapsed;
	};

	//TODO удалить!
	/*var g_nRowOffsetFlag = 0;
	var g_nRowOffsetXf = g_nRowOffsetFlag + 1;
	var g_nRowOffsetHeight = g_nRowOffsetXf + 4;
	var g_nRowStructSize = g_nRowOffsetHeight + 8;*/
	var g_nRowOffsetFlag = 0;
	var g_nRowOffsetXf = g_nRowOffsetFlag + 1;
	var g_nRowOutlineLevel = g_nRowOffsetXf + 4;
	var g_nRowOffsetHeight = g_nRowOutlineLevel + 1;
	var g_nRowStructSize = g_nRowOffsetHeight + 8;

	var g_nRowFlag_empty = 0;
	var g_nRowFlag_init = 1;
	var g_nRowFlag_hd = 2;
	var g_nRowFlag_CustomHeight = 4;
	var g_nRowFlag_CalcHeight = 8;
	var g_nRowFlag_NullHeight = 16;
	var g_nRowFlag_Collapsed = 32;
	var g_nRowFlag_hdView = 64;

	/**
	 * @constructor
	 */
	function Row(worksheet) {
		this.ws = worksheet;
		this.index = null;
		this.xfs = null;
		this.h = null;
		this.outlineLevel = 0;
		this.flags = g_nRowFlag_init;
		this._hasChanged = false;

		this._tempCell = new AscCommonExcel.Cell(worksheet);
	}
	Row.prototype.clear = function () {
		this.index = null;
		this.xfs = null;
		this.h = null;
		this.outlineLevel = 0;
		this.flags = g_nRowFlag_init;
		this._hasChanged = false;
	};
	Row.prototype.saveContent = function (opt_inCaseOfChange) {
		if (this.index >= 0 && (!opt_inCaseOfChange || this._hasChanged)) {
			this._hasChanged = false;
			var sheetMemory = this.ws.rowsData;
			sheetMemory.checkIndex(this.index);
			var xfSave = this.xfs ? this.xfs.getIndexNumber() : 0;
			var flagToSave = this.flags;
			var heightToSave = this.h;
			if (null === heightToSave) {
				flagToSave |= g_nRowFlag_NullHeight;
				heightToSave = 0;
			}
			sheetMemory.setUint8(this.index, g_nRowOffsetFlag, flagToSave);
			sheetMemory.setUint32(this.index, g_nRowOffsetXf, xfSave);
			sheetMemory.setUint8(this.index, g_nRowOutlineLevel, this.outlineLevel);
			sheetMemory.setFloat64(this.index, g_nRowOffsetHeight, heightToSave);
		}
	};
	Row.prototype.loadContent = function (index) {
		var res = false;
		this.clear();
		this.index = index;
		var sheetMemory = this.ws.rowsData;
		if (sheetMemory.hasIndex(this.index)) {
			this.flags = sheetMemory.getUint8(this.index, g_nRowOffsetFlag);
			if (0 != (g_nRowFlag_init & this.flags)) {
				this.xfs = g_StyleCache.getXf(sheetMemory.getUint32(this.index, g_nRowOffsetXf));
				this.outlineLevel = sheetMemory.getUint8(this.index, g_nRowOutlineLevel);
				if (0 !== (g_nRowFlag_NullHeight & this.flags)) {
					this.flags &= ~g_nRowFlag_NullHeight;
					this.h = null;
				} else {
					this.h = sheetMemory.getFloat64(this.index, g_nRowOffsetHeight);
				}
				res = true;
			}
		}
		return res;
	};
	Row.prototype.setChanged = function (val) {
		this._hasChanged = val;
	};
	Row.prototype.isEmpty = function () {
		return this.isEmptyProp();
	};
	Row.prototype.isEmptyProp = function () {
		//todo
		return null == this.xfs && null == this.h && g_nRowFlag_init == this.flags && 0 === this.outlineLevel;
	};
	Row.prototype.clone = function (oNewWs, renameParams) {
		if (!oNewWs) {
			oNewWs = this.ws;
		}
		var oNewRow = new Row(oNewWs);
		oNewRow.index = this.index;
		oNewRow.flags = this.flags;
		if (null != this.xfs) {
			oNewRow.xfs = this.xfs;
		}
		if (null != this.h) {
			oNewRow.h = this.h;
		}
		if(0 !== this.outlineLevel) {
			oNewRow.outlineLevel = this.outlineLevel;
		}
		return oNewRow;
	};
	Row.prototype.copyFrom = function (row) {
		this.flags = row.flags;
		if (null != row.xfs) {
			this.xfs = row.xfs;
		}
		if (null != row.h) {
			this.h = row.h;
		}
		if(0 !== this.outlineLevel) {
			this.outlineLevel = row.outlineLevel;
		}
		this._hasChanged = true;
	};
	Row.prototype.getDefaultXfs = function () {
		var oRes = null;
		if (null != this.ws.oAllCol && null != this.ws.oAllCol.xfs) {
			oRes = this.ws.oAllCol.xfs.clone();
		}
		return oRes;
	};
	Row.prototype.getHeightProp = function () {
		return new AscCommonExcel.UndoRedoData_RowProp(this);
	};
	Row.prototype.setHeightProp = function (prop) {
		if (null != prop) {
			if (null != prop.h) {
				this.h = prop.h;
			} else {
				this.h = null;
			}
			this.setHidden(prop.hd);
			this.setCustomHeight(prop.CustomHeight);
			this.setOutlineLevel(prop.OutlineLevel, null, true);
		}
	};
	Row.prototype.getStyle = function () {
		return this.xfs;
	};
	Row.prototype._getUpdateRange = function () {
		if (AscCommonExcel.g_nAllRowIndex == this.index) {
			return new Asc.Range(0, 0, gc_nMaxCol0, gc_nMaxRow0);
		} else {
			return new Asc.Range(0, this.index, gc_nMaxCol0, this.index);
		}
	};
	Row.prototype.setStyle = function (xfs) {
		var oldVal = this.xfs;
		this.setStyleInternal(xfs);
		if (AscCommon.History.Is_On() && oldVal !== this.xfs) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_SetStyle, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oldVal, this.xfs));
		}
	};
	Row.prototype.setStyleInternal = function (xfs) {
		this.xfs = g_StyleCache.addXf(xfs);
		this._hasChanged = true;
	};
	Row.prototype.setCellStyle = function (val) {
		var oStyle;
		var newVal = this.ws.workbook.CellStyles._prepareCellStyle(val);
		var oRes = this.ws.workbook.oStyleManager.setCellStyle(this, newVal);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			var oldStyleName = this.ws.workbook.CellStyles.getStyleNameByXfId(oRes.oldVal);
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_SetCellStyle, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oldStyleName, val));

			// Выставляем стиль
			oStyle = this.ws.workbook.CellStyles.getStyleByXfId(oRes.newVal);
			if (oStyle.ApplyFont) {
				this.setFont(oStyle.getFont());
			}
			if (oStyle.ApplyFill) {
				this.setFill(oStyle.getFill());
			}
			if (oStyle.ApplyBorder) {
				this.setBorder(oStyle.getBorder());
			}
			if (oStyle.ApplyNumberFormat) {
				this.setNumFormat(oStyle.getNumFormatStr());
			}
		}
		return oStyle;
	};
	Row.prototype.setNumFormat = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setNum(this, new Num({f: val}));
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_Num, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setNum = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setNum(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_Num, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setFont = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setFont(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			var oldVal = null;
			if (null != oRes.oldVal) {
				oldVal = oRes.oldVal.clone();
			}
			var newVal = null;
			if (null != oRes.newVal) {
				newVal = oRes.newVal.clone();
			}
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_SetFont, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oldVal, newVal));
		}
	};
	Row.prototype.setFontname = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setFontname(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_Fontname, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setFontsize = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setFontsize(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_Fontsize, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setFontcolor = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setFontcolor(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_Fontcolor, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setBold = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setBold(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_Bold, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setItalic = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setItalic(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_Italic, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setUnderline = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setUnderline(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_Underline, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setStrikeout = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setStrikeout(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_Strikeout, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setFontAlign = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setFontAlign(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_FontAlign, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setAlignVertical = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setAlignVertical(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_AlignVertical, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setAlignHorizontal = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setAlignHorizontal(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_AlignHorizontal, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setFill = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setFill(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_Fill, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setBorder = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setBorder(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			var oldVal = null;
			if (null != oRes.oldVal) {
				oldVal = oRes.oldVal.clone();
			}
			var newVal = null;
			if (null != oRes.newVal) {
				newVal = oRes.newVal.clone();
			}
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_Border, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oldVal, newVal));
		}
	};
	Row.prototype.setShrinkToFit = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setShrinkToFit(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_ShrinkToFit, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setWrap = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setWrap(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_Wrap, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setAngle = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setAngle(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_Angle, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setIndent = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setIndent(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_Indent, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setApplyProtection = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setApplyProtection(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_ApplyProtection, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setLocked = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setLocked(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_Locked, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setHiddenFormulas = function (val) {
		var oRes = this.ws.workbook.oStyleManager.setHiddenFormulas(this, val);
		if (AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoRow, AscCH.historyitem_RowCol_HiddenFormulas, this.ws.getId(),
				this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oRes.oldVal, oRes.newVal));
		}
	};
	Row.prototype.setHidden = function (val, bViewLocalChange) {
		var inViewAndFilter = this.ws.getActiveNamedSheetViewId() !== null && this.ws.autoFilters && this.ws.autoFilters.containInFilter(this.index);
		if (!bViewLocalChange) {
			var bCollaborativeChanges = !(this.ws.autoFilters && this.ws.autoFilters.useViewLocalChange) && this.ws.workbook.bCollaborativeChanges;
			bViewLocalChange = !bCollaborativeChanges && inViewAndFilter;
		}
		//если находимся в режиме вью, а приходят изменения для дефолта - не меняем hiddenManager
		if (this.index >= 0 && (!this.getHidden() !== !val) && !(inViewAndFilter && !bViewLocalChange)) {
			this.ws.hiddenManager.addHidden(true, this.index);
		}

		var _rowFlag_hd = !bViewLocalChange ? g_nRowFlag_hd : g_nRowFlag_hdView;
		if (true === val) {
			this.flags |= _rowFlag_hd;
		} else {
			this.flags &= ~_rowFlag_hd;
		}
		this._hasChanged = true;
	};
	Row.prototype.setOutlineLevel = function (val, bDel, notAddHistory) {
		var oldProps = this.outlineLevel;
		if(null !== val) {
			this.outlineLevel = val;
		} else {
			if(!this.outlineLevel) {
				this.outlineLevel = 0;
			}
			this.outlineLevel = bDel ? this.outlineLevel - 1 : this.outlineLevel + 1;
		}
		if(this.outlineLevel < 0){
			this.outlineLevel = 0;
		} else if(this.outlineLevel > c_maxOutlineLevel) {
			this.outlineLevel = c_maxOutlineLevel;
		} else {
			this._hasChanged = true;
		}

		if(!notAddHistory && AscCommon.History.Is_On() && oldProps != this.outlineLevel) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_GroupRow, this.ws.getId(), this._getUpdateRange(), new UndoRedoData_IndexSimpleProp(this.index, true, oldProps, this.outlineLevel));
		}
	};
	Row.prototype.getOutlineLevel = function () {
		return this.outlineLevel;
	};

	Row.prototype.getHidden = function (bViewLocalChange) {
		if (undefined === bViewLocalChange) {
			var bCollaborativeChanges = !(this.ws.autoFilters && this.ws.autoFilters.useViewLocalChange) && this.ws.workbook.bCollaborativeChanges;
			bViewLocalChange = !bCollaborativeChanges && this.ws.getActiveNamedSheetViewId() !== null && this.ws.autoFilters.containInFilter(this.index);
		}
		var _rowFlag_hd = bViewLocalChange ? g_nRowFlag_hdView : g_nRowFlag_hd;
		return 0 !== (_rowFlag_hd & this.flags);
	};
	Row.prototype.setCustomHeight = function (val) {
		if (true === val) {
			this.flags |= g_nRowFlag_CustomHeight;
		} else {
			this.flags &= ~g_nRowFlag_CustomHeight;
		}
		this._hasChanged = true;
	};
	Row.prototype.getCustomHeight = function () {
		return 0 != (g_nRowFlag_CustomHeight & this.flags);
	};
	Row.prototype.setCalcHeight = function (val) {
		if (true === val) {
			this.flags |= g_nRowFlag_CalcHeight;
		} else {
			this.flags &= ~g_nRowFlag_CalcHeight;
		}
		this._hasChanged = true;
	};
	Row.prototype.getCalcHeight = function () {
		return 0 != (g_nRowFlag_CalcHeight & this.flags);
	};
	Row.prototype.setCollapsed = function (val) {
		if (true === val) {
			this.flags |= g_nRowFlag_Collapsed;
		} else {
			this.flags &= ~g_nRowFlag_Collapsed;
		}
		this._hasChanged = true;
	};
	Row.prototype.getCollapsed = function () {
		return 0 != (g_nRowFlag_Collapsed & this.flags);
	};
	Row.prototype.setIndex = function (val) {
		this.index = val;
	};
	Row.prototype.getIndex = function () {
		return this.index;
	};
	Row.prototype.setHeight = function (val) {
		if (val < 0) {
			val = 0;
		} else if (val > Asc.c_oAscMaxRowHeight) {
			val = Asc.c_oAscMaxRowHeight;
		}
		this.h = val;
		this._hasChanged = true;
	};
	Row.prototype.getHeight = function () {
		return this.h;
	};
	Row.prototype.fromXLSB = function(stream, aCellXfs) {
		var end = stream.XlsbReadRecordLength() + stream.GetCurPos();

		this.setIndex(stream.GetULongLE() & 0xFFFFF);
		var style = stream.GetULongLE();
		var ht = stream.GetUShortLE();
		stream.Skip2(1);
		var byteExtra2 = stream.GetUChar();
		this.setOutlineLevel(byteExtra2 & 0x7, null, true);
		if (0 !== (byteExtra2 & 0x8)) {
			this.setCollapsed(true);
		}
		if (0 !== (byteExtra2 & 0x10)) {
			this.setHidden(true);
		}
		if (0 !== (byteExtra2 & 0x20)) {
			this.setCustomHeight(true);
		}
		if (ht > 0 || this.getCustomHeight()) {
			this.setHeight(ht / 20);
		}
		if (0 !== (byteExtra2 & 0x40)) {
			var xf = aCellXfs[style];
			if (xf) {
				this.setStyle(xf);
			}
		}
		stream.Seek2(end);
	};
	Row.prototype.toXLSB = function(stream, offsetIndex, stylesForWrite) {
		stream.XlsbStartRecord(AscCommonExcel.XLSB.rt_ROW_HDR, 17);
		stream.WriteULong((this.index + offsetIndex) & 0xFFFFF);
		var nS = 0;
		if (this.xfs) {
			nS = stylesForWrite.add(this.xfs);
		}
		stream.WriteULong(nS);
		var nHt = 0;
		if (null != this.h) {
			nHt = (this.h * 20) & 0x1FFF;
		}
		stream.WriteUShort(nHt);
		stream.WriteByte(0);
		var byteExtra2 = 0;
		if (this.outlineLevel > 0) {
			byteExtra2 |= this.outlineLevel & 0x7;
		}
		if (this.getCollapsed()) {
			byteExtra2 |= 0x8;
		}
		if (this.getHidden()) {
			byteExtra2 |= 0x10;
		}
		if (this.getCustomHeight()) {
			byteExtra2 |= 0x20;
		}
		if (this.xfs) {
			byteExtra2 |= 0x40;
		}
		stream.WriteByte(byteExtra2);
		stream.WriteByte(0);
		stream.WriteULong(0);
		stream.XlsbEndRecord();
	};
	Row.prototype.onStartNode = function(elem, attr, uq, tagend, getStrNode) {
		var attrVals;
		if ('c' === elem) {
			this._tempCell.clear();
			if (this._tempCell.readAttributes) {
				this._tempCell.readAttributes(attr, uq);
			}
			return this._tempCell;
		}
		return this;
	};
	Row.prototype.onEndNode = function(prevContext, elem) {
		var res = true;
		if ('c' === elem) {
			this._tempCell.saveContent();
		} else {
			res = false;
		}
		return res;
	};

	function getStringFromMultiText(multiText) {
		var sRes = "";
		if (multiText) {
			for (var i = 0, length = multiText.length; i < length; ++i) {
				var elem = multiText[i];
				var text = elem.getFragmentText ? elem.getFragmentText() : elem.text;
				if (null != text && !(elem.format && elem.format.getSkip())) {
					sRes += text;
				}
			}
		}
		return sRes;
	}
	function getStringFromMultiTextSkipToSpace(multiText) {
		var sRes = "";
		if (multiText) {
			for (var i = 0, length = multiText.length; i < length; ++i) {
				var elem = multiText[i];
				var text = elem.getFragmentText ? elem.getFragmentText() : elem.text;
				if (null != text) {
					if(elem.format && elem.format.getSkip()) {
						sRes += " ";
					} else if(!(elem.format && elem.format.getRepeat())) {
						sRes += text;
					} else {
						var j = 0;
					}
				}
			}
		}
		return sRes;
	}
	function isEqualMultiText(multiText1, multiText2) {
		if (multiText1 && multiText2) {
			if (multiText1.length === multiText2.length) {
				for (var i = 0, length = multiText1.length; i < length; ++i) {
					var elem1 = multiText1[i];
					var elem2 = multiText2[i];
					if (!elem1.isEqual(elem2)) {
						return false;
					}
				}
				return true;
			} else {
				return false;
			}
		} else {
			return multiText1 === multiText2;
		}
	}
var g_oCMultiTextElemProperties = {
		text: 0,
		format: 1
	};
function CMultiTextElem() {
	this.text = null;
	this.format = null;
}
CMultiTextElem.prototype =
{
	Properties : g_oCMultiTextElemProperties,
	isEqual : function(val)
	{
		if(null == val)
			return false;
		return this.text == val.text && ((null == this.format && null == val.format) || (null != this.format && null != val.format && this.format.isEqual(val.format)));
	},
	clone : function()
	{
		var oRes = new CMultiTextElem();
		if(null != this.text)
			oRes.text = this.text;
		if(null != this.format)
			oRes.format = this.format.clone();
		return oRes;
	},
	getType : function()
	{
		return UndoRedoDataTypes.ValueMultiTextElem;
	},
	getProperties : function()
	{
		return this.Properties;
	},
	getProperty : function(nType)
	{
		switch(nType)
		{
			case this.Properties.text: return this.text;break;
			case this.Properties.format: return this.format;break;
		}
	},
	setProperty : function(nType, value)
	{
		switch(nType)
		{
			case this.Properties.text: this.text = value;break;
			case this.Properties.format: this.format = value;break;
		}
	},
	setText: function(val)
	{
		this.text = val;
	}
};

var g_oCCellValueProperties = {
		text: 0,
		multiText: 1,
		number: 2,
		type: 3
	};

function CCellValue(opt_cell)
{
	if (opt_cell) {
		this.text = opt_cell.text;
		this.multiText = opt_cell.multiText;
		this.number = opt_cell.number;
		this.type = opt_cell.type;
	} else {
		this.text = null;
		this.multiText = null;
		this.number = null;
		this.type = CellValueType.Number;
	}
}
CCellValue.prototype =
{
	Properties: g_oCCellValueProperties,
	isEqual : function(val)
	{
		if(null == val)
			return false;
		if(this.text != val.text)
			return false;
		if(this.number != val.number)
			return false;
		if(this.type != val.type)
			return false;
		if(null != this.multiText && null != val.multiText)
		{
			if(this.multiText.length == val.multiText.length)
			{
				for(var i = 0, length = this.multiText.length; i < length; ++i)
				{
					if(false == this.multiText[i].isEqual(val.multiText[i]))
						return false;
				}
				return true;
			}
			return false;
		}
		else if(null == this.multiText && null == val.multiText)
			return true;

		return false;
	},
	isEqualCell : function(cell)
	{
		return this.isEqual(cell);
	},
	getType : function()
	{
		return UndoRedoDataTypes.CellValue;
	},
	getProperties : function()
	{
		return this.Properties;
	},
	getProperty : function(nType)
	{
		switch(nType)
		{
			case this.Properties.text: return this.text;break;
			case this.Properties.multiText: return this.multiText;break;
			case this.Properties.number: return this.number;break;
			case this.Properties.type: return this.type;break;
		}
	},
	setProperty : function(nType, value)
	{
		switch(nType)
		{
			case this.Properties.text: this.text = value;break;
			case this.Properties.multiText: this.multiText = value;break;
			case this.Properties.number: this.number = value;break;
			case this.Properties.type: this.type = value;break;
		}
	},
	getTextValue : function(num)
	{
		var multiText = this.multiText;
		var numFormat = AscCommon.oNumFormatCache.get(num && num.getFormat() || "General");
		if (null !== this.text) {
			multiText = numFormat.format(this.text, this.type, AscCommon.gc_nMaxDigCount, false);
		} else if (null !== this.number) {
			if (CellValueType.Bool === this.type) {
				//todo Local
				multiText = [{text: ((this.number == 1) ? AscCommon.cBoolLocal.t : AscCommon.cBoolLocal.f)}];
			} else {
				multiText = numFormat.format(this.number, this.type, AscCommon.gc_nMaxDigCount, false);
			}
		}
		if (null !== multiText) {
			return AscCommonExcel.getStringFromMultiTextSkipToSpace(multiText)
		}
		return "";
	}
};

function RangeDataManagerElem(bbox, data)
{
	this.bbox = bbox;
	this.data = data;
}

	function RangeDataManager(fChange) {
		this.tree = new AscCommon.DataIntervalTree2D();
		this.oDependenceManager = null;
		this.fChange = fChange;

		this.initData = null;
		this.worksheet = null;
	}
	RangeDataManager.prototype._delayedInit = function () {
		if (this.initData) {
			var initData = this.initData;
			this.initData = null;
			var t = this;
			AscCommonExcel.executeInR1C1Mode(false, function () {
				AscCommon.History.TurnOff();
				for (var i = 0; i < initData.length; ++i) {
					var range = t.worksheet.getRange2(initData[i]);
					if (null != range) {
						range.mergeOpen();
					}
				}
				AscCommon.History.TurnOn();
			});
		}
	};
	RangeDataManager.prototype.add = function (bbox, data, oChangeParam) {
		this._delayedInit();
		var oNewElem = new RangeDataManagerElem(new Asc.Range(bbox.c1, bbox.r1, bbox.c2, bbox.r2), data);
		this.tree.insert(bbox, oNewElem);
		if (null != this.fChange) {
			this.fChange.call(this, oNewElem.data, null, oNewElem.bbox, oChangeParam);
		}
	};
	RangeDataManager.prototype.get = function (bbox) {
		this._delayedInit();
		var oRes = {all: [], inner: [], outer: []};
		var intervals = this.tree.searchNodes(bbox);
		for (var i = 0; i < intervals.length; i++) {
			var interval = intervals[i];
			var elem = interval.data;
			if (elem.bbox.isIntersect(bbox)) {
				oRes.all.push(elem);
				if (bbox.containsRange(elem.bbox)) {
					oRes.inner.push(elem);
				} else {
					oRes.outer.push(elem);
				}
			}
		}
		return oRes;
	};
	RangeDataManager.prototype.getFirst = function (bbox) {
		this._delayedInit();
		var intervals = this.tree.searchNodes(bbox);
		for (var i = 0; i < intervals.length; i++) {
			var interval = intervals[i];
			var elem = interval.data;
			if (elem.bbox.isIntersect(bbox)) {
				return elem
			}
		}
		return null;
	};
	RangeDataManager.prototype.getAny = function (bbox) {
		this._delayedInit();
		return this.tree.searchAny(bbox);
	};
	RangeDataManager.prototype.getByCell = function (nRow, nCol) {
		this._delayedInit();
		var bbox = new Asc.Range(nCol, nRow, nCol, nRow);
		var res = this.getAny(bbox);
		if (!res && null != this.oDependenceManager) {
			var oDependence = this.oDependenceManager.getAny(bbox);
			if (oDependence) {
				res = this.getAny(oDependence.bbox);
			}
		}
		return res;
	};
	RangeDataManager.prototype.remove = function (bbox, bInnerOnly, oChangeParam) {
		this._delayedInit();
		var aElems = this.get(bbox);
		var aTargetArray;
		if (bInnerOnly) {
			aTargetArray = aElems.inner;
		} else {
			aTargetArray = aElems.all;
		}
		for (var i = 0, length = aTargetArray.length; i < length; ++i) {
			var elem = aTargetArray[i];
			this.removeElement(elem, oChangeParam);
		}
	};
	RangeDataManager.prototype.removeElement = function (elemToDelete, oChangeParam) {
		this._delayedInit();
		if (null != elemToDelete) {
			this.tree.remove(elemToDelete.bbox, elemToDelete);
			if (null != this.fChange) {
				this.fChange.call(this, elemToDelete.data, elemToDelete.bbox, null, oChangeParam);
			}
		}
	};
	RangeDataManager.prototype.shiftGet = function (bbox, bHor) {
		this._delayedInit();
		var bboxGet = shiftGetBBox(bbox, bHor);
		return {bbox: bboxGet, elems: this.get(bboxGet)};
	};
	RangeDataManager.prototype.shift = function (bbox, bAdd, bHor, oGetRes, oChangeParam) {
		this._delayedInit();
		var _this = this;
		if (null == oGetRes) {
			oGetRes = this.shiftGet(bbox, bHor);
		}
		var offset;
		if (bHor) {
			offset = new AscCommon.CellBase(0, bbox.c2 - bbox.c1 + 1);
		} else {
			offset = new AscCommon.CellBase(bbox.r2 - bbox.r1 + 1, 0);
		}
		if (!bAdd) {
			offset.row *= -1;
			offset.col *= -1;
		}
		this._shiftmove(true, bbox, offset, oGetRes.elems, oChangeParam);
	};
	RangeDataManager.prototype.move = function (from, to, oChangeParam) {
		this._delayedInit();
		var offset = new AscCommon.CellBase(to.r1 - from.r1, to.c1 - from.c1);
		var oGetRes = this.get(from);
		this._shiftmove(false, from, offset, oGetRes, oChangeParam);
	};
	RangeDataManager.prototype._shiftmove = function (bShift, bbox, offset, elems, oChangeParam) {
		var aToChange = [];
		var bAdd = offset.row > 0 || offset.col > 0;
		var bHor = 0 != offset.col ? true : false;
		//сдвигаем inner
		if (elems.inner.length > 0) {
			var bboxAsc = new Asc.Range(bbox.c1, bbox.r1, bbox.c2, bbox.r2);
			for (var i = 0, length = elems.inner.length; i < length; i++) {
				var elem = elems.inner[i];
				var from = elem.bbox;
				var to = null;
				if (bShift) {
					if (bAdd) {
						to = from.clone();
						to.setOffset(offset);
					} else if (!bboxAsc.containsRange(from)) {
						to = from.clone();
						if (bHor) {
							if (to.c1 <= bbox.c2) {
								to.setOffsetFirst(new AscCommon.CellBase(0, bbox.c2 - to.c1 + 1));
							}
						} else {
							if (to.r1 <= bbox.r2) {
								to.setOffsetFirst(new AscCommon.CellBase(bbox.r2 - to.r1 + 1, 0));
							}
						}
						to.setOffset(offset);
					}
				} else {
					to = from.clone();
					to.setOffset(offset);
				}
				aToChange.push({elem: elem, to: to});
			}
		}
		//меняем outer
		if (bShift) {
			if (elems.outer.length > 0) {
				for (var i = 0, length = elems.outer.length; i < length; i++) {
					var elem = elems.outer[i];
					var from = elem.bbox;
					var to = null;
					if (bHor) {
						if (from.c1 < bbox.c1 && bbox.r1 <= from.r1 && from.r2 <= bbox.r2) {
							if (bAdd) {
								to = from.clone();
								to.setOffsetLast(new AscCommon.CellBase(0, bbox.c2 - bbox.c1 + 1));
							} else {
								to = from.clone();
								var nTemp1 = from.c2 - bbox.c1 + 1;
								var nTemp2 = bbox.c2 - bbox.c1 + 1;
								to.setOffsetLast(new AscCommon.CellBase(0, -Math.min(nTemp1, nTemp2)));
							}
						}
					} else {
						if (from.r1 < bbox.r1 && bbox.c1 <= from.c1 && from.c2 <= bbox.c2) {
							if (bAdd) {
								to = from.clone();
								to.setOffsetLast(new AscCommon.CellBase(bbox.r2 - bbox.r1 + 1, 0));
							} else {
								to = from.clone();
								var nTemp1 = from.r2 - bbox.r1 + 1;
								var nTemp2 = bbox.r2 - bbox.r1 + 1;
								to.setOffsetLast(new AscCommon.CellBase(-Math.min(nTemp1, nTemp2), 0));
							}
						}
					}
					if (null != to) {
						aToChange.push({elem: elem, to: to});
					}
				}
			}
		}
		//сначала сортируем чтобы не было конфликтов при сдвиге
		aToChange.sort(function (a, b) {
			return shiftSort(a, b, offset);
		});

		if (null != this.fChange) {
			for (var i = 0, length = aToChange.length; i < length; ++i) {
				var item = aToChange[i];
				this.fChange.call(this, item.elem.data, item.elem.bbox, item.to, oChangeParam);
			}
		}
		//убираем fChange, чтобы потом послать его только на одну операцию, а не 2
		var fOldChange = this.fChange;
		this.fChange = null;
		//сначала удаляем все чтобы не было конфликтов
		for (var i = 0, length = aToChange.length; i < length; ++i) {
			var item = aToChange[i];
			var elem = item.elem;
			this.removeElement(elem, oChangeParam);
		}
		//добавляем измененные ячейки
		for (var i = 0, length = aToChange.length; i < length; ++i) {
			var item = aToChange[i];
			if (null != item.to) {
				this.add(item.to, item.elem.data, oChangeParam);
			}
		}
		this.fChange = fOldChange;
	};
	RangeDataManager.prototype.getAll = function () {
		this._delayedInit();
		var res = [];
		var intervals = this.tree.searchNodes(new Asc.Range(0, 0, gc_nMaxCol0, gc_nMaxRow0));
		for (var i = 0; i < intervals.length; i++) {
			var interval = intervals[i];
			res.push(interval.data);
		}
		return res;
	};
	RangeDataManager.prototype.setDependenceManager = function (oDependenceManager) {
		this.oDependenceManager = oDependenceManager;
	};

	/** @constructor */
	function sparklineGroup(addId) {
		// attributes
		this.type = null;
		this.lineWeight = null;
		this.displayEmptyCellsAs = null;
		this.markers = null;
		this.high = null;
		this.low = null;
		this.first = null;
		this.last = null;
		this.negative = null;
		this.displayXAxis = null;
		this.displayHidden = null;
		this.minAxisType = null;
		this.maxAxisType = null;
		this.rightToLeft = null;
		this.manualMax = null;
		this.manualMin = null;

		this.dateAxis = null;

		// elements
		this.colorSeries = null;
		this.colorNegative = null;
		this.colorAxis = null;
		this.colorMarkers = null;
		this.colorFirst = null;
		this.colorLast = null;
		this.colorHigh = null;
		this.colorLow = null;

		this.f = null;
		this.arrSparklines = [];

		//for drawing preview
		this.canvas = null;

		this.worksheet = null;
		this.Id = null;
		if (addId) {
			this.Id = AscCommon.g_oIdCounter.Get_NewId();
			AscCommon.g_oTableId.Add(this, this.Id);
		}
	}

	sparklineGroup.prototype.getObjectType = function () {
		return AscDFH.historyitem_type_Sparkline;
	};
	sparklineGroup.prototype.Get_Id = function () {
		return this.Id;
	};
	sparklineGroup.prototype.Write_ToBinary2 = function (w) {
		w.WriteLong(this.getObjectType());
		w.WriteString2(this.Id);
		w.WriteString2(this.worksheet ? this.worksheet.getId() : '-1');
	};
	sparklineGroup.prototype.Read_FromBinary2 = function (r) {
		this.Id = r.GetString2();

		// ToDDo не самая лучшая схема добавления на лист...
		var api_sheet = Asc['editor'];
		this.worksheet = api_sheet.wbModel.getWorksheetById(r.GetString2());
		/*if (this.worksheet) {
		 this.worksheet.insertSparklineGroup(this);
		 }*/
	};
	sparklineGroup.prototype.default = function () {
		this.type = Asc.c_oAscSparklineType.Line;
		this.lineWeight = 0.75;
		this.displayEmptyCellsAs = Asc.c_oAscEDispBlanksAs.Zero;
		this.markers = false;
		this.high = false;
		this.low = false;
		this.first = false;
		this.last = false;
		this.negative = false;
		this.displayXAxis = false;
		this.displayHidden = false;
		this.minAxisType = Asc.c_oAscSparklineAxisMinMax.Individual;
		this.maxAxisType = Asc.c_oAscSparklineAxisMinMax.Individual;
		this.rightToLeft = false;
		this.manualMax = null;
		this.manualMin = null;
		this.dateAxis = false;

		// elements
		var defaultSeriesColor = 3629202;
		var defaultOtherColor = 13631488;

		this.colorSeries = new RgbColor(defaultSeriesColor);
		this.colorNegative = new RgbColor(defaultOtherColor);
		this.colorAxis = new RgbColor(defaultOtherColor);
		this.colorMarkers = new RgbColor(defaultOtherColor);
		this.colorFirst = new RgbColor(defaultOtherColor);
		this.colorLast = new RgbColor(defaultOtherColor);
		this.colorHigh = new RgbColor(defaultOtherColor);
		this.colorLow = new RgbColor(defaultOtherColor);
	};
	sparklineGroup.prototype.setWorksheet = function (worksheet, oldWorksheet) {

		let sOldId = this.worksheet ? this.worksheet.Id : null;
		let sNewId = worksheet ? worksheet.Id : null;
		AscCommon.History.Add(new AscDFH.CChangesDrawingsString(this, AscDFH.historyitem_Sparkline_Worksheet, sOldId, sNewId));
		this.worksheet = worksheet;

		if (oldWorksheet) {
			var oldSparklines = [];
			var newSparklines = [];
			for (var i = 0; i < this.arrSparklines.length; ++i) {
				oldSparklines.push(this.arrSparklines[i].clone());
				this.arrSparklines[i].updateWorksheet(worksheet.sName, oldWorksheet.sName);
				newSparklines.push(this.arrSparklines[i].clone());
			}
			AscCommon.History.Add(new AscDFH.CChangesSparklinesChangeData(this, oldSparklines, newSparklines));
		}
	};

	sparklineGroup.prototype.checkProperty = function (propOld, propNew, type, fChangeConstructor) {
		if (null !== propNew && propOld !== propNew) {
			AscCommon.History.Add(new fChangeConstructor(this, type, propOld, propNew));
			return propNew;
		}
		return propOld;
	};

	sparklineGroup.prototype.set = function (val) {
		var getColor = function (color) {
			return color instanceof Asc.asc_CColor ? CorrectAscColor(color) : color ? color.clone(): color;
		};

		this.type = this.checkProperty(this.type, val.type, AscDFH.historyitem_Sparkline_Type, AscDFH.CChangesDrawingsLong);
		this.lineWeight = this.checkProperty(this.lineWeight, val.lineWeight, AscDFH.historyitem_Sparkline_LineWeight, AscDFH.CChangesDrawingsDouble);
		this.displayEmptyCellsAs = this.checkProperty(this.displayEmptyCellsAs, val.displayEmptyCellsAs, AscDFH.historyitem_Sparkline_DisplayEmptyCellsAs, AscDFH.CChangesDrawingsLong);
		this.markers = this.checkProperty(this.markers, val.markers, AscDFH.historyitem_Sparkline_Markers, AscDFH.CChangesDrawingsBool);
		this.high = this.checkProperty(this.high, val.high, AscDFH.historyitem_Sparkline_High, AscDFH.CChangesDrawingsBool);
		this.low = this.checkProperty(this.low, val.low, AscDFH.historyitem_Sparkline_Low, AscDFH.CChangesDrawingsBool);
		this.first = this.checkProperty(this.first, val.first, AscDFH.historyitem_Sparkline_First, AscDFH.CChangesDrawingsBool);
		this.last = this.checkProperty(this.last, val.last, AscDFH.historyitem_Sparkline_Last, AscDFH.CChangesDrawingsBool);
		this.negative = this.checkProperty(this.negative, val.negative, AscDFH.historyitem_Sparkline_Negative, AscDFH.CChangesDrawingsBool);
		this.displayXAxis = this.checkProperty(this.displayXAxis, val.displayXAxis, AscDFH.historyitem_Sparkline_DisplayXAxis, AscDFH.CChangesDrawingsBool);
		this.displayHidden = this.checkProperty(this.displayHidden, val.displayHidden, AscDFH.historyitem_Sparkline_DisplayHidden, AscDFH.CChangesDrawingsBool);
		this.minAxisType = this.checkProperty(this.minAxisType, val.minAxisType, AscDFH.historyitem_Sparkline_MinAxisType, AscDFH.CChangesDrawingsLong);
		this.maxAxisType = this.checkProperty(this.maxAxisType, val.maxAxisType, AscDFH.historyitem_Sparkline_MaxAxisType, AscDFH.CChangesDrawingsLong);
		this.rightToLeft = this.checkProperty(this.rightToLeft, val.rightToLeft, AscDFH.historyitem_Sparkline_RightToLeft, AscDFH.CChangesDrawingsBool);
		this.manualMax = this.checkProperty(this.manualMax, val.manualMax, AscDFH.historyitem_Sparkline_ManualMax, AscDFH.CChangesDrawingsDouble);
		this.manualMin = this.checkProperty(this.manualMin, val.manualMin, AscDFH.historyitem_Sparkline_ManualMin, AscDFH.CChangesDrawingsDouble);
		this.dateAxis = this.checkProperty(this.dateAxis, val.dateAxis, AscDFH.historyitem_Sparkline_DateAxis, AscDFH.CChangesDrawingsBool);

		this.colorSeries = this.checkProperty(this.colorSeries, getColor(val.colorSeries), AscDFH.historyitem_Sparkline_ColorSeries, AscDFH.CChangesDrawingsExcelColor);
		this.colorNegative = this.checkProperty(this.colorNegative, getColor(val.colorNegative), AscDFH.historyitem_Sparkline_ColorNegative, AscDFH.CChangesDrawingsExcelColor);
		this.colorAxis = this.checkProperty(this.colorAxis, getColor(val.colorAxis), AscDFH.historyitem_Sparkline_ColorAxis, AscDFH.CChangesDrawingsExcelColor);
		this.colorMarkers = this.checkProperty(this.colorMarkers, getColor(val.colorMarkers), AscDFH.historyitem_Sparkline_ColorMarkers, AscDFH.CChangesDrawingsExcelColor);
		this.colorFirst = this.checkProperty(this.colorFirst, getColor(val.colorFirst), AscDFH.historyitem_Sparkline_ColorFirst, AscDFH.CChangesDrawingsExcelColor);
		this.colorLast = this.checkProperty(this.colorLast, getColor(val.colorLast), AscDFH.historyitem_Sparkline_colorLast, AscDFH.CChangesDrawingsExcelColor);
		this.colorHigh = this.checkProperty(this.colorHigh, getColor(val.colorHigh), AscDFH.historyitem_Sparkline_ColorHigh, AscDFH.CChangesDrawingsExcelColor);
		this.colorLow = this.checkProperty(this.colorLow, getColor(val.colorLow), AscDFH.historyitem_Sparkline_ColorLow, AscDFH.CChangesDrawingsExcelColor);

		this.f = this.checkProperty(this.f, val.f, AscDFH.historyitem_Sparkline_F, AscDFH.CChangesDrawingsString);

		this.cleanCache();
	};
	sparklineGroup.prototype.clone = function (onlyProps) {
		var res = new sparklineGroup(!onlyProps);
		res.set(this);
		res.f = this.f;

		if (!onlyProps) {
			var newSparklines = [];
			for (var i = 0; i < this.arrSparklines.length; ++i) {
				res.arrSparklines.push(this.arrSparklines[i].clone());
				newSparklines.push(this.arrSparklines[i].clone());
			}
			AscCommon.History.Add(new AscDFH.CChangesSparklinesChangeData(res, null, newSparklines));
		}

		return res;
	};
	sparklineGroup.prototype.draw = function (oDrawingContext) {
		var oCacheView;
		var graphics = new AscCommon.CGraphics();
		graphics.init(oDrawingContext.ctx, oDrawingContext.getWidth(0), oDrawingContext.getHeight(0),
			oDrawingContext.getWidth(3), oDrawingContext.getHeight(3));
		graphics.m_oFontManager = AscCommon.g_fontManager;
		for (var i = 0; i < this.arrSparklines.length; ++i) {
			if (oCacheView = this.arrSparklines[i].oCacheView) {
				oCacheView.draw(graphics);
			}
		}
	};
	sparklineGroup.prototype.cleanCache = function () {
		// ToDo clean only colors (for color scheme)
		for (var i = 0; i < this.arrSparklines.length; ++i) {
			this.arrSparklines[i].oCacheView = null;
		}
	};
	sparklineGroup.prototype.updateCache = function (sheet, ranges) {
		var sparklineRange;
		for (var i = 0; i < this.arrSparklines.length; ++i) {
			sparklineRange = this.arrSparklines[i]._f;
			if (sparklineRange) {
				for (var j = 0; j < ranges.length; ++j) {
					if (sparklineRange.isIntersect(ranges[j], sheet)) {
						this.arrSparklines[i].oCacheView = null;
						break;
					}
				}
			}
		}
	};
	sparklineGroup.prototype.contains = function (c, r) {
		for (var j = 0; j < this.arrSparklines.length; ++j) {
			if (this.arrSparklines[j].contains(c, r)) {
				return j;
			}
		}
		return -1;
	};
	sparklineGroup.prototype.intersectionSimple = function (range) {
		for (var j = 0; j < this.arrSparklines.length; ++j) {
			if (this.arrSparklines[j].intersectionSimple(range)) {
				return j;
			}
		}
		return -1;
	};
	sparklineGroup.prototype.remove = function (range) {
		for (var i = 0; i < this.arrSparklines.length; ++i) {
			if (this.arrSparklines[i].checkInRange(range)) {
				AscCommon.History.Add(new AscDFH.CChangesSparklinesRemoveData(this, this.arrSparklines[i]));
				this.arrSparklines.splice(i, 1);
				--i;
			}
		}
		return (0 === this.arrSparklines.length);
	};
	sparklineGroup.prototype.getLocationRanges = function (onlySingle) {
		var result = new AscCommonExcel.SelectionRange();
		this.arrSparklines.forEach(function (item, i) {
			if (0 === i) {
				result.assign2(item.sqRef);
			} else {
				result.addRange();
				result.getLast().assign2(item.sqRef);
			}
		});
		var unionRange = result.getUnion();
		return (!onlySingle || unionRange.isSingleRange()) ? unionRange : result;
	};
	sparklineGroup.prototype.getDataRanges = function () {
		var isUnion = true;
		var sheet = this.worksheet.getName();
		var result = new AscCommonExcel.SelectionRange();
		this.arrSparklines.forEach(function (item, i) {
			if (item._f) {
				isUnion = isUnion && sheet === item._f.sheet;
				if (0 === i) {
					result.assign2(item._f);
				} else {
					result.addRange();
					result.getLast().assign2(item._f);
				}
			} else {
				isUnion = false;
			}
		});
		var unionRange = isUnion ? result.getUnion() : result;
		return unionRange.isSingleRange() ? unionRange : result;
	};
	sparklineGroup.prototype.setSparklinesFromRange = function (dataRange, locationRange, addToHistory) {
		var t = this;
		var isVertLocationRange = locationRange.c1 === locationRange.c2;
		var count = !isVertLocationRange ? locationRange.c2 - locationRange.c1 + 1 :
			locationRange.r2 - locationRange.r1 + 1;
		var countDataRangeRow = dataRange.r2 - dataRange.r1 + 1;
		var isVertDataRange = countDataRangeRow === count;

		var newArrSparklines = [];
		for (var i = 0; i < count; i++) {
			var sL = new sparkline();

			var _r1 = isVertDataRange ? dataRange.r1 + i : dataRange.r1;
			var _r2 = isVertDataRange ? dataRange.r1 + i : dataRange.r2;
			var _c1 = !isVertDataRange ? dataRange.c1 + i: dataRange.c1;
			var _c2 = !isVertDataRange ? dataRange.c1 + i: dataRange.c2;
			AscCommonExcel.executeInR1C1Mode(false, function () {
				var sheetName = dataRange.sheet ? dataRange.sheet : t.worksheet.sName;
				var f = AscCommon.parserHelp.get3DRef(sheetName, new Asc.Range(_c1, _r1, _c2, _r2).getName());
				sL.setF(f);
			});

			var _col = !isVertLocationRange ? locationRange.c1 + i : locationRange.c1;
			var _row = isVertLocationRange ? locationRange.r1 + i : locationRange.r1;
			sL.sqRef = new Asc.Range(_col, _row, _col, _row);
			sL.sqRef.setAbs(true, true, true, true);
			newArrSparklines.push(sL);
		}
		this.setSparklines(newArrSparklines, addToHistory);
	};
	sparklineGroup.prototype.setSparklines = function (val, addToHistory, opt_addOldPr) {
		var oldPr = null;
		var i;
		if (addToHistory && opt_addOldPr) {
			oldPr = [];
			for (i = 0; i < this.arrSparklines.length; i++) {
				oldPr.push(this.arrSparklines[i].clone());
			}
		}
		this.arrSparklines = val;
		if (addToHistory) {
			var newPr = [];
			for (i = 0; i < this.arrSparklines.length; i++) {
				newPr.push(this.arrSparklines[i].clone());
			}
			AscCommon.History.Add(new AscDFH.CChangesSparklinesChangeData(this, oldPr, newPr));
		}
	};
	sparklineGroup.prototype.isValidDataRef = function (dataRange, locationRange) {
		if (!dataRange || !locationRange) {
			return Asc.c_oAscError.ID.DataRangeError;
		}
		var isOneRow = locationRange.r1 === locationRange.r2;
		var isOneCol = locationRange.c1 === locationRange.c2;
		if (!isOneRow && !isOneCol) {
			return Asc.c_oAscError.ID.SingleColumnOrRowError;
		}
		var count = isOneRow ? locationRange.c2 - locationRange.c1 + 1 : locationRange.r2 - locationRange.r1 + 1;
		var countDataRangeRow = dataRange.r2 - dataRange.r1 + 1;
		var countDataRangeCol = dataRange.c2 - dataRange.c1 + 1;
		if (count !== countDataRangeRow && count !== countDataRangeCol) {
			return Asc.c_oAscError.ID.LocationOrDataRangeError;
		}
		return null;
	};
	sparklineGroup.prototype.getModifySparklinesForPromote = function (from, to, offset) {
		var newArr = [];
		var needAdd = false;
		for (var i = 0; i < this.arrSparklines.length; i++) {
			if (this.arrSparklines[i].sqRef && from.containsRange(this.arrSparklines[i].sqRef)) {
				var cloneElem = this.arrSparklines[i].clone();
				cloneElem.sqRef.setOffset(offset);
				if (to.containsRange(cloneElem.sqRef)) {
					newArr.push(this.arrSparklines[i]);
					if (cloneElem._f) {
						cloneElem._f.setOffset(offset);
						cloneElem.f = cloneElem._f.getName();
					}
					needAdd = true;
				}

				newArr.push(cloneElem);
			} else {
				newArr.push(this.arrSparklines[i].clone());
			}
		}
		return newArr.length && needAdd ? newArr : null;
	};
	sparklineGroup.prototype.asc_getId = function () {
		return this.Id;
	};
	sparklineGroup.prototype.asc_getType = function () {
		return null !== this.type ? this.type : Asc.c_oAscSparklineType.Line;
	};
	sparklineGroup.prototype.asc_getLineWeight = function () {
		return null !== this.lineWeight ? this.lineWeight : 0.75;
	};
	sparklineGroup.prototype.asc_getDisplayEmpty = function () {
		return null !== this.displayEmptyCellsAs ? this.displayEmptyCellsAs : Asc.c_oAscEDispBlanksAs.Zero;
	};
	sparklineGroup.prototype.asc_getMarkersPoint = function () {
		return !!this.markers;
	};
	sparklineGroup.prototype.asc_getHighPoint = function () {
		return !!this.high;
	};
	sparklineGroup.prototype.asc_getLowPoint = function () {
		return !!this.low;
	};
	sparklineGroup.prototype.asc_getFirstPoint = function () {
		return !!this.first;
	};
	sparklineGroup.prototype.asc_getLastPoint = function () {
		return !!this.last;
	};
	sparklineGroup.prototype.asc_getNegativePoint = function () {
		return !!this.negative;
	};
	sparklineGroup.prototype.asc_getDisplayXAxis = function () {
		return this.displayXAxis;
	};
	sparklineGroup.prototype.asc_getDisplayHidden = function () {
		return this.displayHidden;
	};
	sparklineGroup.prototype.asc_getMinAxisType = function () {
		return null !== this.minAxisType ? this.minAxisType : Asc.c_oAscSparklineAxisMinMax.Individual;
	};
	sparklineGroup.prototype.asc_getMaxAxisType = function () {
		return null !== this.maxAxisType ? this.maxAxisType : Asc.c_oAscSparklineAxisMinMax.Individual;
	};
	sparklineGroup.prototype.asc_getRightToLeft = function () {
		return this.rightToLeft;
	};
	sparklineGroup.prototype.asc_getManualMax = function () {
		return this.manualMax;
	};
	sparklineGroup.prototype.asc_getManualMin = function () {
		return this.manualMin;
	};
	sparklineGroup.prototype.asc_getColorSeries = function () {
		return Asc.colorObjToAscColor(this.colorSeries);
	};
	sparklineGroup.prototype.asc_getColorNegative = function () {
		return Asc.colorObjToAscColor(this.colorNegative);
	};
	sparklineGroup.prototype.asc_getColorAxis = function () {
		return Asc.colorObjToAscColor(this.colorAxis);
	};
	sparklineGroup.prototype.asc_getColorMarkers = function () {
		return Asc.colorObjToAscColor(this.colorMarkers);
	};
	sparklineGroup.prototype.asc_getColorFirst = function () {
		return Asc.colorObjToAscColor(this.colorFirst);
	};
	sparklineGroup.prototype.asc_getColorLast = function () {
		return Asc.colorObjToAscColor(this.colorLast);
	};
	sparklineGroup.prototype.asc_getColorHigh = function () {
		return Asc.colorObjToAscColor(this.colorHigh);
	};
	sparklineGroup.prototype.asc_getColorLow = function () {
		return Asc.colorObjToAscColor(this.colorLow);
	};
	sparklineGroup.prototype.asc_getDataRanges = function () {
		var arrResultData = [];
		var arrResultLocation = [];
		var oLocationRanges = this.getLocationRanges(true);
		var oDataRanges = oLocationRanges.isSingleRange() && this.getDataRanges();
		if (oLocationRanges.isSingleRange() && oDataRanges.isSingleRange()) {
			for (var i = 0; i < oLocationRanges.ranges.length; ++i) {
				arrResultData.push(oDataRanges.ranges[i].getName());
				arrResultLocation.push(oLocationRanges.ranges[i].getAbsName());
			}
		} else {
			this.arrSparklines.forEach(function (item) {
				arrResultData.push(item.f || AscCommon.cErrorOrigin['ref']);
				arrResultLocation.push(item.sqRef.getAbsName());
			});
		}
		return [arrResultData.join(AscCommon.FormulaSeparators.functionArgumentSeparator),
			arrResultLocation.join(AscCommon.FormulaSeparators.functionArgumentSeparator)];
	};
	sparklineGroup.prototype.asc_setType = function (val) {
		this.type = val;
	};
	sparklineGroup.prototype.asc_setLineWeight = function (val) {
		this.lineWeight = val;
	};
	sparklineGroup.prototype.asc_setDisplayEmpty = function (val) {
		this.displayEmptyCellsAs = val;
	};
	sparklineGroup.prototype.asc_setMarkersPoint = function (val) {
		this.markers = val;
	};
	sparklineGroup.prototype.asc_setHighPoint = function (val) {
		this.high = val;
	};
	sparklineGroup.prototype.asc_setLowPoint = function (val) {
		this.low = val;
	};
	sparklineGroup.prototype.asc_setFirstPoint = function (val) {
		this.first = val;
	};
	sparklineGroup.prototype.asc_setLastPoint = function (val) {
		this.last = val;
	};
	sparklineGroup.prototype.asc_setNegativePoint = function (val) {
		this.negative = val;
	};
	sparklineGroup.prototype.asc_setDisplayXAxis = function (val) {
		this.displayXAxis = val;
	};
	sparklineGroup.prototype.asc_setDisplayHidden = function (val) {
		this.displayHidden = val;
	};
	sparklineGroup.prototype.asc_setMinAxisType = function (val) {
		this.minAxisType = val;
	};
	sparklineGroup.prototype.asc_setMaxAxisType = function (val) {
		this.maxAxisType = val;
	};
	sparklineGroup.prototype.asc_setRightToLeft = function (val) {
		this.rightToLeft = val;
	};
	sparklineGroup.prototype.asc_setManualMax = function (val) {
		this.manualMax = val;
	};
	sparklineGroup.prototype.asc_setManualMin = function (val) {
		this.manualMin = val;
	};
	sparklineGroup.prototype.asc_setColorSeries = function (val) {
		this.colorSeries = val;
	};
	sparklineGroup.prototype.asc_setColorNegative = function (val) {
		this.colorNegative = val;
	};
	sparklineGroup.prototype.asc_setColorAxis = function (val) {
		this.colorAxis = val;
	};
	sparklineGroup.prototype.asc_setColorMarkers = function (val) {
		this.colorMarkers = val;
	};
	sparklineGroup.prototype.asc_setColorFirst = function (val) {
		this.colorFirst = val;
	};
	sparklineGroup.prototype.asc_setColorLast = function (val) {
		this.colorLast = val;
	};
	sparklineGroup.prototype.asc_setColorHigh = function (val) {
		this.colorHigh = val;
	};
	sparklineGroup.prototype.asc_setColorLow = function (val) {
		this.colorLow = val;
	};

	sparklineGroup.prototype.createExcellColor = function (aColor) {
		var oExcellColor = null;
		if (Array.isArray(aColor)) {
			if (2 === aColor.length) {
				oExcellColor = AscCommonExcel.g_oColorManager.getThemeColor(aColor[0], aColor[1]);
			} else if (1 === aColor.length) {
				oExcellColor = new AscCommonExcel.RgbColor(0x00ffffff & aColor[0]);
			}
		}
		return oExcellColor;
	};

	sparklineGroup.prototype._generateThumbCache = function () {
		function createItem(value) {
			return {numFormatStr: "General", isDateTimeFormat: false, val: value, isHidden: false};
		}

		switch (this.asc_getType()) {
			case Asc.c_oAscSparklineType.Line: {
				return [createItem(4), createItem(-58), createItem(51), createItem(-124), createItem(124), createItem(60)];
			}
			case Asc.c_oAscSparklineType.Column: {
				return [createItem(88), createItem(56), createItem(144), createItem(64), createItem(-56), createItem(-104),
					createItem(-40), createItem(-24), createItem(-56), createItem(104), createItem(56), createItem(80),
					createItem(-56), createItem(88)];
			}
			case Asc.c_oAscSparklineType.Stacked: {
				return [createItem(1), createItem(-1), createItem(-1), createItem(-2), createItem(1), createItem(1),
					createItem(-1), createItem(1), createItem(1), createItem(1), createItem(1), createItem(2), createItem(-1),
					createItem(1)];
			}
		}
		return [];
	};

	sparklineGroup.prototype._drawThumbBySparklineGroup = function (oSparkline, oSparklineGroup, oSparklineView, oGraphics) {
		oSparklineView.initFromSparkline(oSparkline, oSparklineGroup, null, true);
		var t = this;
		AscFormat.ExecuteNoHistory(function () {
			oSparklineView.chartSpace.setWorksheet(t.worksheet);
		}, this, []);

		oSparklineView.chartSpace.extX = 100;
		oSparklineView.chartSpace.extY = 100;
		oSparklineView.chartSpace.x = 0;
		oSparklineView.chartSpace.y = 0;
		var type = oSparklineGroup.asc_getType();
		if (type === Asc.c_oAscSparklineType.Stacked) {
			AscFormat.ExecuteNoHistory(function () {
				var oPlotArea = oSparklineView.chartSpace.chart.plotArea;
				if (!oPlotArea.layout) {
					oPlotArea.setLayout(new AscFormat.CLayout());
				}
				var fPos = 0.32;
				oPlotArea.layout.setWMode(AscFormat.LAYOUT_MODE_FACTOR);
				oPlotArea.layout.setW(1.0);
				oPlotArea.layout.setHMode(AscFormat.LAYOUT_MODE_FACTOR);
				oPlotArea.layout.setH(1 - 2 * fPos);
				oPlotArea.layout.setYMode(AscFormat.LAYOUT_MODE_EDGE);
				oPlotArea.layout.setY(fPos);
			}, this, []);
		}
		if (type === Asc.c_oAscSparklineType.Line) {
			AscFormat.ExecuteNoHistory(function () {
				var oPlotArea = oSparklineView.chartSpace.chart.plotArea;
				if (!oPlotArea.layout) {
					oPlotArea.setLayout(new AscFormat.CLayout());
				}
				var fPos = 0.16;
				oPlotArea.layout.setWMode(AscFormat.LAYOUT_MODE_FACTOR);
				oPlotArea.layout.setW(1 - fPos);
				oPlotArea.layout.setHMode(AscFormat.LAYOUT_MODE_FACTOR);
				oPlotArea.layout.setH(1 - fPos);
			}, this, []);
		}
		AscFormat.ExecuteNoHistory(function () {
			AscFormat.CheckSpPrXfrm(oSparklineView.chartSpace);
		}, this, []);
		oSparklineView.chartSpace.recalculate();
		oSparklineView.chartSpace.brush = AscFormat.CreateSolidFillRGBA(0xFF, 0xFF, 0xFF, 0xFF);

		oSparklineView.chartSpace.draw(oGraphics);
	};

	sparklineGroup.prototype._isEqualStyle = function (oSparklineGroup) {
		var equalColors = function (color1, color2) {
			return color1 ? color1.isEqual(color2) : color1 === color2;
		};
		return equalColors(this.colorSeries, oSparklineGroup.colorSeries) &&
			equalColors(this.colorNegative, oSparklineGroup.colorNegative) &&
			equalColors(this.colorMarkers, oSparklineGroup.colorMarkers) &&
			equalColors(this.colorFirst, oSparklineGroup.colorFirst) &&
			equalColors(this.colorLast, oSparklineGroup.colorLast) &&
			equalColors(this.colorHigh, oSparklineGroup.colorHigh) &&
			equalColors(this.colorLow, oSparklineGroup.colorLow);
	};

	sparklineGroup.prototype.asc_getStyles = function (type) {
		AscCommon.History.TurnOff();
		var aRet = [];
		var nStyleIndex = -1;
		var oSparklineGroup = this.clone(true);
		if ('undefined' !== typeof type) {
			oSparklineGroup.asc_setType(type);
		}

		var canvas = document.createElement('canvas');
		canvas.width = AscCommon.AscBrowser.convertToRetinaValue(50, true);
		canvas.height = AscCommon.AscBrowser.convertToRetinaValue(50, true);

		var oSparklineView = new AscFormat.CSparklineView();
		var oSparkline = new sparkline();
		oSparkline.oCache = this._generateThumbCache();
		var oGraphics = new AscCommon.CGraphics();
		oGraphics.init(canvas.getContext('2d'), canvas.width, canvas.height, 100, 100);
		oGraphics.m_oFontManager = AscCommon.g_fontManager;
		oGraphics.transform(1, 0, 0, 1, 0, 0);

		for (var i = 0; i < 36; ++i) {
			oSparklineGroup.asc_setStyle(i);
			if (nStyleIndex === -1 && this._isEqualStyle(oSparklineGroup)) {
				nStyleIndex = i;
			}

			this._drawThumbBySparklineGroup(oSparkline, oSparklineGroup, oSparklineView, oGraphics);
			aRet.push(canvas.toDataURL("image/png"));
		}
		aRet.push(nStyleIndex);
		AscCommon.History.TurnOn();
		return aRet;
	};

	sparklineGroup.prototype.asc_setStyle = function (nStyleIndex) {
		var oStyle = AscFormat.aSparklinesStyles[nStyleIndex];
		if (oStyle) {
			this.colorSeries = this.createExcellColor(oStyle[0]);
			this.colorNegative = this.createExcellColor(oStyle[1]);
			this.colorAxis = this.createExcellColor(0xff000000);
			this.colorMarkers = this.createExcellColor(oStyle[2]);
			this.colorFirst = this.createExcellColor(oStyle[3]);
			this.colorLast = this.createExcellColor(oStyle[4]);
			this.colorHigh = this.createExcellColor(oStyle[5]);
			this.colorLow = this.createExcellColor(oStyle[6]);
		}
	};
	/** @constructor */
	function sparkline() {
		this.sqRef = null;
		this.f = null;
		this._f = null;

		//for preview
		this.oCache = null;
		this.oCacheView = null;
	}

	sparkline.prototype.clone = function () {
		var res = new sparkline();

		res.sqRef = this.sqRef ? this.sqRef.clone() : null;
		res.f = this.f;
		res._f = this._f ? this._f.clone() : null;

		return res;
	};
	sparkline.prototype.setSqRef = function (sqRef) {
		this.sqRef = AscCommonExcel.g_oRangeCache.getAscRange(sqRef).clone();
		this.sqRef.setAbs(true, true, true, true);
	};
	sparkline.prototype.setF = function (f) {
		//TODO AscCommonExcel.executeInR1C1Mode. пока выставляю сверху. перепроверить и добавить здесь.
		this.f = f;
		this._f = AscCommonExcel.g_oRangeCache.getRange3D(this.f);
	};
	sparkline.prototype.updateWorksheet = function (sheet, oldSheet) {
		var t = this;
		if (this._f && oldSheet === this._f.sheet && (null === this._f.sheet2 || oldSheet === this._f.sheet2)) {
			this._f.setSheet(sheet);
			AscCommonExcel.executeInR1C1Mode(false, function () {
				t.f = t._f.getName();
			});
		}
	};
	sparkline.prototype.checkInRange = function (range) {
		return this.sqRef ? range.isIntersect(this.sqRef) : false;
	};
	sparkline.prototype.contains = function (c, r) {
		return this.sqRef ? this.sqRef.contains(c, r) : false;
	};
	sparkline.prototype.intersectionSimple = function (range) {
		return this.sqRef ? this.sqRef.intersectionSimple(range) : false;
	};

	// For Auto Filters
	/** @constructor */
	function TablePart(handlers) {
		this.Ref = null;
		this.HeaderRowCount = null;
		this.TotalsRowCount = null;
		this.DisplayName = null;
		this.AutoFilter = null;
		this.SortState = null;
		this.TableColumns = null;
		this.TableStyleInfo = null;

		this.QueryTable = null;
		this.tableType = null;

		this.altText = null;
		this.altTextSummary = null;

		this.result = null;
		this.handlers = handlers;
	}

	TablePart.prototype.clone = function () {
		let i, res = new TablePart(this.handlers);
		res.Ref = this.Ref ? this.Ref.clone() : null;
		res.HeaderRowCount = this.HeaderRowCount;
		res.TotalsRowCount = this.TotalsRowCount;
		if (this.AutoFilter) {
			res.AutoFilter = this.AutoFilter.clone();
		}
		if (this.SortState) {
			res.SortState = this.SortState.clone();
		}
		if (this.TableColumns) {
			res.TableColumns = [];
			for (i = 0; i < this.TableColumns.length; ++i) {
				res.TableColumns.push(this.TableColumns[i].clone());
			}
		}
		if (this.TableStyleInfo) {
			res.TableStyleInfo = this.TableStyleInfo.clone();
		}

		if (this.result) {
			res.result = [];
			for (i = 0; i < this.result.length; ++i) {
				res.result.push(this.result[i].clone());
			}
		}

		if (this.QueryTable) {
			res.QueryTable = this.QueryTable.clone();
		}
		res.tableType = this.tableType;

		res.DisplayName = this.DisplayName;

		res.altText = this.altText;
		res.altTextSummary = this.altTextSummary;

		return res;
	};
	TablePart.prototype.renameSheetCopy = function (ws, renameParams) {
		for (let i = 0; i < this.TableColumns.length; ++i) {
			this.TableColumns[i].renameSheetCopy(ws, renameParams);
		}
	};
	TablePart.prototype.removeDependencies = function (opt_cols) {
		if (!opt_cols) {
			opt_cols = this.TableColumns;
		}
		for (let i = 0; i < opt_cols.length; ++i) {
			opt_cols[i].removeDependencies();
		}
	};
	TablePart.prototype.buildDependencies = function () {
		for (let i = 0; i < this.TableColumns.length; ++i) {
			this.TableColumns[i].buildDependencies();
		}
	};
	TablePart.prototype.getAllFormulas = function (formulas) {
		for (let i = 0; i < this.TableColumns.length; ++i) {
			this.TableColumns[i].getAllFormulas(formulas);
		}
	};
	TablePart.prototype.forEachFormula = function (callback) {
		for (let i = 0; i < this.TableColumns.length; ++i) {
			this.TableColumns[i].forEachFormula(callback);
		}
	};
	TablePart.prototype.moveRef = function (col, row) {
		let ref = this.Ref.clone();
		ref.setOffset(new AscCommon.CellBase(row || 0, col || 0));

		this.Ref = ref;
		//event
		this.handlers.trigger("changeRefTablePart", this);

		if (this.AutoFilter) {
			this.AutoFilter.moveRef(col, row);
		}
		if (this.SortState) {
			this.SortState.moveRef(col, row);
		}
	};
	TablePart.prototype.changeRef = function (col, row, bIsFirst, bIsNotChangeAutoFilter) {
		let ref = this.Ref.clone();
		let offset = new AscCommon.CellBase(row || 0, col || 0);
		if (bIsFirst) {
			ref.setOffsetFirst(offset);
		} else {
			ref.setOffsetLast(offset);
		}

		this.Ref = ref;

		//event
		this.handlers.trigger("changeRefTablePart", this);

		if (this.AutoFilter && !bIsNotChangeAutoFilter) {
			this.AutoFilter.changeRef(col, row, bIsFirst);
		}
		if (this.SortState) {
			this.SortState.changeRef(col, row, bIsFirst);
		}
	};
	TablePart.prototype.changeRefOnRange = function (range, autoFilters, generateNewTableColumns) {
		if (!range) {
			return;
		}

		//add table columns
		if (generateNewTableColumns) {
			let newTableColumns = [];
			let intersectionRanges = this.Ref.intersection(range);

			if (null !== intersectionRanges) {
				this.removeDependencies();
				let tableColumn;
				let headerRow = this.isHeaderRow() ? this.Ref.r1 : this.Ref.r1 - 1;
				for (let i = range.c1; i <= range.c2; i++) {
					if (i >= intersectionRanges.c1 && i <= intersectionRanges.c2) {
						let tableIndex = i - this.Ref.c1;
						tableColumn = this.TableColumns[tableIndex];
					} else {
						tableColumn = new TableColumn();
						let cell = autoFilters.worksheet.getCell3(headerRow, i);
						if (!cell.isNullText()) {
							tableColumn.Name =
								autoFilters.checkTableColumnName(newTableColumns.concat(this.TableColumns),
									cell.getValueWithoutFormat());
						}
					}

					newTableColumns.push(tableColumn);
				}

				for (let j = 0; j < newTableColumns.length; j++) {
					tableColumn = newTableColumns[j];
					if (!tableColumn) {
						tableColumn = newTableColumns[j] = new TableColumn();
					}
					if (tableColumn.Name === null) {
						tableColumn.Name = autoFilters._generateColumnName2(newTableColumns);
					}
				}

				this.TableColumns = newTableColumns;
				this.buildDependencies();
			}
		}
		let wb = autoFilters.worksheet.workbook;
		if (this.isTotalsRow() && this.Ref.r2 !== range.r2 && !wb.bUndoChanges && !wb.bRedoChanges) {
			let rangeTotal = autoFilters.worksheet.getRange3(this.Ref.r2, this.Ref.c1, this.Ref.r2, this.Ref.c2);
			rangeTotal.cleanText()
		}

		if (this.QueryTable) {
			this.cleanQueryTables();
			//this.QueryTable.syncIndexes(this.TableColumns);
		}

		this.Ref = new Asc.Range(range.c1, range.r1, range.c2, range.r2);
		//event
		this.handlers.trigger("changeRefTablePart", this);

		if (this.AutoFilter) {
			let filterRange = new Asc.Range(range.c1, range.r1, range.c2, this.isTotalsRow() ? range.r2 - 1 : range.r2);
			this.AutoFilter.changeRefOnRange(filterRange);
		}
	};
	TablePart.prototype.isApplyAutoFilter = function () {
		let res = false;

		if (this.AutoFilter) {
			res = this.AutoFilter.isApplyAutoFilter();
		}

		return res;
	};
	TablePart.prototype.isApplySortConditions = function () {
		let res = false;

		if (this.SortState && this.SortState.SortConditions && this.SortState.SortConditions[0]) {
			res = true;
		}

		return res;
	};

	TablePart.prototype.setHandlers = function (handlers) {
		if (this.handlers === null) {
			this.handlers = handlers;
		}
	};

	TablePart.prototype.deleteTableColumns = function (activeRange) {
		if (!activeRange) {
			return;
		}

		let diff = null, startCol;
		if (activeRange.c1 < this.Ref.c1 && activeRange.c2 >= this.Ref.c1 && activeRange.c2 < this.Ref.c2)//until
		{
			diff = activeRange.c2 - this.Ref.c1 + 1;
			startCol = 0;
		} else if (activeRange.c1 <= this.Ref.c2 && activeRange.c2 > this.Ref.c2 && activeRange.c1 > this.Ref.c1)//after
		{
			diff = this.Ref.c2 - activeRange.c1 + 1;
			startCol = activeRange.c1 - this.Ref.c1;
		} else if (activeRange.c1 >= this.Ref.c1 && activeRange.c2 <= this.Ref.c2)//inside
		{
			diff = activeRange.c2 - activeRange.c1 + 1;
			startCol = activeRange.c1 - this.Ref.c1;
		}

		if (diff !== null) {
			let deleted = this.TableColumns.splice(startCol, diff);
			this.removeDependencies(deleted);

			if (this.QueryTable) {
				this.cleanQueryTables();
				//this.QueryTable.deleteTableColumns(deleted);
				//this.QueryTable.syncIndexes(this.TableColumns);
			}

			//todo undo
			let deletedMap = {};
			for (let i = 0; i < deleted.length; ++i) {
				deletedMap[deleted[i].Name] = 1;
			}
			this.handlers.trigger("deleteColumnTablePart", this.DisplayName, deletedMap);

			if (this.SortState) {
				let bIsDeleteSortState = this.SortState.changeColumns(activeRange, true);
				if (bIsDeleteSortState) {
					this.SortState = null;
				}
			}
		}

	};

	TablePart.prototype.addTableColumns = function (activeRange, autoFilters) {
		let newTableColumns = [], num = 0;
		this.removeDependencies();
		for (let j = 0; j < this.TableColumns.length;) {
			let curCol = num + this.Ref.c1;
			if (activeRange.c1 <= curCol && activeRange.c2 >= curCol) {
				newTableColumns[newTableColumns.length] = new TableColumn();
			} else {
				newTableColumns[newTableColumns.length] = this.TableColumns[j];
				j++
			}

			num++;
		}

		for (let j = 0; j < newTableColumns.length; j++) {
			let tableColumn = newTableColumns[j];
			if (tableColumn.Name === null) {
				tableColumn.Name = autoFilters._generateColumnName2(newTableColumns);
			}
		}

		this.TableColumns = newTableColumns;

		if (this.QueryTable) {
			this.cleanQueryTables();
			//this.QueryTable.syncIndexes(this.TableColumns);
		}

		/*if(this.SortState && this.SortState.SortConditions && this.SortState.SortConditions[0])
		 {
		 let SortConditions = this.SortState.SortConditions[0];
		 if(activeRange.c1 <= SortConditions.Ref.c1)
		 {
		 let offset = activeRange.c2 - activeRange.c1 + 1;
		 SortConditions.Ref.c1 += offset;
		 SortConditions.Ref.c2 += offset;
		 }
		 }*/

		if (this.SortState) {
			this.SortState.changeColumns(activeRange);
		}

		this.buildDependencies();
	};

	TablePart.prototype.addTableLastColumn = function (activeRange, autoFilters, isAddLastColumn) {
		this.removeDependencies();
		let newTableColumns = this.TableColumns;
		newTableColumns.push(new TableColumn());
		newTableColumns[newTableColumns.length - 1].Name = autoFilters._generateColumnName2(newTableColumns);

		this.TableColumns = newTableColumns;
		if (this.QueryTable) {
			this.cleanQueryTables();
			//this.QueryTable.syncIndexes(this.TableColumns);
		}
		this.buildDependencies();
	};

	TablePart.prototype.isAutoFilter = function () {
		return false;
	};

	TablePart.prototype.getAutoFilter = function () {
		return this.AutoFilter;
	};

	TablePart.prototype.getTableRangeForFormula = function (objectParam) {
		let res = null;
		let startRow = this.HeaderRowCount === null ? this.Ref.r1 + 1 : this.Ref.r1;
		let endRow = this.TotalsRowCount ? this.Ref.r2 - 1 : this.Ref.r2;
		switch (objectParam.param) {
			case FormulaTablePartInfo.all: {
				res = new Asc.Range(this.Ref.c1, this.Ref.r1, this.Ref.c2, this.Ref.r2);
				break;
			}
			case FormulaTablePartInfo.data: {
				res = new Asc.Range(this.Ref.c1, startRow, this.Ref.c2, endRow);
				break;
			}
			case FormulaTablePartInfo.headers: {
				if (this.HeaderRowCount === null) {
					res = new Asc.Range(this.Ref.c1, this.Ref.r1, this.Ref.c2, this.Ref.r1);
				} else if (!objectParam.toRef || objectParam.bConvertTableFormulaToRef) {
					res = new Asc.Range(this.Ref.c1, startRow, this.Ref.c2, endRow);
				}
				break;
			}
			case FormulaTablePartInfo.totals: {
				if (this.TotalsRowCount) {
					res = new Asc.Range(this.Ref.c1, this.Ref.r2, this.Ref.c2, this.Ref.r2);
				} else if (!objectParam.toRef || objectParam.bConvertTableFormulaToRef) {
					res = new Asc.Range(this.Ref.c1, startRow, this.Ref.c2, endRow);
				}
				break;
			}
			case FormulaTablePartInfo.thisRow: {
				if (objectParam.cell) {
					if (startRow <= objectParam.cell.r1 && objectParam.cell.r1 <= endRow) {
						res = new Asc.Range(this.Ref.c1, objectParam.cell.r1, this.Ref.c2, objectParam.cell.r1);
					} else if (objectParam.bConvertTableFormulaToRef) {
						res = new Asc.Range(this.Ref.c1, startRow, this.Ref.c2, endRow);
					}
				} else {
					if (objectParam.bConvertTableFormulaToRef) {
						res = new Asc.Range(this.Ref.c1, 0, this.Ref.c2, 0);
					} else {
						res = new Asc.Range(this.Ref.c1, startRow, this.Ref.c2, endRow);
					}
				}
				break;
			}
			case FormulaTablePartInfo.columns: {
				let startCol = this.getTableIndexColumnByName(objectParam.startCol);
				let endCol = this.getTableIndexColumnByName(objectParam.endCol);

				if (startCol === null) {
					break;
				}
				if (endCol === null) {
					endCol = startCol;
				}

				res = new Asc.Range(this.Ref.c1 + startCol, startRow, this.Ref.c1 + endCol, endRow);
				break;
			}
		}
		if (res) {
			if (objectParam.param === FormulaTablePartInfo.thisRow) {
				res.setAbs(false, true, false, true);
			} else {
				res.setAbs(true, true, true, true);
			}
		}
		return res;
	};

	TablePart.prototype.getTableIndexColumnByName = function (name) {
		let res = null;
		if (name === null || name === undefined || !this.TableColumns) {
			return res;
		}

		for (let i = 0; i < this.TableColumns.length; i++) {
			if (name.toLowerCase() === this.TableColumns[i].Name.toLowerCase()) {
				res = i;
				break;
			}
		}

		return res;
	};

	TablePart.prototype.getTableRangeColumnByName = function (name) {
		let res = null;
		if (name === null || name === undefined || !this.TableColumns) {
			return res;
		}

		for (let i = 0; i < this.TableColumns.length; i++) {
			if (name.toLowerCase() === this.TableColumns[i].Name.toLowerCase()) {
				res = new Asc.Range(this.Ref.c1 + i, this.Ref.r1, this.Ref.c1 + i, this.Ref.r2);
				break;
			}
		}

		return res;
	};

	TablePart.prototype.getTableNameColumnByIndex = function (index) {
		let res = null;
		if (index === null || index === undefined || !this.TableColumns) {
			return res;
		}

		if (this.TableColumns[index]) {
			res = this.TableColumns[index].Name;
		}

		return res;
	};

	TablePart.prototype.getIndexByColumnName = function (name) {
		let res = null;
		if (name === null || name === undefined || !this.TableColumns) {
			return res;
		}

		for (let i = 0; i < this.TableColumns.length; i++) {
			if (name.toLowerCase() === this.TableColumns[i].Name.toLowerCase()) {
				res = i;
				break;
			}
		}

		return res;
	};

	TablePart.prototype.showButton = function (val) {
		if (val === false) {
			if (!this.AutoFilter) {
				this.AutoFilter = new AutoFilter();
				this.AutoFilter.Ref = this.Ref;
			}

			this.AutoFilter.showButton(val);
		} else {
			if (this.AutoFilter && this.AutoFilter.FilterColumns && this.AutoFilter.FilterColumns.length) {
				this.AutoFilter.showButton(val);
			}
		}
	};

	TablePart.prototype.isShowButton = function () {
		let res = false;

		if (this.AutoFilter) {
			res = this.AutoFilter.isShowButton();
		} else {
			res = null;
		}

		return res;
	};

	TablePart.prototype.generateTotalsRowLabel = function (ws) {
		if (!this.TableColumns) {
			return;
		}

		//в случае одной колонки выставляем только формулу
		if (this.TableColumns.length > 1) {
			this.TableColumns[0].generateTotalsRowLabel();
		}
		this.TableColumns[this.TableColumns.length - 1].generateTotalsRowFunction(ws, this);
	};

	TablePart.prototype.changeDisplayName = function (newName) {
		this.DisplayName = newName;
	};

	TablePart.prototype.getRangeWithoutHeaderFooter = function () {
		let startRow = this.HeaderRowCount === null ? this.Ref.r1 + 1 : this.Ref.r1;
		let endRow = this.TotalsRowCount ? this.Ref.r2 - 1 : this.Ref.r2;

		return Asc.Range(this.Ref.c1, startRow, this.Ref.c2, endRow);
	};

	TablePart.prototype.getColumnRange = function (index, withoutHeader, withoutFooter, opt_range) {
		let tableRange = opt_range ? opt_range : this.Ref;
		let startRow = tableRange.r1;
		if (withoutHeader && this.isHeaderRow()) {
			startRow++;
		}
		let endRow = tableRange.r2;
		if (withoutFooter && this.isTotalsRow()) {
			endRow--;
		}
		return Asc.Range(tableRange.c1 + index, startRow, tableRange.c1 + index, endRow);
	};

	TablePart.prototype.checkTotalRowFormula = function (ws) {
		for (let i = 0; i < this.TableColumns.length; i++) {
			this.TableColumns[i].checkTotalRowFormula(ws, this);
		}
	};

	TablePart.prototype.changeAltText = function (val) {
		this.altText = val;
	};

	TablePart.prototype.changeAltTextSummary = function (val) {
		this.altTextSummary = val;
	};

	TablePart.prototype.addAutoFilter = function () {
		let autoFilter = new AscCommonExcel.AutoFilter();
		let cloneRef = this.Ref.clone();
		if (this.TotalsRowCount) {
			cloneRef.r2--
		}
		autoFilter.Ref = cloneRef;

		this.AutoFilter = autoFilter;
		return autoFilter;
	};

	TablePart.prototype.isHeaderRow = function () {
		return null === this.HeaderRowCount || this.HeaderRowCount > 0;
	};

	TablePart.prototype.isTotalsRow = function () {
		return this.TotalsRowCount > 0;
	};

	TablePart.prototype.getTotalsRowRange = function () {
		let res = null;

		if (this.TotalsRowCount > 0) {
			res = new Asc.Range(this.Ref.c1, this.Ref.r2, this.Ref.c2, this.Ref.r2);
		}

		return res;
	};

	TablePart.prototype.generateSortState = function () {
		this.SortState = new AscCommonExcel.SortState();
		this.SortState.SortConditions = [];
		this.SortState.SortConditions[0] = new AscCommonExcel.SortCondition();
	};

	//при открытии в случае если не валидный Ref приходит в объекте AutoFilter
	//получаем этот Ref из табличного
	TablePart.prototype.generateAutoFilterRef = function () {
		let res = null;
		if (this.Ref) {
			if (this.isTotalsRow()) {
				res = new Asc.Range(this.Ref.c1, this.Ref.r1, this.Ref.c2, this.Ref.r2 - 1);
			} else {
				res = new Asc.Range(this.Ref.c1, this.Ref.r1, this.Ref.c2, this.Ref.r2);
			}
		}
		return res;
	};

	TablePart.prototype.syncTotalLabels = function (ws) {
		if (this.Ref) {
			if (this.isTotalsRow()) {
				for (let i = 0; i < this.TableColumns.length; i++) {
					if (null !== this.TableColumns[i].TotalsRowLabel) {
						let cell = ws.getCell3(this.Ref.r2, this.Ref.c1 + i);
						if (cell.isFormula()) {
							this.TableColumns[i].TotalsRowLabel = null;
							if (null === this.TableColumns[i].TotalsRowFunction) {
								this.TableColumns[i].TotalsRowFunction = Asc.ETotalsRowFunction.totalrowfunctionCustom;
							}
						} else {
							let val = cell.getValue();
							if (val !== this.TableColumns[i].TotalsRowLabel) {
								this.TableColumns[i].TotalsRowLabel = val;
							}
						}
					}
				}
			}
		}
	};
	TablePart.prototype.cleanQueryTables = function () {
		//удаляю инфомарцию об queryTables после удаления/добавления колонки таблицы
		//связано это с тем, что необходимо следить за всеми полями из queryTables + синхронизировать их с tableColumns,
		//+ есть нюанс - id колонок таблиц сейчас записывается в x2t по порядку в массиве. queryTables связаны с id таблиц
		//необходимо перейти на генерацию id в js и следить за id
		//+ записывать в историю, обрабатывать undo/redo
		//поля, которые попадают в список удаленных, необходимо при undo вовращать с прежними индексами и айдишниками
		//TODO в следующих версиях необходимо реализовать данный функционал в полном объеме
		this.QueryTable = null;
		this.tableType = null;
		for (let i = 0; i < this.TableColumns.length; i++) {
			this.TableColumns[i].queryTableFieldId = null;
			this.TableColumns[i].uniqueName = null;
		}
	};

	TablePart.prototype.getColIdByName = function (name) {
		for (let i = 0; i < this.TableColumns.length; i++) {
			if (name === this.TableColumns[i].Name) {
				return i;
			}
		}
		return null;
	};

	TablePart.prototype.getIndexTableColumnById = function (id) {
		for (let i = 0; i < this.TableColumns.length; i++) {
			if (id === this.TableColumns[i].id) {
				return i + 1;
			}
		}
		return null;
	};

	TablePart.prototype.getSelectionString = function(activeCell, handleSelectionRange) {
		let getColumnNameRange = function (_start, _end, _needAddParenthesis) {
			if (_start !== _end) {
				return "[" + _start + "]" +  ":" +  "[" + _end + "]";
			} else {
				return _needAddParenthesis ? "[" + _start + "]" : _start;
			}
		};

		if (this.Ref.containsRange(handleSelectionRange)) {

			let argsSeparator = AscCommon.FormulaSeparators.functionArgumentSeparator;
			let startCol = this.getTableNameColumnByIndex(handleSelectionRange.c1 - this.Ref.c1);
			let endCol = this.getTableNameColumnByIndex(handleSelectionRange.c2 - this.Ref.c1);

			/* add special character escaping for string inside the table (escaping with single quote) */
			startCol = parserHelp.escapeTableCharacters(startCol, true);
			endCol = parserHelp.escapeTableCharacters(endCol, true);

			if (this.Ref.isEqual(handleSelectionRange)) {
				//Table1[#All]
				return this.DisplayName + "[" + AscCommon.cStrucTableReservedWords.all + "]";
			} else if (this.Ref.r1 === handleSelectionRange.r1 && this.Ref.r2 === handleSelectionRange.r2) {
				//Table1[[#All];[Column1]]
				//Table1[[#All];[Column1]:[Column2]]
				return this.DisplayName + "[" + "[" + AscCommon.cStrucTableReservedWords.all + "]" + argsSeparator + getColumnNameRange(startCol, endCol, true) + "]";
			}

			let dataContains = this._isDataTableContainsRange(handleSelectionRange);
			if (dataContains && dataContains.partialIntersection) {
				//check on @
				if (handleSelectionRange.isOneRow()) {
					if (activeCell.row === handleSelectionRange.r1) {
						if (handleSelectionRange.c1 === this.Ref.c1 && handleSelectionRange.c2 === this.Ref.c2) {
							//all row
							//Table1[@]

							return this.DisplayName + "[" + AscCommon.cStrucTableReservedWords.at + "]";
						} else {
							//part of row
							//Table1[@[Column2]:[Column3]]
							//Table1[@Column1]

							return this.DisplayName + "[" + AscCommon.cStrucTableReservedWords.at + getColumnNameRange(startCol, endCol) + "]";
						}
					}
				}
				return null;
			}
			let totalContains = this._isTotalRowContainsRange(handleSelectionRange);
			if (totalContains && totalContains.partialIntersection) {
				return null;
			}
			let headerContains = this._isHeaderRowContainsRange(handleSelectionRange);
			if (headerContains && headerContains.partialIntersection) {
				return null;
			}


			//1. only data - Table1
			//2. only data Table4[[Column1]:[Column2]] / Table4[Column1]
			if (dataContains && !totalContains && !headerContains) {
				return dataContains.all ? this.DisplayName : this.DisplayName + "[" + getColumnNameRange(startCol, endCol) + "]";
			}

			//3. only all totals - Table4[#Totals]
			//4. only totals - Table4[[#Totals];[Column1]:[Column2]]
			if (!dataContains && totalContains && !headerContains) {
				if (totalContains.all) {
					return this.DisplayName + "[" + AscCommon.cStrucTableReservedWords.totals + "]";
				} else {
					return this.DisplayName + "[" + "[" +  AscCommon.cStrucTableReservedWords.totals + "]" + argsSeparator + getColumnNameRange(startCol, endCol, true) + "]";
				}
			}

			//5. only all headers - Table4[#Headers]
			//6. only headers - Table4[[#Headers];[Column1]:[Column2]]
			if (!dataContains && !totalContains && headerContains) {
				if (headerContains.all) {
					return this.DisplayName + "[" + AscCommon.cStrucTableReservedWords.headers + "]";
				} else {
					return this.DisplayName + "[" + "[" +  AscCommon.cStrucTableReservedWords.headers + "]" + argsSeparator + getColumnNameRange(startCol, endCol, true) + "]";
				}
			}


			//hybrid totals/headers/data

			//Table4[[#Data];[#Totals];[Column1]]
			//Table4[[#Headers];[#Data];[Column1]]
			//Table4[[#Totals];[Column1]]
			//Table4[[#Headers];[#Data];[Column1]:[Column2]]


			//all hybrid
			//Table4[[#Data];[#Totals]]
			if (dataContains || totalContains || headerContains) {
				let isAll;
				let res = this.DisplayName + "[";
				let needDelimiter = false;
				if (headerContains) {
					res += "[" + AscCommon.cStrucTableReservedWords.headers + "]";
					needDelimiter = true;
					if (headerContains.all) {
						isAll = true;
					}
				}
				if (dataContains) {
					if (needDelimiter) {
						res += argsSeparator;
					}
					res += "[" + AscCommon.cStrucTableReservedWords.data + "]";
					needDelimiter = true;
					if (dataContains.all) {
						isAll = true;
					}
				}
				if (totalContains) {
					if (needDelimiter) {
						res += argsSeparator;
					}
					res += "[" + AscCommon.cStrucTableReservedWords.totals + "]";
					needDelimiter = true;
					if (totalContains.all) {
						isAll = true;
					}
				}

				if (!isAll) {
					res += argsSeparator + getColumnNameRange(startCol, endCol, true);
				}

				res += "]";

				return res;
			}
		}
		return null;
	};

	//return {all: bool, partIntersection: bool}/true/false
	TablePart.prototype._getContainsRange = function(range, tablePartRange, checkPartialIntersection) {
		let intersection = range.intersection(tablePartRange);
		if (intersection) {
			if (tablePartRange.isEqual(intersection)) {
				return {all: true};
			} else {
				if (checkPartialIntersection && (intersection.r1 !== tablePartRange.r1 || intersection.r2 !== tablePartRange.r2)) {
					return {partialIntersection: true};
				}
				return true;
			}
		}

		return false;
	};

	TablePart.prototype._isTotalRowContainsRange = function(range) {
		if (range && this.isTotalsRow()) {
			if (this.Ref.containsRange(range)) {
				let _totalRange = this.getTableRangeForFormula({param: FormulaTablePartInfo.totals});
				return this._getContainsRange(range, _totalRange);
			}
		}
		return false;
	};

	TablePart.prototype._isHeaderRowContainsRange = function(range) {
		if (range && this.isHeaderRow()) {
			if (this.Ref.containsRange(range)) {
				let _headerRange = this.getTableRangeForFormula({param: FormulaTablePartInfo.headers});
				return this._getContainsRange(range, _headerRange);
			}
		}
		return false;
	};

	TablePart.prototype._isDataTableContainsRange = function(range) {
		if (this.Ref.containsRange(range)) {
			let _dataRange = this.getTableRangeForFormula({param: FormulaTablePartInfo.data});
			return this._getContainsRange(range, _dataRange, true);
		}
		return false;
	};


	/** @constructor */
	function AutoFilter() {
		this.Ref = null;
		this.FilterColumns = null;
		this.SortState = null;

		this.result = null;
	}

	AutoFilter.prototype.clone = function () {
		var i, res = new AutoFilter();
		res.Ref = this.Ref ? this.Ref.clone() : null;
		res.refTable = this.refTable ? this.refTable.clone() : null;
		if (this.FilterColumns) {
			res.FilterColumns = [];
			for (i = 0; i < this.FilterColumns.length; ++i) {
				res.FilterColumns.push(this.FilterColumns[i].clone());
			}
		}

		if (this.SortState) {
			res.SortState = this.SortState.clone();
		}

		if (this.result) {
			res.result = [];
			for (i = 0; i < this.result.length; ++i) {
				res.result.push(this.result[i].clone());
			}
		}

		return res;
	};

	AutoFilter.prototype.addFilterColumn = function () {
		if (this.FilterColumns === null) {
			this.FilterColumns = [];
		}

		var oNewElem = new FilterColumn();
		this.FilterColumns.push(oNewElem);

		return oNewElem;
	};
	AutoFilter.prototype.getFilterColumnByIndex = function (index) {
		return this.FilterColumns && this.FilterColumns[index];
	};
	AutoFilter.prototype.setStringRef = function (ref) {
		//TODO пересмотреть проверку
		//возможно здесь 3d ref - проверяем
		if(-1 !== ref.indexOf("!"))
		{
			var is3DRef = AscCommon.parserHelp.parse3DRef(ref);
			if(is3DRef){
				ref = is3DRef.range;
			}
		}
		this.Ref = AscCommonExcel.g_oRangeCache.getAscRange(ref);
	};
	AutoFilter.prototype.moveRef = function (col, row) {
		var ref = this.Ref.clone();
		ref.setOffset(new AscCommon.CellBase(row || 0, col || 0));

		if (this.SortState) {
			this.SortState.moveRef(col, row);
		}

		this.Ref = ref;
	};
	AutoFilter.prototype.changeRef = function (col, row, bIsFirst) {
		var ref = this.Ref.clone();
		var offset = new AscCommon.CellBase(row || 0, col || 0);
		if (bIsFirst) {
			ref.setOffsetFirst(offset);
		} else {
			ref.setOffsetLast(offset);
		}

		this.Ref = ref;

		if (this.SortState) {
			this.SortState.changeRef(col, row, bIsFirst);
		}
	};
	AutoFilter.prototype.changeRefOnRange = function (range) {
		if (!range) {
			return;
		}

		this.Ref = new Asc.Range(range.c1, range.r1, range.c2, range.r2);
	};
	AutoFilter.prototype.isApplyAutoFilter = function () {
		var res = false;

		if (this.FilterColumns && this.FilterColumns.length) {
			for (var i = 0; i < this.FilterColumns.length; i++) {
				if (this.FilterColumns[i].isApplyAutoFilter()) {
					res = true;
					break;
				}
			}
		}

		return res;
	};

	AutoFilter.prototype.isApplySortConditions = function () {
		var res = false;

		if (this.SortState && this.SortState.SortConditions && this.SortState.SortConditions[0]) {
			res = true;
		}

		return res;
	};

	AutoFilter.prototype.isAutoFilter = function () {
		return true;
	};

	AutoFilter.prototype.cleanFilters = function () {
		if (!this.FilterColumns) {
			return;
		}

		for (var i = 0; i < this.FilterColumns.length; i++) {
			if (this.FilterColumns[i].ShowButton === false) {
				this.FilterColumns[i].clean();
			} else {
				this.FilterColumns.splice(i, 1);
				i--;
			}
		}
	};

	AutoFilter.prototype.showButton = function (val) {

		if (val === false) {
			if (this.FilterColumns === null) {
				this.FilterColumns = [];
			}

			var columnsLength = this.Ref.c2 - this.Ref.c1 + 1;
			for (var i = 0; i < columnsLength; i++) {
				var filterColumn = this._getFilterColumnByColId(i);
				if (filterColumn) {
					filterColumn.ShowButton = false;
				} else {
					filterColumn = new FilterColumn();
					filterColumn.ColId = i;
					filterColumn.ShowButton = false;
					this.FilterColumns.push(filterColumn);
				}
			}
		} else {
			if (this.FilterColumns && this.FilterColumns.length) {
				for (var i = 0; i < this.FilterColumns.length; i++) {
					this.FilterColumns[i].ShowButton = true;
				}
			}
		}
	};

	AutoFilter.prototype.isShowButton = function () {
		var res = true;

		if (this.FilterColumns && this.FilterColumns.length) {
			for (var i = 0; i < this.FilterColumns.length; i++) {
				if (this.FilterColumns[i].ShowButton === false) {
					res = false;
					break;
				}
			}
		}

		return res;
	};

	AutoFilter.prototype.getRangeWithoutHeaderFooter = function () {
		return Asc.Range(this.Ref.c1, this.Ref.r1 + 1, this.Ref.c2, this.Ref.r2);
	};

	AutoFilter.prototype._getFilterColumnByColId = function (colId) {
		var res = false;

		if (this.FilterColumns && this.FilterColumns.length) {
			for (var i = 0; i < this.FilterColumns.length; i++) {
				if (this.FilterColumns[i].ColId === colId) {
					res = this.FilterColumns[i];
					break;
				}
			}
		}

		return res;
	};

	//функция используется только для изменения данных сортировки, называется так как и в классе TablePart. возможно стоит переименовать.
	AutoFilter.prototype.deleteTableColumns = function (activeRange) {
		if (this.SortState) {
			var bIsDeleteSortState = this.SortState.changeColumns(activeRange, true);
			if (bIsDeleteSortState) {
				this.SortState = null;
			}
		}
	};

	//функция используется только для изменения данных сортировки, называется так как и в классе TablePart. возможно стоит переименовать.
	AutoFilter.prototype.addTableColumns = function (activeRange) {
		if (this.SortState) {
			this.SortState.changeColumns(activeRange);
		}
	};

	AutoFilter.prototype.getIndexByColId = function (colId) {
		var res = null;

		if (!this.FilterColumns) {
			return res;
		}

		for (var i = 0; i < this.FilterColumns.length; i++) {
			if (this.FilterColumns[i].ColId === colId) {
				res = i;
				break;
			}
		}

		return res;
	};

	AutoFilter.prototype.getFilterColumn = function (colId) {
		var res = null;

		if (!this.FilterColumns) {
			return res;
		}

		for (var i = 0; i < this.FilterColumns.length; i++) {
			if (this.FilterColumns[i].ColId === colId) {
				res = this.FilterColumns[i];
				break;
			}
		}

		return res;
	};

	AutoFilter.prototype.isHideButton = function (colId) {
		var filterColumn = this.getFilterColumn(colId);
		return filterColumn && false === filterColumn.ShowButton;
	};

	AutoFilter.prototype.getAutoFilter = function () {
		return this;
	};

	AutoFilter.prototype.hiddenByAnotherFilter = function (worksheet, cellId, row, col, opt_columnsFilter) {
		let result = false;

		let filterColumns = opt_columnsFilter ? opt_columnsFilter : this.FilterColumns;
		if (filterColumns) {
			for (let j = 0; j < filterColumns.length; j++) {
				let filterColumn = opt_columnsFilter ? filterColumns[j].filter : filterColumns[j];
				let colId = filterColumn.ColId;
				if (colId !== cellId) {
					let cell = worksheet.getCell3(row, colId + col);
					let cellFormat = cell.getNumFormat();
					let isDateTimeFormat = cellFormat && cellFormat.isDateTimeFormat() &&
						cell.getType() === window["AscCommon"].CellValueType.Number &&
						cellFormat.getType() !== Asc.c_oAscNumFormatType.Time;

					let isNumberFilter = filterColumn.isApplyCustomFilter();
					let val = (isDateTimeFormat || isNumberFilter) ? cell.getValueWithoutFormat() : cell.getValueWithFormat();
					if (filterColumn.isHideValue(val, isDateTimeFormat, null, cell)) {
						result = true;
						break;
					}
				}
			}
		}

		return result;
	};

	AutoFilter.prototype.setRowHidden = function(worksheet, newFilterColumn, opt_columnsFilter) {
		var startRow = this.Ref.r1 + 1;
		var endRow = this.Ref.r2;

		var colId = newFilterColumn ? newFilterColumn.ColId : null;
		var minChangeRow = null;
		var hiddenObj = {start: this.Ref.r1 + 1, h: null};
		var nHiddenRowCount = 0;
		var nRowsCount = 0;
		for (var i = startRow; i <= endRow; i++) {

			var isHidden = this.hiddenByAnotherFilter(worksheet, colId, i, this.Ref.c1, opt_columnsFilter);
			if(!newFilterColumn) {
				if (isHidden !== worksheet.getRowHidden(i)) {
					if (minChangeRow === null) {
						minChangeRow = i;
					}
				}

				if (true === isHidden) {
					worksheet.setRowHidden(isHidden, i, i);
				}
			} else {
				if (!isHidden) {
					var cell = worksheet.getCell3(i, colId + this.Ref.c1);
					let cellFormat = cell.getNumFormat();
					let isDateTimeFormat = cellFormat && cellFormat.isDateTimeFormat() &&
						cell.getType() === window["AscCommon"].CellValueType.Number &&
						cellFormat.getType() !== Asc.c_oAscNumFormatType.Time;
					var isNumberFilter = false;
					if (newFilterColumn.CustomFiltersObj || newFilterColumn.Top10 || newFilterColumn.DynamicFilter) {
						isNumberFilter = true;
					}

					var currentValue = (isDateTimeFormat || isNumberFilter) ? cell.getValueWithoutFormat() : cell.getValueWithFormat();
					var isSetHidden = newFilterColumn.isHideValue(currentValue, isDateTimeFormat, null, cell);

					if (isSetHidden !== worksheet.getRowHidden(i) && minChangeRow === null) {
						minChangeRow = i;
					}

					//скрываем строки
					if (hiddenObj.h === null) {
						hiddenObj.h = isSetHidden;
						hiddenObj.start = i;
					} else if (hiddenObj.h !== isSetHidden) {
						worksheet.setRowHidden(hiddenObj.h, hiddenObj.start, i - 1);
						if (true === hiddenObj.h) {
							nHiddenRowCount += i - hiddenObj.start;
						}

						hiddenObj.h = isSetHidden;
						hiddenObj.start = i;
					}

					if (i === endRow) {
						worksheet.setRowHidden(hiddenObj.h, hiddenObj.start, i);
						if (true === hiddenObj.h) {
							nHiddenRowCount += i + 1 - hiddenObj.start;
						}
					}
					nRowsCount++;
				} else if (hiddenObj.h !== null) {
					worksheet.setRowHidden(hiddenObj.h, hiddenObj.start, i - 1);
					if (true === hiddenObj.h) {
						nHiddenRowCount += i - hiddenObj.start;
					}
					hiddenObj.h = null
				}

				if (true === isHidden) {
					worksheet.setRowHidden(isHidden, i, i);
				}
			}
		}

		return {nOpenRowsCount: nRowsCount - nHiddenRowCount, nAllRowsCount: endRow - startRow + 1, minChangeRow: minChangeRow};
	};

	AutoFilter.prototype.generateSortState = function () {
		this.SortState = new AscCommonExcel.SortState();
		this.SortState.SortConditions = [];
		this.SortState.SortConditions[0] = new AscCommonExcel.SortCondition();
	};

	AutoFilter.prototype.readAttributes = function(attr, uq) {
		if (attr()) {
			var vals = attr();
			var val;
			val = vals["ref"];
			if (undefined !== val) {
				this.setStringRef(AscCommon.unleakString(uq(val)));
			}
		}
	};
	AutoFilter.prototype.onStartNode = function(elem, attr, uq) {
		var newContext = this;
		if ("filterColumn" === elem) {
			if (!this.FilterColumns) {
				this.FilterColumns = [];
			}
			newContext = new FilterColumn();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.FilterColumns.push(newContext);
		} else if ("sortState" === elem) {
		//todo
			newContext = null;
		} else {
			newContext = null;
		}
		return newContext;
	};
	AutoFilter.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}
		writer.WriteXmlNodeStart(ns + name/*"autoFilter"*/);
		if (null !== this.Ref) {
			writer.WriteXmlAttributeStringEncode("ref", this.Ref.getName());
		}
		writer.WriteXmlAttributesEnd();
		if (this.FilterColumns) {
			for (var i = 0; i < this.FilterColumns.length; ++i) {
				var elem = this.FilterColumns[i];
				elem.toXml(writer, "filterColumn");
			}
		}
		if (this.SortState) {
			this.SortState.toXml(writer, "sortState");
		}

		writer.WriteXmlNodeEnd(ns + name);
	};
	AutoFilter.prototype.deleteFilterColumn = function(index) {
		if (this.FilterColumns && this.FilterColumns[index]) {
			this.FilterColumns.splice(index, 1)
		}
	};

	function FilterColumns() {
		this.ColId = null;
		this.CustomFiltersObj = null;
	}

	FilterColumns.prototype.clone = function () {
		var res = new FilterColumns();
		res.ColId = this.ColId;
		if (this.CustomFiltersObj) {
			res.CustomFiltersObj = this.CustomFiltersObj.clone();
		}

		return res;
	};

	/** @constructor */
	function SortState() {
		this.Ref = null;
		this.CaseSensitive = null;
		this.ColumnSort = null;//false
		this.SortMethod = null;//none
		this.SortConditions = null;
	}

	SortState.prototype.clone = function () {
		var i, res = new SortState();
		res.Ref = this.Ref ? this.Ref.clone() : null;
		res.CaseSensitive = this.CaseSensitive;
		res.ColumnSort = this.ColumnSort;
		res.SortMethod = this.SortMethod;
		if (this.SortConditions) {
			res.SortConditions = [];
			for (i = 0; i < this.SortConditions.length; ++i) {
				res.SortConditions.push(this.SortConditions[i].clone());
			}
		}
		return res;
	};

	SortState.prototype.getType = function() {
		return AscCommonExcel.UndoRedoDataTypes.SortState;
	};
	SortState.prototype.Read_FromBinary2 = function(r) {
		if (r.GetBool()) {
			var r1 = r.GetLong();
			var c1 = r.GetLong();
			var r2 = r.GetLong();
			var c2 = r.GetLong();

			this.Ref = new Asc.Range(c1, r1, c2, r2);
		}
		if (r.GetBool()) {
			this.CaseSensitive = r.GetBool();
		}
		if (r.GetBool()) {
			this.ColumnSort = r.GetBool();
		}
		if (r.GetBool()) {
			this.SortMethod = r.GetBool();
		}

		var length = r.GetLong();
		for (var i = 0; i < length; ++i) {
			var reply = new SortCondition();
			reply.Read_FromBinary2(r);
			if(!this.SortConditions) {
				this.SortConditions = [];
			}
			this.SortConditions.push(reply);
		}
	};
	SortState.prototype.Write_ToBinary2 = function(w) {
		if (null != this.Ref) {
			w.WriteBool(true);
			w.WriteLong(this.Ref.r1);
			w.WriteLong(this.Ref.c1);
			w.WriteLong(this.Ref.r2);
			w.WriteLong(this.Ref.c2);
		} else {
			w.WriteBool(false);
		}
		if (null != this.CaseSensitive) {
			w.WriteBool(true);
			w.WriteBool(this.CaseSensitive);
		} else {
			w.WriteBool(false);
		}
		if (null != this.ColumnSort) {
			w.WriteBool(true);
			w.WriteBool(this.ColumnSort);
		} else {
			w.WriteBool(false);
		}
		if (null != this.SortMethod) {
			w.WriteBool(true);
			w.WriteBool(this.SortMethod);
		} else {
			w.WriteBool(false);
		}

		w.WriteLong(this.SortConditions ? this.SortConditions.length : 0);
		if (this.SortConditions) {
			for (var i = 0; i < this.SortConditions.length; ++i) {
				this.SortConditions[i].Write_ToBinary2(w);
			}
		}
	};
	/*SortState.prototype.applyCollaborative = function (nSheetId, collaborativeEditing) {
		this.nCol = collaborativeEditing.getLockMeColumn2(nSheetId, this.nCol);
		this.nRow = collaborativeEditing.getLockMeRow2(nSheetId, this.nRow);
	};*/

	SortState.prototype.moveRef = function (col, row) {
		var ref = this.Ref.clone();
		ref.setOffset(new AscCommon.CellBase(row || 0, col || 0));

		this.Ref = ref;

		if (this.SortConditions) {
			for (var i = 0; i < this.SortConditions.length; ++i) {
				this.SortConditions[i].moveRef(col, row);
			}
		}
	};

	SortState.prototype.changeRef = function (col, row, bIsFirst) {
		var ref = this.Ref.clone();
		var offset = new AscCommon.CellBase(row || 0, col || 0);
		if (bIsFirst) {
			ref.setOffsetFirst(offset);
		} else {
			ref.setOffsetLast(offset);
		}

		this.Ref = ref;
	};

	SortState.prototype.changeColumns = function (activeRange, isDelete) {
		var bIsSortStateDelete = true;
		//если изменяем диапазон так, что удаляется колонка с сортировкой, удаляем ее
		if (this.SortConditions) {
			for (var i = 0; i < this.SortConditions.length; ++i) {
				var bIsSortConditionsDelete = this.SortConditions[i].changeColumns(activeRange, isDelete);
				if (bIsSortConditionsDelete) {
					this.SortConditions[i] = null;
				} else {
					bIsSortStateDelete = false;
				}
			}
		}
		return bIsSortStateDelete;
	};

	SortState.prototype.setOffset = function(offset, ws, addToHistory) {
		var oldSortState = this.clone();
		var ref = this.Ref.clone();
		ref.setOffset(offset);
		this.Ref = ref;

		if (this.SortConditions) {
			for (var i = 0; i < this.SortConditions.length; ++i) {
				this.SortConditions[i].setOffset(offset);
			}
		}

		if (addToHistory) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoSortState, AscCH.historyitem_SortState_Add, ws.getId(), null,
				new AscCommonExcel.UndoRedoData_SortState(oldSortState, this.clone()));
		}
	};

	SortState.prototype.shift = function(range, offset, ws, addToHistory) {
		var oldSortState = this.clone();

		var from = this.Ref;
		var to = null;
		var bAdd = offset.row > 0 || offset.col > 0;
		var bHor = 0 != offset.col;
		var nTemp1, nTemp2;
		if (bHor) {
			if (from.c1 < range.c1 && range.r1 <= from.r1 && from.r2 <= range.r2) {
				if (bAdd) {
					to = from.clone();
					to.setOffsetLast(new AscCommon.CellBase(0, range.c2 - range.c1 + 1));
				} else {
					to = from.clone();
					nTemp1 = from.c2 - range.c1 + 1;
					nTemp2 = range.c2 - range.c1 + 1;
					to.setOffsetLast(new AscCommon.CellBase(0, -Math.min(nTemp1, nTemp2)));
				}
			} else if(range.c1 <= from.c1 && range.c2 >= from.c1 && range.c2 <= from.c2 && range.r1 <= from.r1 && from.r2 <= range.r2 && !bAdd) {
				to = from.clone();
				nTemp1 = range.c2 - from.c1 + 1;
				nTemp2 = range.c2 - range.c1 + 1;
				to.setOffsetFirst(new AscCommon.CellBase(0, Math.min(nTemp1, nTemp2)));
			}
		} else {
			if (from.r1 < range.r1 && range.c1 <= from.c1 && from.c2 <= range.c2) {
				if (bAdd) {
					to = from.clone();
					to.setOffsetLast(new AscCommon.CellBase(range.r2 - range.r1 + 1, 0));
				} else {
					to = from.clone();
					nTemp1 = from.r2 - range.r1 + 1;
					nTemp2 = range.r2 - range.r1 + 1;
					to.setOffsetLast(new AscCommon.CellBase(-Math.min(nTemp1, nTemp2), 0));
				}
			} else if(range.r1 <= from.r1 && range.r2 >= from.r1 && range.r2 <= from.r2 && range.c1 <= from.c1 && from.c2 <= range.c2 && !bAdd) {
				to = from.clone();
				nTemp1 = range.r2 - from.r1 + 1;
				nTemp2 = range.r2 - range.r1 + 1;
				to.setOffsetFirst(new AscCommon.CellBase(Math.min(nTemp1, nTemp2), 0));
			}
		}

		if (null != to) {
			this.Ref = to;
			if (this.SortConditions) {
				var deleteIndexes = [];
				for (var i = 0; i < this.SortConditions.length; ++i) {
					var sortCondition = this.SortConditions[i];
					var ref = sortCondition.Ref;
					if (offset.row < 0 || offset.col < 0) {
						//смотрим, не попал ли в выделение целиком
						if(range.containsRange(ref)) {
							deleteIndexes[i] = true;
						}
					}
					if(!deleteIndexes[i]) {
						var bboxShift = AscCommonExcel.shiftGetBBox(range, 0 !== offset.col);
						//проверяем, не сдвинулся ли целиком
						if(bboxShift.containsRange(ref)) {
							sortCondition.setOffset(offset);
						} else {
							//осталось проверить на изменение диапазона
							sortCondition.shift(range, offset, this.ColumnSort);
						}
					}
				}
				for(var j in deleteIndexes) {
					this.SortConditions.splice(j, 1);
				}
			}
		}

		if (addToHistory && null != to) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoSortState, AscCH.historyitem_SortState_Add, ws.getId(), null,
				new AscCommonExcel.UndoRedoData_SortState(oldSortState, this.clone()));
		}
	};


	/** @constructor */
	function TableColumn() {
		this.Name = null;
		this.TotalsRowLabel = null;
		this.TotalsRowFunction = null;
		this.TotalsRowFormula = null;
		this.dxf = null;
		this.CalculatedColumnFormula = null;

		//формируется на сохранения
		this.queryTableFieldId = null;
		this.uniqueName = null;

		//queryTableField
		this.clipped = null;
		this.dataBound = null;//default true
		this.fillFormulas = null;
		this.queryName = null;
		this.rowNumbers = null;

		this.id = null;
		//формируется на сохранения
		//this.tableColumnId = null;
	}

	TableColumn.prototype.onFormulaEvent = function (type, eventData) {
	};
	TableColumn.prototype.renameSheetCopy = function (ws, renameParams) {
		if (this.TotalsRowFormula) {
			this.buildDependencies();
			this.TotalsRowFormula.renameSheetCopy(renameParams);
			this.applyTotalRowFormula(this.TotalsRowFormula.assemble(true), ws, true);
		}
	};
	TableColumn.prototype.buildDependencies = function () {
		if (this.TotalsRowFormula) {
			this.TotalsRowFormula.parse();
			this.TotalsRowFormula.buildDependencies();
		}
	};
	TableColumn.prototype.removeDependencies = function () {
		if (this.TotalsRowFormula) {
			this.TotalsRowFormula.removeDependencies();
		}
	};
	TableColumn.prototype.getAllFormulas = function (formulas) {
		if (this.TotalsRowFormula) {
			formulas.push(this.TotalsRowFormula);
		}
	};
	TableColumn.prototype.forEachFormula = function (callback) {
		if (this.TotalsRowFormula) {
			callback(this.TotalsRowFormula);
		}
	};
	TableColumn.prototype.clone = function () {
		var res = new TableColumn();
		res.Name = this.Name;
		res.TotalsRowLabel = this.TotalsRowLabel;
		res.TotalsRowFunction = this.TotalsRowFunction;

		if (this.TotalsRowFormula) {
		res.applyTotalRowFormula(this.TotalsRowFormula.getFormula(), this.TotalsRowFormula.ws, false);
		}
		if (this.dxf) {
			res.dxf = this.dxf.clone;
		}
		res.CalculatedColumnFormula = this.CalculatedColumnFormula;

		res.uniqueName = this.uniqueName;
		res.clipped = this.clipped;
		res.dataBound = this.dataBound;
		res.fillFormulas = this.fillFormulas;
		res.queryName = this.queryName;
		res.rowNumbers = this.rowNumbers;

		res.queryTableFieldId = this.queryTableFieldId;
		res.uniqueName = this.uniqueName;

		res.id = this.id;

		return res;
	};
	TableColumn.prototype.generateTotalsRowLabel = function () {
		//TODO добавить в перевод
		if (this.TotalsRowLabel === null) {
			this.TotalsRowLabel = "Summary";
		}
	};
	TableColumn.prototype.generateTotalsRowFunction = function (ws, tablePart) {
		//TODO добавить в перевод
		if (null === this.TotalsRowFunction && null === this.TotalsRowLabel) {
			var columnRange = this.getRange(tablePart);

			var totalFunction = Asc.ETotalsRowFunction.totalrowfunctionSum;
			if (null !== columnRange) {
				for (var i = columnRange.r1; i <= columnRange.r2; i++) {
					var type = ws.getCell3(i, columnRange.c1).getType();
					if (null !== type && CellValueType.Number !== type) {
						totalFunction = Asc.ETotalsRowFunction.totalrowfunctionCount;
						break;
					}
				}
			}

			this.TotalsRowFunction = totalFunction;
		}
	};

	TableColumn.prototype.getTotalRowFormula = function (tablePart, bLocale) {
		var t = this;
		var res = null;
		var funcNum;
		if (null !== this.TotalsRowFunction) {
			var generateFunction = function (val) {
				var _name = "SUBTOTAL";
				var _f = bLocale && AscCommonExcel.cFormulaFunctionToLocale ? AscCommonExcel.cFormulaFunctionToLocale[_name] : _name;
				var _separator = AscCommon.FormulaSeparators.functionArgumentSeparator;
				return _f + "(" + val + _separator + tablePart.DisplayName + "[" + t.Name + "])";
			};

			switch (this.TotalsRowFunction) {
				case Asc.ETotalsRowFunction.totalrowfunctionAverage: {
					res = generateFunction(101);
					break;
				}
				case Asc.ETotalsRowFunction.totalrowfunctionCount: {
					res = generateFunction(103);
					break;
				}
				case Asc.ETotalsRowFunction.totalrowfunctionCountNums: {
					break;
				}
				case Asc.ETotalsRowFunction.totalrowfunctionCustom: {
					res = this.getTotalsRowFormula();
					break;
				}
				case Asc.ETotalsRowFunction.totalrowfunctionMax: {
					res = generateFunction(104);
					break;
				}
				case Asc.ETotalsRowFunction.totalrowfunctionMin: {
					res = generateFunction(105);
					break;
				}
				case Asc.ETotalsRowFunction.totalrowfunctionNone: {
					break;
				}
				case Asc.ETotalsRowFunction.totalrowfunctionStdDev: {
					res = generateFunction(107);
					break;
				}
				case Asc.ETotalsRowFunction.totalrowfunctionSum: {
					res = generateFunction(109);
					break;
				}
				case Asc.ETotalsRowFunction.totalrowfunctionVar: {
					res = generateFunction(110);
					break;
				}
			}
		}

		return res;
	};

	TableColumn.prototype.cleanTotalsData = function () {
		this.CalculatedColumnFormula = null;
		this.applyTotalRowFormula(null, null, false);
		this.TotalsRowFunction = null;
		this.TotalsRowLabel = null;
	};
	TableColumn.prototype.getTotalsRowFormula = function () {
		return this.TotalsRowFormula ? this.TotalsRowFormula.getFormula() : null;
	};
	TableColumn.prototype.getTotalsRowFunction = function () {
		return this.TotalsRowFunction;
	};
	TableColumn.prototype.getTotalsRowLabel = function () {
		return this.TotalsRowLabel ? this.TotalsRowLabel : null;
	};
	TableColumn.prototype.setTotalsRowFormula = function (val, ws) {
		this.cleanTotalsData();
		if ("=" === val[0]) {
			val = val.substring(1);
		}
		this.applyTotalRowFormula(val, ws, true);
		this.TotalsRowFunction = Asc.ETotalsRowFunction.totalrowfunctionCustom;
	};
	TableColumn.prototype.setTotalsRowFunction = function (val) {
		//функция работает только на undo/redo
		//для того, чтобы работала из меню, необходимо генерировать и добавлять формулу в ячейку
		this.cleanTotalsData();
		this.TotalsRowFunction = val;
	};

	TableColumn.prototype.setTotalsRowLabel = function (val) {
		this.cleanTotalsData();
		this.TotalsRowLabel = val;
	};

	TableColumn.prototype.checkTotalRowFormula = function (ws, tablePart) {
		if (null !== this.TotalsRowFunction &&
			Asc.ETotalsRowFunction.totalrowfunctionCustom !== this.TotalsRowFunction) {
			var totalRowFormula = this.getTotalRowFormula(tablePart, true);

			if (null !== totalRowFormula) {
				this.applyTotalRowFormula(totalRowFormula, ws, true);
				this.TotalsRowFunction = Asc.ETotalsRowFunction.totalrowfunctionCustom;
			}
		}
	};
	TableColumn.prototype.applyTotalRowFormula = function (val, opt_ws, opt_buildDep) {
		this.removeDependencies();
		if (val) {
			this.TotalsRowFormula = new AscCommonExcel.parserFormula(val, this, opt_ws);
			if (opt_buildDep) {
				this.buildDependencies();
			}
		} else {
			this.TotalsRowFormula = null;
		}
	};

	TableColumn.prototype.getRange = function (tablePart, includeHeader, includeTotal) {
		var res = null;

		var ref = tablePart.Ref;
		var startRow = (includeHeader && tablePart.isHeaderRow()) || (!tablePart.isHeaderRow()) ? ref.r1 : ref.r1 + 1;
		var endRow = (includeTotal && tablePart.isTotalsRow()) || (!tablePart.isTotalsRow()) ? ref.r2 : ref.r2 - 1;
		var col = null;
		for (var i = 0; i < tablePart.TableColumns.length; i++) {
			if (this.Name === tablePart.TableColumns[i].Name) {
				col = ref.c1 + i;
				break;
			}
		}

		if (null !== col) {
			res = new Asc.Range(col, startRow, col, endRow);
		}

		return res;
	};

	/** @constructor */
	function TableStyleInfo() {
		this.Name = null;
		this.ShowColumnStripes = null;
		this.ShowRowStripes = null;
		this.ShowFirstColumn = null;
		this.ShowLastColumn = null;
	}

	TableStyleInfo.prototype.clone = function () {
		var res = new TableStyleInfo();
		res.Name = this.Name;
		res.ShowColumnStripes = this.ShowColumnStripes;
		res.ShowRowStripes = this.ShowRowStripes;
		res.ShowFirstColumn = this.ShowFirstColumn;
		res.ShowLastColumn = this.ShowLastColumn;
		return res;
	};
	TableStyleInfo.prototype.setName = function (val) {
		this.Name = val;
	};

	/** @constructor */
	function FilterColumn() {
		this.ColId = null;
		this.Filters = null;
		this.CustomFiltersObj = null;
		this.DynamicFilter = null;
		this.ColorFilter = null;
		this.Top10 = null;
		this.ShowButton = true;
	}

	FilterColumn.prototype.Write_ToBinary2 = function(writer) {
		writer.WriteLong(this.ColId);

		if (null !== this.Filters) {
			writer.WriteBool(true);
			this.Filters.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}

		if (null !== this.CustomFiltersObj) {
			writer.WriteBool(true);
			this.CustomFiltersObj.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}

		if (null !== this.DynamicFilter) {
			writer.WriteBool(true);
			this.DynamicFilter.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}

		if (null !== this.ColorFilter) {
			writer.WriteBool(true);
			this.ColorFilter.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}

		if (null !== this.Top10) {
			writer.WriteBool(true);
			this.Top10.Write_ToBinary2(writer);
		} else {
			writer.WriteBool(false);
		}

		writer.WriteBool(this.ShowButton);
	};

	FilterColumn.prototype.Read_FromBinary2 = function(reader) {
		this.ColId = reader.GetLong();
		var newObj;
		if (reader.GetBool()) {
			newObj = new Filters();
			newObj.Read_FromBinary2(reader);
			this.Filters = newObj;
		}
		if (reader.GetBool()) {
			newObj = new CustomFilters();
			newObj.Read_FromBinary2(reader);
			this.CustomFiltersObj = newObj;
		}
		if (reader.GetBool()) {
			newObj = new DynamicFilter();
			newObj.Read_FromBinary2(reader);
			this.DynamicFilter = newObj;
		}
		if (reader.GetBool()) {
			newObj = new ColorFilter();
			newObj.Read_FromBinary2(reader);
			this.ColorFilter = newObj;
		}
		if (reader.GetBool()) {
			newObj = new Top10();
			newObj.Read_FromBinary2(reader);
			this.Top10 = newObj;
		}

		this.ShowButton = reader.GetBool();
	};


	FilterColumn.prototype.clone = function () {
		var res = new FilterColumn();
		res.ColId = this.ColId;
		if (this.Filters) {
			res.Filters = this.Filters.clone();
		}
		if (this.CustomFiltersObj) {
			res.CustomFiltersObj = this.CustomFiltersObj.clone();
		}
		if (this.DynamicFilter) {
			res.DynamicFilter = this.DynamicFilter.clone();
		}
		if (this.ColorFilter) {
			res.ColorFilter = this.ColorFilter.clone();
		}
		if (this.Top10) {
			res.Top10 = this.Top10.clone();
		}
		res.ShowButton = this.ShowButton;
		return res;
	};
	FilterColumn.prototype.isHideValue = function (val, isDateTimeFormat, top10Length, cell, isLabelFilter) {
		var res = false;
		if (this.Filters) {
			this.Filters._initLowerCaseValues();
			res = this.Filters.isHideValue(val.toLowerCase(), isDateTimeFormat);
		} else if (this.CustomFiltersObj) {
			res = this.CustomFiltersObj.isHideValue(val, isLabelFilter, cell);
		} else if (this.Top10) {
			res = this.Top10.isHideValue(val, top10Length);
		} else if (this.ColorFilter) {
			res = this.ColorFilter.isHideValue(cell);
		} else if (this.DynamicFilter) {
			res = this.DynamicFilter.isHideValue(val);
		}

		return res;
	};
	FilterColumn.prototype.clean = function () {
		this.Filters = null;
		this.CustomFiltersObj = null;
		this.DynamicFilter = null;
		this.ColorFilter = null;
		this.Top10 = null;
	};
	FilterColumn.prototype.createFilter = function (obj) {


		//распознаем формат
		/*var res = AscCommon.g_oFormatParser.parse(val);
		if(null != res)
		{
			//Сравниваем с текущим форматом, если типы совпадают - меняем только значение ячейки
			var nFormatType = oNumFormat.getType();
			if(!((c_oAscNumFormatType.Percent == nFormatType && res.bPercent) ||
				(c_oAscNumFormatType.Currency == nFormatType && res.bCurrency) ||
				(c_oAscNumFormatType.Date == nFormatType && res.bDate) ||
				(c_oAscNumFormatType.Time == nFormatType && res.bTime)) && res.format != oNumFormat.sFormat) {
				this.setNumFormat(res.format);
			}
			this.setTypeInternal(CellValueType.Number);
			this.setValueNumberInternal(res.value);
		}
		else
		{
			this.setTypeInternal(CellValueType.String);
			//проверяем QuotePrefix
			if(val.length > 0 && "'" == val[0])
			{
				this.setQuotePrefix(true);
				val = val.substring(1);
			}
			this.setValueTextInternal(val);
		}*/




		var allFilterOpenElements = false;
		var newFilter;

		switch (obj.filter.type) {
			case c_oAscAutoFilterTypes.ColorFilter: {
				this.ColorFilter = obj.filter.filter.clone();
				break;
			}
			case c_oAscAutoFilterTypes.CustomFilters: {
				obj.filter.filter.check();
				obj.filter.filter.correctFromInterface();
				this.CustomFiltersObj = obj.filter.filter.clone();
				break;
			}
			case c_oAscAutoFilterTypes.DynamicFilter: {
				this.DynamicFilter = obj.filter.filter.clone();
				break;
			}
			case c_oAscAutoFilterTypes.Top10: {
				this.Top10 = obj.filter.filter.clone();
				break;
			}
			case c_oAscAutoFilterTypes.Filters: {
				//если приходит только скрытое Blank, тогда добавляем CustomFilter, так же делает MS
				var addCustomFilter = false;
				for (var i = 0; i < obj.values.length; i++) {
					if ("" === obj.values[i].text && false === obj.values[i].visible) {
						addCustomFilter = true;
					} else if ("" !== obj.values[i].text && false === obj.values[i].visible) {
						addCustomFilter = false;
						break;
					}
				}

				if (addCustomFilter) {
					this.CustomFiltersObj = new CustomFilters();
					this.CustomFiltersObj._generateEmptyValueFilter();
				} else {
					newFilter = new Filters();
					allFilterOpenElements = newFilter.init(obj);

					if (!allFilterOpenElements) {
						this.Filters = newFilter;
					}
				}

				break;
			}
		}

		let visibleDropDown = obj.asc_getVisibleDropDown();
		this.ShowButton = visibleDropDown === false ? false : null;


		return allFilterOpenElements;
	};

	FilterColumn.prototype.isApplyAutoFilter = function () {
		var res = false;

		if (this.Filters !== null || this.CustomFiltersObj !== null || this.DynamicFilter !== null ||
			this.ColorFilter !== null || this.Top10 !== null) {
			res = true;
		}

		return res;
	};

	FilterColumn.prototype.init = function (range) {

		//добавляем данные, которые не передаются из меню при примененни а/ф(в данном случае только DynamicFilter)
		if (null !== this.DynamicFilter) {
			this.DynamicFilter.init(range);
		} else if (null !== this.Top10) {
			this.Top10.init(range);
		}

	};
	FilterColumn.prototype.hasInitByArray = function ()
	{
		return null !== this.Top10;
	};
	FilterColumn.prototype.initByArray = function (arr, isSum) {

		if (null !== this.Top10) {
			this.Top10.initByArray(arr, isSum);
		}

	};

	FilterColumn.prototype.isApplyCustomFilter = function () {
		var res = false;

		if (this.Top10 || this.CustomFiltersObj || this.ColorFilter || this.DynamicFilter) {
			res = true;
		}

		return res;
	};

	FilterColumn.prototype.isOnlyNotEqualEmpty = function () {
		var res = false;

		if (this.CustomFiltersObj) {
			var customFilters = this.CustomFiltersObj.CustomFilters;
			var customFilter = customFilters && 1 === customFilters.length ? customFilters[0] : null;
			if (customFilter && c_oAscCustomAutoFilter.doesNotEqual === customFilter.Operator && " " === customFilter.Val) {
				res = true;
			}
		}

		return res;
	};

	FilterColumn.prototype.isAllClean = function () {

		if (this.Filters === null && this.CustomFiltersObj === null &&
			this.DynamicFilter === null && this.ColorFilter === null && this.Top10 === null &&
			(this.ShowButton === true || this.ShowButton === null)) {
			return true;
		}

		return false;
	};
	FilterColumn.prototype.isColorFilter = function () {
		return this.ColorFilter !== null;
	};
	FilterColumn.prototype.readAttributes = function(attr, uq) {
		if (attr()) {
			var vals = attr();
			var val;
			val = vals["colId"];
			if (undefined !== val) {
				this.ColId = val - 0;
			}
			val = vals["hiddenButton"];
			if (undefined !== val) {
				this.ShowButton = !AscCommon.getBoolFromXml(val);
			}
			val = vals["showButton"];
			if (undefined !== val) {
				this.ShowButton = AscCommon.getBoolFromXml(val);
			}
		}
	};
	FilterColumn.prototype.onStartNode = function(elem, attr, uq) {
		var newContext = this;
		if ("colorFilter" === elem) {
			newContext = new ColorFilter();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.ColorFilter = newContext;
		} else if ("customFilters" === elem) {
			newContext = new CustomFilters();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.CustomFiltersObj = newContext;
		} else if ("dynamicFilter" === elem) {
			newContext = new DynamicFilter();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.DynamicFilter = newContext;
		} else if ("filters" === elem) {
			newContext = new Filters();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.Filters = newContext;
		} else if ("iconFilter" === elem) {
			//todo iconFilter
			newContext = null;
		} else if ("top10" === elem) {
			newContext = new Top10();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.Top10 = newContext;
		} else {
			newContext = null;
		}
		return newContext;
	};
	FilterColumn.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}
		if (!childns) {
			childns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*"filterColumn"*/);

		if (null !== this.ColId) {
			writer.WriteXmlAttributeNumber("colId", this.ColId);
		}
		if (true !== this.ShowButton) {
			if (this.ShowButton) {
				writer.WriteXmlAttributeBool("showButton", this.ShowButton);
			} else {
				writer.WriteXmlAttributeBool("hiddenButton", !this.ShowButton);
			}
		}
		writer.WriteXmlAttributesEnd();

		if (null !== this.ColorFilter) {
			this.ColorFilter.toXml(writer, "colorFilter", childns, childns);
		}
		if (null !== this.CustomFiltersObj) {
			this.CustomFiltersObj.toXml(writer, "customFilters", childns, childns);
		}
		if (null !== this.DynamicFilter) {
			this.DynamicFilter.toXml(writer, "dynamicFilter", childns, childns);
		}
		if (null !== this.Filters) {
			this.Filters.toXml(writer, "filters", childns, childns);
		}
		if (null !== this.Top10) {
			this.Top10.toXml(writer, "top10", childns, childns);
		}
		writer.WriteXmlNodeEnd(ns + name);

	};

	function CT_Filter() {
		//Attributes
		this.Val = null;
	}
	CT_Filter.prototype.readAttributes = function(attr, uq) {
		if (attr()) {
			var vals = attr();
			var val;
			val = vals["val"];
			if (undefined !== val) {
				this.Val = AscCommon.unleakString(uq(val));
			}
		}
	};

	CT_Filter.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}
		writer.WriteXmlNodeStart(ns + name);
		if (null !== this.Val) {
			writer.WriteXmlAttributeStringEncode("val", this.Val);
		}
		writer.WriteXmlAttributesEnd(true);
	};

	/** @constructor */
	function Filters() {
		this.Values = {};
		this.Dates = [];
		this.Blank = null;

		this.lowerCaseValues = null;
	}

	Filters.prototype.Write_ToBinary2 = function (writer) {
		var i, length;
		if(null != this.Values) {
			writer.WriteBool(true);

			length = 0;
			for (i in this.Values) {
				length++;
			}
			writer.WriteLong(length);
			for (i in this.Values) {
				writer.WriteLong(i);
			}

		} else {
			writer.WriteBool(false);
		}

		if(null != this.Dates) {
			writer.WriteBool(true);
			writer.WriteLong(this.Dates.length);
			for(i = 0, length = this.Dates.length; i < length; ++i) {
				this.Dates[i].Write_ToBinary2(writer);
			}
		} else {
			writer.WriteBool(false);
		}

		if(null != this.Blank) {
			writer.WriteBool(true);
			writer.WriteBool(this.Blank);
		} else {
			writer.WriteBool(false);
		}
	};

	Filters.prototype.Read_FromBinary2 = function (reader) {
		var i, length;
		if (reader.GetBool()) {
			length = reader.GetLong();
			for (i = 0; i < length; ++i) {
				this.Values[reader.GetLong()] = true;
			}
		}
		if (reader.GetBool()) {
			length = reader.GetLong();
			for (i = 0; i < length; ++i) {
				var _date = new AutoFilterDateElem();
				this.Dates.push(_date.Read_FromBinary2(reader));
			}
		}
		if (reader.GetBool()) {
			this.Blank = reader.GetBool();
		}
	};

	Filters.prototype.clone = function () {
		var i, res = new Filters();
		for (i in this.Values) {
			res.Values[i] = this.Values[i];
		}
		if (this.Dates) {
			for (i = 0; i < this.Dates.length; ++i) {
				res.Dates.push(this.Dates[i].clone());
			}
		}
		res.Blank = this.Blank;
		return res;
	};
	Filters.prototype.init = function (obj) {
		var allFilterOpenElements = true;
		for (var i = 0; i < obj.values.length; i++) {
			if (obj.values[i].visible) {
				if (obj.values[i].isDateFormat) {
					if (obj.values[i].text === "") {
						this.Blank = true;
					} else {
						var dateGroupItem = new DateGroupItem();
						var autoFilterDateElem = new AutoFilterDateElem(obj.values[i].val, obj.values[i].val, obj.values[i].dateTimeGrouping);
						dateGroupItem.convertRangeToDateGroupItem(autoFilterDateElem);
						autoFilterDateElem.convertDateGroupItemToRange(dateGroupItem);

						this.Dates.push(autoFilterDateElem);
					}
				} else {
					if (obj.values[i].text === "") {
						this.Blank = true;
					} else {
						this.Values[obj.values[i].text] = true;
					}
				}
			} else {
				allFilterOpenElements = false;
			}
		}
		this.sortDate();
		this._initLowerCaseValues();

		return allFilterOpenElements;
	};
	Filters.prototype.isHideValue = function (val, isDateTimeFormat) {
		var res = false;

		val = window["Asc"].trim(val);
		if (isDateTimeFormat && this.Dates) {
			if (val === "") {
				res = !this.Blank ? true : false;
			} else {
				res = this.binarySearch(val, this.Dates) !== -1 ? false : true;
			}
		} else if (this.Values) {
			if (val === "") {
				res = !this.Blank ? true : false;
			} else {
				res = !this.lowerCaseValues[val] ? true : false;
			}
		}

		return res;
	};

	Filters.prototype.binarySearch = function (val, array) {
		var i = 0, j = array.length - 1, k;
		val = parseFloat(val);

		while (i <= j) {
			k = Math.floor((i + j) / 2);

			if (val >= array[k].start && val < array[k].end) {
				return k;
			} else if (val < array[k].start) {
				j = k - 1;
			} else {
				i = k + 1;
			}
		}

		return -1;
	};

	Filters.prototype.linearSearch = function (val, array) {
		var n = array.length, i = 0;
		val = parseFloat(val);

		while (i <= n && !(array[i] && val >= array[i].start && val < array[i].end))
			i++;

		if (i < n) {
			return i;
		} else {
			return -1;
		}
	};
	Filters.prototype._initLowerCaseValues = function () {
		if (this.lowerCaseValues === null) {
			this.lowerCaseValues = {};
			for (var i in this.Values) {
				this.lowerCaseValues[i.toLowerCase()] = true;
			}
		}
	};
	Filters.prototype.sortDate = function () {
		if (this.Dates && this.Dates.length) {
			this.Dates.sort(function sortArr(a, b) {
				return a.start - b.start;
			})
		}
	};

	Filters.prototype.clean = function () {
		this.Values = {};
		this.Dates = [];
		this.Blank = null;
	};
	Filters.prototype.readAttributes = function(attr, uq) {
		if (attr()) {
			var vals = attr();
			var val;
			val = vals["blank"];
			if (undefined !== val) {
				this.Blank = AscCommon.getBoolFromXml(val);
			}
			//todo calendarType
		}
	};
	Filters.prototype.onStartNode = function(elem, attr, uq) {
		var newContext = this;
		if ("filter" === elem) {
			newContext = new CT_Filter();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			if(null != newContext.Val)
				this.Values[newContext.Val] = 1;
		} else if ("dateGroupItem" === elem) {
			newContext = new DateGroupItem();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			var autoFilterDateElem = new AscCommonExcel.AutoFilterDateElem();
			autoFilterDateElem.convertDateGroupItemToRange(newContext);
			this.Dates.push(autoFilterDateElem);
		} else {
			newContext = null;
		}
		return newContext;
	};
	Filters.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}
		if (!childns) {
			childns = "";
		}

		writer.WriteXmlNodeStart(ns + name);
		if (null !== this.Blank) {
			writer.WriteXmlAttributeBool("blank", this.Blank);
		}
		writer.WriteXmlAttributesEnd();
		for (var val in this.Values) {
			var filter = new AscCommonExcel.CT_Filter();
			filter.Val = val;
			filter.toXml(writer, "filter", childns, childns);
		}
		for (var i = 0; i < this.Dates.length; ++i) {
			var elem = this.Dates[i];
			var dateGroupItem = new AscCommonExcel.DateGroupItem();
			dateGroupItem.convertRangeToDateGroupItem(elem);
			dateGroupItem.toXml(writer, "dateGroupItem", childns, childns);
		}
		writer.WriteXmlNodeEnd(ns + name);
	};

	/** @constructor */
	function Filter() {
		this.Val = null;
	}
	/** @constructor */
	function DateGroupItem() {
		this.DateTimeGrouping = null;
		this.Day = null;
		this.Hour = null;
		this.Minute = null;
		this.Month = null;
		this.Second = null;
		this.Year = null;
	}
	DateGroupItem.prototype.clone = function() {
		var res = new DateGroupItem();
		res.DateTimeGrouping = this.DateTimeGrouping;
		res.Day = this.Day;
		res.Hour = this.Hour;
		res.Minute = this.Minute;
		res.Month = this.Month;
		res.Second = this.Second;
		res.Year = this.Year;
		return res;
	};
	DateGroupItem.prototype.convertRangeToDateGroupItem = function(range) {
		var startUtcDate = AscCommon.NumFormat.prototype.parseDate(range.start);
		var year = startUtcDate.year;
		var month = startUtcDate.month + 1;
		var day = startUtcDate.d;
		var hour = startUtcDate.hour;
		var minute = startUtcDate.min;
		var second = startUtcDate.sec;

		this.DateTimeGrouping = range.dateTimeGrouping;

		switch(this.DateTimeGrouping)
		{
			case Asc.EDateTimeGroup.datetimegroupDay://day
			{
				this.Year = year;
				this.Month = month;
				this.Day = day;
				break;
			}
			case Asc.EDateTimeGroup.datetimegroupHour://hour
			{
				this.Year = year;
				this.Month = month;
				this.Day = day;
				this.Hour = hour;
				break;
			}
			case Asc.EDateTimeGroup.datetimegroupMinute://minute
			{
				this.Year = year;
				this.Month = month;
				this.Day = day;
				this.Hour = hour;
				this.Minute = minute;
				break;
			}
			case Asc.EDateTimeGroup.datetimegroupMonth://month
			{
				this.Year = year;
				this.Month = month;
				break;
			}
			case Asc.EDateTimeGroup.datetimegroupSecond://second
			{
				this.Year = year;
				this.Month = month;
				this.Day = day;
				this.Hour = hour;
				this.Minute = minute;
				this.Second = second;
				break;
			}
			case Asc.EDateTimeGroup.datetimegroupYear://year
			{
				this.Year = year;
				break;
			}
		}
	};
	DateGroupItem.prototype.readAttributes = function(attr, uq) {
		if (attr()) {
			var vals = attr();
			var val;
			val = vals["year"];
			if (undefined !== val) {
				this.Year = val - 0;
			}
			val = vals["month"];
			if (undefined !== val) {
				this.Month = val - 0;
			}
			val = vals["day"];
			if (undefined !== val) {
				this.Day = val - 0;
			}
			val = vals["hour"];
			if (undefined !== val) {
				this.Hour = val - 0;
			}
			val = vals["minute"];
			if (undefined !== val) {
				this.Minute = val - 0;
			}
			val = vals["second"];
			if (undefined !== val) {
				this.Second = val - 0;
			}
			val = vals["dateTimeGrouping"];
			if (undefined !== val) {
				val = AscCommonExcel.FromXml_ST_DateTimeGrouping(val);
				if (-1 !== val) {
					this.DateTimeGrouping = val;
				}
			}
		}
	};
	DateGroupItem.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*"dateGroupItem"*/);

		writer.WriteXmlNullableAttributeNumber("year", this.Year);
		writer.WriteXmlNullableAttributeNumber("month", this.Month);
		writer.WriteXmlNullableAttributeNumber("day", this.Day);
		writer.WriteXmlNullableAttributeNumber("hour", this.Hour);
		writer.WriteXmlNullableAttributeNumber("minute", this.Minute);
		writer.WriteXmlNullableAttributeNumber("second", this.Second);
		writer.WriteXmlAttributeStringEncode("dateTimeGrouping", AscCommonExcel.ToXml_ST_DateTimeGrouping(this.DateTimeGrouping));
		writer.WriteXmlAttributesEnd();

		writer.WriteXmlNodeEnd(ns + name);

	};

	var g_oCustomFilters = {
		And	 : 0,
		CustomFilters	: 1
	};
	/** @constructor */
	function CustomFilters() {
		this.Properties = g_oCustomFilters;

		this.And = false;
		this.CustomFilters = null;
	}
	CustomFilters.prototype.getType = function() {
		return UndoRedoDataTypes.CustomFilters;
	};
	CustomFilters.prototype.getProperties = function() {
		return this.Properties;
	};
	CustomFilters.prototype.getProperty = function(nType) {
		switch (nType) {
			case this.Properties.And: return this.And; break;
			case this.Properties.CustomFilters: return this.CustomFilters; break;
		}
		return null;
	};
	CustomFilters.prototype.setProperty = function(nType, value) {
		switch (nType) {
			case this.Properties.And: this.And = value;break;
			case this.Properties.CustomFilters: this.CustomFilters = value;break;
		}
	};
	CustomFilters.prototype.Write_ToBinary2 = function(writer) {
		writer.WriteBool(this.And);

		writer.WriteLong(this.CustomFilters ? this.CustomFilters.length : 0);
		for (var i = 0; i < this.CustomFilters.length; ++i) {
			this.CustomFilters[i].Write_ToBinary2(writer);
		}
	};
	CustomFilters.prototype.Read_FromBinary2 = function(reader) {
		this.And = reader.GetBool();
		var length = reader.GetLong();
		for (var i = 0; i < length; ++i) {
			var reply = new CustomFilter();
			reply.Read_FromBinary2(reader);
			if(!this.CustomFilters) {
				this.CustomFilters = [];
			}
			this.CustomFilters.push(reply);
		}
	};
	CustomFilters.prototype.clone = function() {
		var i, res = new CustomFilters();
		res.And = this.And;
		if (this.CustomFilters) {
			res.CustomFilters = [];
			for (i = 0; i < this.CustomFilters.length; ++i)
				res.CustomFilters.push(this.CustomFilters[i].clone());
		}
		return res;
	};
	CustomFilters.prototype.init = function(obj) {
		this.And = !obj.isChecked;
		this.CustomFilters = [];

		if(obj.filter1 != null)
			this.CustomFilters[0] = new CustomFilter(obj.filter1, obj.valFilter1);
		if(obj.filter2 != null)
			this.CustomFilters[1] = new CustomFilter(obj.filter2, obj.valFilter2);
	};
	CustomFilters.prototype.isHideValue = function(val, isLabelFilter, cell){

		var res = false;
		var filterRes1 = this.CustomFilters[0] ? this.CustomFilters[0].isHideValue(val, isLabelFilter, cell) : null;
		var filterRes2 = this.CustomFilters[1] ? this.CustomFilters[1].isHideValue(val, isLabelFilter, cell) : null;

		if(!this.And && ((filterRes1 === null && filterRes2 === true || filterRes1 === true && filterRes2 === null || filterRes1 === true && filterRes2 === true)))
			res = true;
		if(this.And && ((filterRes1 === true || filterRes2 === true)))
			res = true;

		return res;
	};
	CustomFilters.prototype.asc_getAnd = function () { return this.And; };
	CustomFilters.prototype.asc_getCustomFilters = function () { return this.CustomFilters; };

	CustomFilters.prototype.asc_setAnd = function (val) { this.And = val; };
	CustomFilters.prototype.asc_setCustomFilters = function (val) { this.CustomFilters = val; };

	CustomFilters.prototype.check = function () {
		if(this.CustomFilters) {
			for(var i = 0; i < this.CustomFilters.length; i++) {
				this.CustomFilters[i].check();
			}
		}
	};
	CustomFilters.prototype.correctFromInterface = function () {
		if(this.CustomFilters) {
			for(var i = 0; i < this.CustomFilters.length; i++) {
				this.CustomFilters[i].correctFromInterface();
			}
		}
	};

	CustomFilters.prototype._generateEmptyValueFilter = function() {
		this.And = true;
		this.CustomFilters = [];
		var customFilter = new CustomFilter();
		customFilter._generateEmptyValueFilter();
		this.CustomFilters.push(customFilter);
	};
	CustomFilters.prototype.readAttributes = function(attr, uq) {
		if (attr()) {
			var vals = attr();
			var val;
			val = vals["and"];
			if (undefined !== val) {
				this.And = AscCommon.getBoolFromXml(val);
			}
		}
	};
	CustomFilters.prototype.onStartNode = function(elem, attr, uq) {
		var newContext = this;
		if ("customFilter" === elem) {
			if (!this.CustomFilters) {
				this.CustomFilters = [];
			}
			newContext = new CustomFilter();
			if (newContext.readAttributes) {
				newContext.readAttributes(attr, uq);
			}
			this.CustomFilters.push(newContext);
		} else {
			newContext = null;
		}
		return newContext;
	};
	CustomFilters.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}
		if (!childns) {
			childns = "";
		}

		writer.WriteXmlNodeStart(ns + name);

		writer.WriteXmlNullableAttributeString("and", this.And ? 1 : null);
		writer.WriteXmlAttributesEnd();

		for (var i = 0; i < this.CustomFilters.length; ++i) {
			this.CustomFilters[i].toXml(writer, "customFilter", childns, childns);
		}

		writer.WriteXmlNodeEnd(ns + name);
	};
	CustomFilters.prototype.changeForInterface = function (filterTypes) {
		var res = this.clone();
		if(res.CustomFilters) {
			for(var i = 0; i < res.CustomFilters.length; i++) {
				res.CustomFilters[i].changeForInterface(filterTypes);
			}
		}
		return res;
	};

	var g_oCustomFilter = {
		Operator	 : 0,
		Val	: 1
	};

	/** @constructor */
	function CustomFilter(operator, val) {
		this.Properties = g_oCustomFilter;

		this.Operator = operator != undefined ? operator : c_oAscCustomAutoFilter.equals;
		this.Val = val != undefined ? val : null;
	}
	CustomFilter.prototype.getType = function() {
		return UndoRedoDataTypes.CustomFilter;
	};
	CustomFilter.prototype.getProperties = function() {
		return this.Properties;
	};
	CustomFilter.prototype.getProperty = function(nType) {
		switch (nType) {
			case this.Properties.Operator: return this.Operator; break;
			case this.Properties.Val: return this.Val; break;
		}
		return null;
	};
	CustomFilter.prototype.setProperty = function(nType, value) {
		switch (nType) {
			case this.Properties.Operator: this.Operator = value;break;
			case this.Properties.Val: this.Val = value;break;
		}
	};

	CustomFilter.prototype.clone = function() {
		var res = new CustomFilter();
		res.Operator = this.Operator;
		res.Val = this.Val;
		return res;
	};
	CustomFilter.prototype.init = function(operator, val) {
		this.Operator = operator;
		this.Val = val;
	};
	CustomFilter.prototype.isHideValue = function (val, isLabelFilter, cell) {

		var result = false;
		var isDigitValue = !isNaN(val);
		if (!isDigitValue) {
			val = val.toLowerCase();
		} else {
			let isQuotePrefix = cell && cell.getQuotePrefix();
			if (isQuotePrefix) {
				isDigitValue = false;
				val = val.toLowerCase();
			}
		}

		var checkComplexSymbols = null, filterVal;
		if (checkComplexSymbols != null) {
			result = checkComplexSymbols;
		} else {
			var isNumberFilter = this.Operator === c_oAscCustomAutoFilter.isGreaterThan || this.Operator === c_oAscCustomAutoFilter.isGreaterThanOrEqualTo || this.Operator === c_oAscCustomAutoFilter.isLessThan || this.Operator === c_oAscCustomAutoFilter.isLessThanOrEqualTo;

			if (isLabelFilter && !isNumberFilter) {
				isDigitValue = false;
			}
			if (c_oAscCustomAutoFilter.equals === this.Operator || c_oAscCustomAutoFilter.doesNotEqual === this.Operator) {
				filterVal = isNaN(this.Val) ? this.Val.toLowerCase() : this.Val;
			} else if (isNumberFilter) {
				filterVal = this.Val;
				if (isLabelFilter) {
					//'1'0' - число для данного случая
					filterVal = this.Val.replace(/'/g, "");
					if (isNaN(filterVal)) {
						filterVal = this.Val;
					}
				}

				if (isLabelFilter && isNaN(filterVal)) {
					//в случае label filter - сравнивам строки
					filterVal = filterVal.toLowerCase();
					isDigitValue = false;
					val = val.toLowerCase();
				} else if (isNaN(filterVal) && isNaN(val)) {
					filterVal = filterVal.toLowerCase();
				} else {
					filterVal = parseFloat(filterVal);
					val = parseFloat(val);
				}
			} else {
				filterVal = isNaN(this.Val) ? this.Val.toLowerCase() : this.Val;
			}

			if (cell && c_oAscCustomAutoFilter.equals === this.Operator) {
				val = cell.getValueWithFormat();
			}

			var trimVal = "string" === typeof(val) ? window["Asc"].trim(val) : val;
			var trimFilterVal = "string" === typeof(filterVal) ? window["Asc"].trim(filterVal) : filterVal;


			var matchingValues = function (_val1, _val2, op) {
				_val1 = _val1 + "";
				_val2 = _val2 + "";
				var matchingInfo = AscCommonExcel.matchingValue(new AscCommonExcel.cString(_val1));
				if (op) {
					matchingInfo.op = op;
				}
				return AscCommonExcel.matching(new AscCommonExcel.cString(_val2), matchingInfo);
			};

			switch (this.Operator) {
				case c_oAscCustomAutoFilter.equals://equals
				{
					if (!isDigitValue) {
						result = matchingValues(trimFilterVal, trimVal);
					} else if (trimVal === trimFilterVal) {
						result = true;
					}

					break;
				}
				case c_oAscCustomAutoFilter.doesNotEqual://doesNotEqual
				{
					if (!isDigitValue) {
						result = matchingValues(trimFilterVal, trimVal, "<>");
					} else if (trimVal !== trimFilterVal) {
						result = true;
					}

					break;
				}

				case c_oAscCustomAutoFilter.isGreaterThan://isGreaterThan
				{
					if (!isDigitValue) {
						result = matchingValues(trimFilterVal, trimVal, ">");
					} else if (val > filterVal) {
						result = true;
					}

					break;
				}
				case c_oAscCustomAutoFilter.isGreaterThanOrEqualTo://isGreaterThanOrEqualTo
				{
					if (!isDigitValue) {
						result = matchingValues(trimFilterVal, trimVal, ">=");
					} else if (val >= filterVal) {
						result = true;
					}

					break;
				}
				case c_oAscCustomAutoFilter.isLessThan://isLessThan
				{
					if (!isDigitValue) {
						result = matchingValues(trimFilterVal, trimVal, "<");
					} else if (val < filterVal) {
						result = true;
					}

					break;
				}
				case c_oAscCustomAutoFilter.isLessThanOrEqualTo://isLessThanOrEqualTo
				{
					if (!isDigitValue) {
						result = matchingValues(trimFilterVal, trimVal, "<=");
					} else if (val <= filterVal) {
						result = true;
					}

					break;
				}
				case c_oAscCustomAutoFilter.beginsWith://beginsWith
				{
					if (!isDigitValue) {
						result = matchingValues(trimFilterVal + "*", trimVal);
					}

					break;
				}
				case c_oAscCustomAutoFilter.doesNotBeginWith://doesNotBeginWith
				{
					if (!isDigitValue) {
						result = matchingValues(trimFilterVal + "*", trimVal, "<>");
					} else {
						result = true;
					}

					break;
				}
				case c_oAscCustomAutoFilter.endsWith://endsWith
				{
					if (!isDigitValue) {
						result = matchingValues("*" + trimFilterVal, trimVal);
					}

					break;
				}
				case c_oAscCustomAutoFilter.doesNotEndWith://doesNotEndWith
				{
					if (!isDigitValue) {
						result = matchingValues("*" + trimFilterVal, trimVal, "<>");
					} else {
						result = true;
					}

					break;
				}
				case c_oAscCustomAutoFilter.contains://contains
				{
					if (!isDigitValue) {
						result = matchingValues("*" + trimFilterVal + "*", trimVal);
					}

					break;
				}
				case c_oAscCustomAutoFilter.doesNotContain://doesNotContain
				{
					if (!isDigitValue) {
						result = matchingValues("*" + trimFilterVal + "*", trimVal, "<>");
					} else {
						result = true;
					}

					break;
				}
			}
		}

		return !result;
	};

	CustomFilter.prototype.asc_getOperator = function () { return this.Operator; };
	CustomFilter.prototype.asc_getVal = function () { return this.Val; };

	CustomFilter.prototype.asc_setOperator = function (val) { this.Operator = val; };
	CustomFilter.prototype.asc_setVal = function (val) {
		this.Val = val;
	};

	CustomFilter.prototype.check = function () {
		if(c_oAscCustomAutoFilter.doesNotEqual === this.Operator) {
			if("" === this.Val.replace(/ /g, "")){
				this.Val = " ";
			}
		}

		if (c_oAscCustomAutoFilter.beginsWith === this.Operator) {
			this.Operator = c_oAscCustomAutoFilter.equals;
			this.Val = this.Val + "*";
		} else if (c_oAscCustomAutoFilter.doesNotBeginWith === this.Operator) {
			this.Operator = c_oAscCustomAutoFilter.doesNotEqual;
			this.Val = this.Val + "*";
		} else if (c_oAscCustomAutoFilter.endsWith === this.Operator) {
			this.Operator = c_oAscCustomAutoFilter.equals;
			this.Val = "*" + this.Val;
		} else if (c_oAscCustomAutoFilter.doesNotEndWith === this.Operator) {
			this.Operator = c_oAscCustomAutoFilter.doesNotEqual;
			this.Val = "*" + this.Val;
		} else if (c_oAscCustomAutoFilter.contains === this.Operator) {
			this.Operator = c_oAscCustomAutoFilter.equals;
			this.Val = "*" + this.Val + "*";
		} else if (c_oAscCustomAutoFilter.doesNotContain === this.Operator) {
			this.Operator = c_oAscCustomAutoFilter.doesNotEqual;
			this.Val = "*" + this.Val + "*";
		}
	};
	CustomFilter.prototype.correctFromInterface = function () {
		if (this.isNumberCustomFilterType()) {
			if (!AscCommon.g_oFormatParser.isLocaleNumber(this.Val)) {
				//try parse format and save without format
				let resParse = AscCommon.g_oFormatParser.parse(this.Val);
				if (resParse) {
					this.Val = resParse.value + "";
				}
			}
		}
	};
	CustomFilter.prototype.isNumberCustomFilterType = function () {
		if (this.Operator === c_oAscCustomAutoFilter.isGreaterThan ||
			this.Operator === c_oAscCustomAutoFilter.isGreaterThanOrEqualTo ||
			this.Operator === c_oAscCustomAutoFilter.isLessThan ||
			this.Operator === c_oAscCustomAutoFilter.isLessThanOrEqualTo ||
			this.Operator === c_oAscCustomAutoFilter.doesNotEqual) {
			return true;
		}
		return false;
	};
	CustomFilter.prototype._generateEmptyValueFilter = function () {
		this.Operator = c_oAscCustomAutoFilter.doesNotEqual;
		this.Val = " ";
	};
	CustomFilter.prototype.readAttributes = function(attr, uq) {
		if (attr()) {
			var vals = attr();
			var val;
			val = vals["operator"];
			if (undefined !== val) {
				val = AscCommonExcel.FromXml_ST_FilterOperator(val);
				if (-1 !== val) {
					this.Operator = val;
				}
			}
			val = vals["val"];
			if (undefined !== val) {
				this.Val = AscCommon.unleakString(uq(val));
			}
		}
	};
	CustomFilter.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*"сustomFilter"*/);
		writer.WriteXmlAttributeString("operator", AscCommonExcel.ToXml_ST_FilterOperator(this.Operator));
		writer.WriteXmlAttributeStringEncode("val", this.Val);
		writer.WriteXmlAttributesEnd();

		writer.WriteXmlNodeEnd(ns + name);
	};
	CustomFilter.prototype.Write_ToBinary2 = function(writer) {
		if (null != this.Operator) {
			writer.WriteBool(true);
			writer.WriteLong(this.Operator);
		} else {
			writer.WriteBool(false);
		}

		if (null != this.Val) {
			writer.WriteBool(true);
			writer.WriteString2(this.Val);
		} else {
			writer.WriteBool(false);
		}
	};
	CustomFilter.prototype.Read_FromBinary2 = function(reader) {
		if (reader.GetBool()) {
			this.Operator = reader.GetLong();
		}
		if (reader.GetBool()) {
			this.Val = reader.GetString2();
		}
	};
	CustomFilter.prototype.changeForInterface = function(filterTypes) {
		if (!this.Val || this.Val.length <= 1) {
			return;
		}

		var isStartSpecSymbol = this.Val && this.Val.length > 1 && this.Val[0] === "*";
		var isEndSpecSymbol;
		if (!isStartSpecSymbol || (isStartSpecSymbol && this.Val.length >= 2)) {
			isEndSpecSymbol = this.Val && this.Val[this.Val.length - 1] === "*";
		}
		if (isStartSpecSymbol && isEndSpecSymbol && this.Val.length <= 2) {
			return;
		}
		if (isStartSpecSymbol || isEndSpecSymbol) {
			this.Val = this.Val.substring(isStartSpecSymbol ? 1 : 0, isEndSpecSymbol ? this.Val.length - 1 : this.Val.length);
			if(c_oAscCustomAutoFilter.doesNotEqual === this.Operator) {
				if (isStartSpecSymbol && isEndSpecSymbol) {
					this.Operator = c_oAscCustomAutoFilter.doesNotContain;
				} else if (isStartSpecSymbol) {
					this.Operator = c_oAscCustomAutoFilter.doesNotEndWith;
				} else {
					this.Operator = c_oAscCustomAutoFilter.doesNotBeginWith;
				}
			} else {
				if (isStartSpecSymbol && isEndSpecSymbol) {
					this.Operator = c_oAscCustomAutoFilter.contains;
				} else if (isStartSpecSymbol) {
					this.Operator = c_oAscCustomAutoFilter.endsWith;
				} else {
					this.Operator = c_oAscCustomAutoFilter.beginsWith;
				}
			}
		}

		//filterTypes
		/*autoFilterObject.asc_setIsTextFilter(filterTypes.text);
		autoFilterObject.asc_setIsDateFilter(filterTypes.date);
		autoFilterObject.asc_setColorsFill(filterTypes.colors);
		autoFilterObject.asc_setColorsFont(filterTypes.fontColors);*/
		if (filterTypes && filterTypes.date && AscCommon.g_oFormatParser.isLocaleNumber(this.Val + "")) {
			let api = Asc.editor || editor;
			this.Val = api.asc_getLocaleExample("m/d/yyyy", this.Val);
		}

	};

	var g_oDynamicFilter = {
		Type : 0,
		Val	: 1,
		MaxVal: 2
	};

	/** @constructor */
	function DynamicFilter() {
		this.Properties = g_oDynamicFilter;

		this.Type = null;
		this.Val = null;
		this.MaxVal = null;
	}
	DynamicFilter.prototype.getType = function() {
		return UndoRedoDataTypes.DynamicFilter;
	};
	DynamicFilter.prototype.getProperties = function() {
		return this.Properties;
	};
	DynamicFilter.prototype.getProperty = function(nType) {
		switch (nType) {
			case this.Properties.Type: return this.Type;
			case this.Properties.Val: return this.Val;
			case this.Properties.MaxVal: return this.MaxVal;
		}
		return null;
	};
	DynamicFilter.prototype.setProperty = function(nType, value) {
		switch (nType) {
			case this.Properties.Type: this.Type = value;break;
			case this.Properties.Val: this.Val = value;break;
			case this.Properties.MaxVal: this.MaxVal = value;break;
		}
	};
	DynamicFilter.prototype.clone = function() {
		var res = new DynamicFilter();
		res.Type = this.Type;
		res.Val = this.Val;
		res.MaxVal = this.MaxVal;
		return res;
	};

	DynamicFilter.prototype.init = function(range) {
		let val = null, maxVal;

		/*if (range) {
			for (let i in Asc.c_oAscDynamicAutoFilter) {
				if (i !== "aboveAverage" && i !== "belowAverage") {
					this.Type = Asc.c_oAscDynamicAutoFilter[i];
					this.init();
				}
			}
		}*/

		//by test:
		//today -> 45049
		let referenceValues = {};
		referenceValues[Asc.c_oAscDynamicAutoFilter.lastMonth] = [45017, 45047];
		referenceValues[Asc.c_oAscDynamicAutoFilter.lastQuarter] = [44927, 45017];
		referenceValues[Asc.c_oAscDynamicAutoFilter.lastWeek] = [45039, 45046];
		referenceValues[Asc.c_oAscDynamicAutoFilter.lastYear] = [44562, 44927];

		referenceValues[Asc.c_oAscDynamicAutoFilter.nextMonth] = [45078, 45108];
		referenceValues[Asc.c_oAscDynamicAutoFilter.nextQuarter] = [45108, 45200];
		referenceValues[Asc.c_oAscDynamicAutoFilter.nextWeek] = [45053, 45060];
		referenceValues[Asc.c_oAscDynamicAutoFilter.nextYear] = [45292, 45658];

		referenceValues[Asc.c_oAscDynamicAutoFilter.thisMonth] = [45047, 45078];
		referenceValues[Asc.c_oAscDynamicAutoFilter.thisQuarter] = [45017, 45108];
		referenceValues[Asc.c_oAscDynamicAutoFilter.thisWeek] = [45046, 45053];
		referenceValues[Asc.c_oAscDynamicAutoFilter.thisYear] = [44927, 45292];

		referenceValues[Asc.c_oAscDynamicAutoFilter.today] = [45049, 45050];
		referenceValues[Asc.c_oAscDynamicAutoFilter.tomorrow] = [45050, 45051];
		referenceValues[Asc.c_oAscDynamicAutoFilter.yearToDate] = [44927, 45050];
		referenceValues[Asc.c_oAscDynamicAutoFilter.yesterday] = [45048, 45049];


		let _calcDateInterval = function (_diff) {
			let today = new cDate().getCurrentDate();
			let start, end;

			if (isSimpleWeek) {
				start = new cDate(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate() - today.getUTCDay() + _diff * 7);
				//+1 -> ms excel
				end = new cDate(start.getUTCFullYear(), start.getUTCMonth(), start.getUTCDate() + 7 + 1);
			} else if (isSimpleMonth) {
				start = new cDate(today.getUTCFullYear(), today.getUTCMonth() + _diff, 1);
				//+1 -> ms excel
				end = new cDate(start.getUTCFullYear(), start.getUTCMonth() + 1 + 1, 1);
			} else if (isSimpleQuarter) {
				let quarter = Math.floor((today.getUTCMonth() / 3));
				start = new cDate(today.getUTCFullYear(), quarter * 3 + _diff * 3, 1);
				end = new cDate(start.getUTCFullYear(), start.getUTCMonth() + 3 + 1, 1);
			} else if (isSimpleYear) {
				start = new cDate(today.getUTCFullYear() + _diff, 0, 1);
				//+1 -> ms excel
				end = new cDate(start.getUTCFullYear() + 1, 11, 31 + 1);
			}

			val = start.getExcelDate(true);
			maxVal = end.getExcelDate(true);
		};

		let oTypes = Asc.c_oAscDynamicAutoFilter;
		let isSimpleWeek = this.Type === oTypes.lastWeek || this.Type === oTypes.nextWeek || this.Type === oTypes.thisWeek;
		let isSimpleMonth = this.Type === oTypes.lastMonth || this.Type === oTypes.nextMonth || this.Type === oTypes.thisMonth;
		let isSimpleQuarter = this.Type === oTypes.lastQuarter || this.Type === oTypes.nextQuarter || this.Type === oTypes.thisQuarter;
		let isSimpleYear = this.Type === oTypes.lastYear || this.Type === oTypes.nextYear || this.Type === oTypes.thisYear;

		switch (this.Type) {
			case Asc.c_oAscDynamicAutoFilter.aboveAverage:
			case Asc.c_oAscDynamicAutoFilter.belowAverage: {
				let sum = 0;
				let counter = 0;

				range._foreachNoEmpty(function (cell) {
					let cellVal = parseFloat(cell.getValueWithoutFormat());

					if (!isNaN(cellVal)) {
						sum += parseFloat(cellVal);
						counter++;
					}

				});
				val = sum / counter;

				break;
			}
			case Asc.c_oAscDynamicAutoFilter.lastMonth:
			case Asc.c_oAscDynamicAutoFilter.lastQuarter:
			case Asc.c_oAscDynamicAutoFilter.lastWeek:
			case Asc.c_oAscDynamicAutoFilter.lastYear: {
				_calcDateInterval(-1);
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.nextMonth:
			case Asc.c_oAscDynamicAutoFilter.nextQuarter:
			case Asc.c_oAscDynamicAutoFilter.nextWeek:
			case Asc.c_oAscDynamicAutoFilter.nextYear: {
				_calcDateInterval(1);
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.thisMonth:
			case Asc.c_oAscDynamicAutoFilter.thisQuarter:
			case Asc.c_oAscDynamicAutoFilter.thisWeek:
			case Asc.c_oAscDynamicAutoFilter.thisYear: {
				_calcDateInterval(0);
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.nullType: {
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.today: {
				let today = new cDate().getCurrentDate();
				val = today.getExcelDate(true);
				today.addDays(1);
				maxVal = today.getExcelDate(true);
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.tomorrow: {
				let today = new cDate().getCurrentDate();
				today.addDays(1);
				val = today.getExcelDate(true);
				today.addDays(1);
				maxVal = today.getExcelDate(true);
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.yearToDate: {
				let today = new cDate().getCurrentDate();
				val = new cDate(today.getUTCFullYear(), 0, 1).getExcelDate(true);
				today.addDays(1);
				maxVal = today.getExcelDate(true);
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.yesterday: {
				let today = new cDate().getCurrentDate();
				maxVal = today.getExcelDate(true);
				today.addDays(-1);
				val = today.getExcelDate(true);
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m1: {
				//<dynamicFilter type="m1"/>
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m10: {
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m11: {
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m12: {
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m2: {
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m3: {
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m4: {
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m5: {
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m6: {
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m7: {
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m8: {
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m9: {
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.q1: {
				//<dynamicFilter type="Q2"/>
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.q2: {
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.q3: {
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.q4: {
				break;
			}
		}

		/*if (val != undefined && referenceValues[this.Type] && (val !== referenceValues[this.Type][0] || maxVal !== referenceValues[this.Type][1])) {
			console.log(this.Type);
		}*/


		this.Val = val;
		this.MaxVal = maxVal;
	};

	DynamicFilter.prototype.isHideValue = function (val) {
		let res = false;
		let minMonth, maxMonth;
		switch (this.Type) {
			case Asc.c_oAscDynamicAutoFilter.aboveAverage: {
				res = !(val > this.Val);
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.belowAverage: {
				res = !(val < this.Val);
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.lastMonth:
			case Asc.c_oAscDynamicAutoFilter.lastQuarter:
			case Asc.c_oAscDynamicAutoFilter.lastWeek:
			case Asc.c_oAscDynamicAutoFilter.lastYear:
			case Asc.c_oAscDynamicAutoFilter.nextMonth:
			case Asc.c_oAscDynamicAutoFilter.nextQuarter:
			case Asc.c_oAscDynamicAutoFilter.nextWeek:
			case Asc.c_oAscDynamicAutoFilter.nextYear:
			case Asc.c_oAscDynamicAutoFilter.thisMonth:
			case Asc.c_oAscDynamicAutoFilter.thisQuarter:
			case Asc.c_oAscDynamicAutoFilter.thisWeek:
			case Asc.c_oAscDynamicAutoFilter.today:
			case Asc.c_oAscDynamicAutoFilter.tomorrow:
			case Asc.c_oAscDynamicAutoFilter.yearToDate:
			case Asc.c_oAscDynamicAutoFilter.yesterday:
			case Asc.c_oAscDynamicAutoFilter.thisYear: {
				res = !(val >= this.Val && val < this.MaxVal);
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.nullType: {
				break;
			}

			case Asc.c_oAscDynamicAutoFilter.m1: {
				//<dynamicFilter type="m1"/>
				//Shows the dates that are in January, regardless of year
				minMonth = maxMonth = 0;

				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m10: {
				//Shows the dates that are in October, regardless of year
				minMonth = maxMonth = 9;
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m11: {
				//Shows the dates that are in November, regardless of year
				minMonth = maxMonth = 10;
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m12: {
				//Shows the dates that are in December, regardless of year
				minMonth = maxMonth = 11;
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m2: {
				//Shows the dates that are in Februray, regardless of year
				minMonth = maxMonth = 1;
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m3: {
				//Shows the dates that are in March, regardless of year
				minMonth = maxMonth = 2;
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m4: {
				//Shows the dates that are in April, regardless of year
				minMonth = maxMonth = 3;
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m5: {
				//Shows the dates that are in May, regardless of year
				minMonth = maxMonth = 4;
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m6: {
				//Shows the dates that are in June, regardless of year
				minMonth = maxMonth = 5;
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m7: {
				//Shows the dates that are in July, regardless of year
				minMonth = maxMonth = 6;
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m8: {
				//Shows the dates that are in August, regardless of year
				minMonth = maxMonth = 7;
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.m9: {
				//Shows the dates that are in September, regardless of year
				minMonth = maxMonth = 8;
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.q1: {
				//<dynamicFilter type="Q2"/>
				//Shows the dates that are in the 1st quarter, regardless of year
				minMonth = 0;
				maxMonth = 2;
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.q2: {
				//Shows the dates that are in the 2st quarter, regardless of year
				minMonth = 3;
				maxMonth = 5;
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.q3: {
				//Shows the dates that are in the 3st quarter, regardless of year
				minMonth = 6;
				maxMonth = 8;
				break;
			}
			case Asc.c_oAscDynamicAutoFilter.q4: {
				//Shows the dates that are in the 4st quarter, regardless of year
				minMonth = 9;
				maxMonth = 11;
				break;
			}
		}

		if (minMonth !== undefined) {
			let valToDate = new cDate().getDateFromExcel(val);
			let valMonth = valToDate.getUTCMonth();
			res = !(valMonth >= minMonth && valMonth <= maxMonth);
		}

		return res;
	};

	DynamicFilter.prototype.asc_getType = function () { return this.Type; };
	DynamicFilter.prototype.asc_getVal = function () { return this.Val; };
	DynamicFilter.prototype.asc_getMaxVal = function () { return this.MaxVal; };

	DynamicFilter.prototype.asc_setType = function (val) {
		this.Type = val; };
	DynamicFilter.prototype.asc_setVal = function (val) { this.Val = val; };
	DynamicFilter.prototype.asc_setMaxVal = function (val) { this.MaxVal = val; };
	DynamicFilter.prototype.readAttributes = function(attr, uq) {
		if (attr()) {
			var vals = attr();
			var val;
			val = vals["type"];
			if (undefined !== val) {
				val = AscCommonExcel.FromXml_ST_DynamicFilterType(val);
				if (-1 !== val) {
					this.Type = val;
				}
			}
			val = vals["val"];
			if (undefined !== val) {
				this.Val = val - 0;
			}
			val = vals["maxVal"];
			if (undefined !== val) {
				this.MaxVal = val - 0;
			}
		}
	};
	DynamicFilter.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*"dynamicFilter"*/);

		writer.WriteXmlAttributeString("type", AscCommonExcel.ToXml_ST_DynamicFilterType(this.Type));
		writer.WriteXmlNullableAttributeNumber("val", this.val);
		writer.WriteXmlNullableAttributeNumber("maxVal", this.MaxVal);

		writer.WriteXmlAttributesEnd(true);
	};
	DynamicFilter.prototype.Write_ToBinary2 = function(writer) {
		if (null !== this.Type) {
			writer.WriteBool(true);
			writer.WriteLong(this.Type);
		} else {
			writer.WriteBool(false);
		}

		if (null !== this.Val) {
			writer.WriteBool(true);
			writer.WriteLong(this.Val);
		} else {
			writer.WriteBool(false);
		}

		if (null !== this.MaxVal) {
			writer.WriteBool(true);
			writer.WriteLong(this.MaxVal);
		} else {
			writer.WriteBool(false);
		}
	};
	DynamicFilter.prototype.Read_FromBinary2 = function(reader) {
		if (reader.GetBool()) {
			this.Type = reader.GetLong();
		}
		if (reader.GetBool()) {
			this.Val = reader.GetLong();
		}
		if (reader.GetBool()) {
			this.MaxVal = reader.GetLong();
		}
	};


	var g_oColorFilter = {
		CellColor : 0,
		dxf	: 1
	};

	/** @constructor */
	function ColorFilter() {
		this.Properties = g_oColorFilter;

		this.CellColor = null;
		this.dxf = null;
	}
	ColorFilter.prototype.getType = function() {
		return UndoRedoDataTypes.ColorFilter;
	};
	ColorFilter.prototype.getProperties = function() {
		return this.Properties;
	};
	ColorFilter.prototype.getProperty = function(nType) {
		switch (nType) {
			case this.Properties.CellColor: return this.CellColor; break;
			case this.Properties.dxf: return this.dxf; break;
		}
		return null;
	};
	ColorFilter.prototype.setProperty = function(nType, value) {
		switch (nType) {
			case this.Properties.CellColor: this.CellColor = value;break;
			case this.Properties.dxf: this.dxf = value;break;
		}
	};
	ColorFilter.prototype.clone = function() {
		var res = new ColorFilter();
		res.CellColor = this.CellColor;
		if (this.dxf) {
			res.dxf = this.dxf.clone();
		}
		return res;
	};
	ColorFilter.prototype.isHideValue = function(cell) {

		var res = true;
		var t = this;

		var isEqualColors = function(filterColor, cellColor)
		{
			var res = false;
			if(filterColor === cellColor)
			{
				res = true;
			}
			else if(!filterColor && (!cellColor || null === cellColor.rgb || 0 === cellColor.rgb))
			{
				res = true;
			}
			else if(!cellColor && (!filterColor || null === filterColor.rgb || 0 === filterColor.rgb))
			{
				res = true;
			}
			else if(cellColor && filterColor && cellColor.rgb === filterColor.rgb)
			{
				res = true;
			}

			return res;
		};

		if(this.dxf && this.dxf.fill && cell)
		{
			var filterColor = this.dxf.fill.bg();
			cell.getLeftTopCellNoEmpty(function(cell) {
				var fontColor;
				var xfs = cell ? cell.getCompiledStyleCustom(false, true, true) : null;
				if(false === t.CellColor)//font color
				{
					var multiText;
					if(cell && (multiText = cell.getValueMultiText()) !== null)
					{
						for(var j = 0; j < multiText.length; j++)
						{
							fontColor = multiText[j].format ? multiText[j].format.getColor() : null;
							if(null === fontColor) {
								fontColor = xfs && xfs.font ? xfs.font.getColor() : null;
							}
							if(isEqualColors(filterColor,fontColor ))
							{
								res = false;
								break;
							}
						}
					}
					else
					{
						fontColor = xfs && xfs.font ? xfs.font.getColor() : null;
						if(isEqualColors(filterColor,fontColor))
						{
							res = false;
						}
					}
				}
				else
				{
					var cellColor =  xfs !== null && xfs.fill && xfs.fill.bg ? xfs.fill.bg() : null;

					if(isEqualColors(filterColor, cellColor))
					{
						res = false;
					}
				}
			});
		}

		return res;
	};

	ColorFilter.prototype.asc_getCellColor = function () { return this.CellColor; };
	ColorFilter.prototype.asc_getDxf = function () { return this.dxf; };

	ColorFilter.prototype.asc_setCellColor = function (val) { this.CellColor = val; };
	ColorFilter.prototype.asc_setDxf = function (val) { this.dxf = val; };
	ColorFilter.prototype.asc_getCColor = function ()
	{
		var res = null;

		if(this.dxf && this.dxf.fill && null !== this.dxf.fill.bg() && null !== this.dxf.fill.bg().rgb)
		{
			var color = this.dxf.fill.bg();

			var res = new Asc.asc_CColor();
			res.asc_putR(color.getR());
			res.asc_putG(color.getG());
			res.asc_putB(color.getB());
			res.asc_putA(color.getA());
		}

		return res;
	};
	ColorFilter.prototype.asc_setCColor = function (asc_CColor)
	{
		if(!this.dxf)
		{
			this.dxf = new CellXfs();
		}

		if(!this.dxf.fill)
		{
			this.dxf.fill = new Fill();
		}

		if(null === asc_CColor)
		{
			this.dxf.fill.fromColor(null);
		}
		else
		{
			this.dxf.fill.fromColor(new RgbColor((asc_CColor.asc_getR() << 16) + (asc_CColor.asc_getG() << 8) + asc_CColor.asc_getB()));
		}
	};
	ColorFilter.prototype.readAttributes = function(attr, uq) {
		if (attr()) {
			var vals = attr();
			var val;
			//todo
			// val = vals["dxfId"];
			// if (undefined !== val) {
			// 	this.dxfId = val - 0;
			// }
			val = vals["cellColor"];
			if (undefined !== val) {
				this.CellColor = AscCommon.getBoolFromXml(val);
			}
		}
	};
	ColorFilter.prototype.toXml = function (writer, name, ns, childns) {
		if (!ns) {
			ns = "";
		}

		writer.WriteXmlNodeStart(ns + name/*colorFilter*/);
		if (this.dxf != null) {
			writer.WriteXmlNullableAttributeNumber("dxfId", writer.context.InitSaveManager.aDxfs.length);
			writer.context.InitSaveManager.aDxfs.push(this.dxf);
		}

		if (this.CellColor === false) {
			writer.WriteXmlNullableAttributeNumber("cellColor", 0);
		}

		writer.WriteXmlAttributesEnd(true);
	};
	ColorFilter.prototype.Write_ToBinary2 = function(writer) {
		if (null !== this.CellColor) {
			writer.WriteBool(true);
			writer.WriteBool(this.CellColor);
		} else {
			writer.WriteBool(false);
		}

		if(null != this.dxf) {
			var dxf = this.dxf;
			writer.WriteBool(true);
			var oBinaryStylesTableWriter = new AscCommonExcel.BinaryStylesTableWriter(writer);
			oBinaryStylesTableWriter.bs.WriteItem(0, function(){oBinaryStylesTableWriter.WriteDxf(dxf);});
		}else {
			writer.WriteBool(false);
		}
	};
	ColorFilter.prototype.Read_FromBinary2 = function(reader) {
		if (reader.GetBool()) {
			this.CellColor = reader.GetBool();
		}
		if (reader.GetBool()) {
			var api_sheet = Asc['editor'];
			var wb = api_sheet.wbModel;
			var bsr = new AscCommonExcel.Binary_StylesTableReader(reader, wb);
			var bcr = new AscCommon.Binary_CommonReader(reader);
			var oDxf = new AscCommonExcel.CellXfs();
			reader.GetUChar();
			var length = reader.GetULongLE();
			bcr.Read1(length, function(t,l){
				return bsr.ReadDxf(t,l,oDxf);
			});
			this.dxf = oDxf;
		}
	};


	var g_oTop10 = {
		FilterVal : 0,
		Percent	: 1,
		Top: 2,
		Val: 3
	};

	/** @constructor */
	function Top10() {
		this.Properties = g_oTop10;

		this.FilterVal = null;
		this.Percent = false;
		this.Top = true;
		this.Val = null;
	}
	Top10.prototype.getType = function() {
		return UndoRedoDataTypes.Top10;
	};
	Top10.prototype.getProperties = function() {
		return this.Properties;
	};
	Top10.prototype.getProperty = function(nType) {
		switch (nType) {
			case this.Properties.FilterVal: return this.FilterVal; break;
			case this.Properties.Percent: return this.Percent; break;
			case this.Properties.Top: return this.Top; break;
			case this.Properties.Val: return this.Val; break;
		}
		return null;
	};
	Top10.prototype.setProperty = function(nType, value) {
		switch (nType) {
			case this.Properties.FilterVal: this.FilterVal = value;break;
			case this.Properties.Percent: this.Percent = value;break;
			case this.Properties.Top: this.Top = value;break;
			case this.Properties.Val: this.Val = value;break;
		}
	};
	Top10.prototype.clone = function() {
		var res = new Top10();
		res.FilterVal = this.FilterVal;
		res.Percent = this.Percent;
		res.Top = this.Top;
		res.Val = this.Val;
		return res;
	};
	Top10.prototype.isHideValue = function(val) {
		// ToDo работает не совсем правильно.
		var res = false;

		if(null !== this.FilterVal)
		{
			if(this.Top)
			{
				if(val < this.FilterVal)
				{
					res = true;
				}
			}
			else
			{
				if(val > this.FilterVal)
				{
					res = true;
				}
			}
		}

		return res;
	};

	Top10.prototype.init = function(range, reWrite){
		var t = this;

		if(null === this.FilterVal || true === reWrite)
		{
			if(range)
			{
				var arr = [];
				var alreadyAddValues = {};
				var count = 0;
				range._setPropertyNoEmpty(null, null, function(cell){
					var val = parseFloat(cell.getValueWithoutFormat());

					if(!isNaN(val) && !alreadyAddValues[val])
					{
						arr.push(val);
						alreadyAddValues[val] = 1;
						count++;
					}
				});
				this.initByArray(arr);
			}
		}
	};
	Top10.prototype.initByArray = function(arr, isSum){
		var res = null;
		var t = this;
		if(arr && arr.length)
		{
			arr.sort(function(a, b){
				var res;
				if(t.Top)
				{
					res = b - a;
				}
				else
				{
					res = a - b;
				}

				return res;
			});

			if(this.Percent)
			{
				var num = parseInt(arr.length * (this.Val / 100));
				if(0 === num)
				{
					num = 1;
				}

				res = arr[num - 1];
			}
			else
			{
				if (isSum) {
					var index = 0;
					var sum = res = arr[index++];
					while (index < arr.length && sum < this.Val) {
						res = arr[index++];
						sum += res;
					}
				} else {
					res = arr[this.Val - 1];
				}
			}
		}
		if(null != res)
		{
			this.FilterVal = res;
		}
	};

	Top10.prototype.asc_getFilterVal = function () { return this.FilterVal; };
	Top10.prototype.asc_getPercent = function () { return this.Percent; };
	Top10.prototype.asc_getTop = function () { return this.Top; };
	Top10.prototype.asc_getVal = function () { return this.Val; };

	Top10.prototype.asc_setFilterVal = function (val) { this.FilterVal = val; };
	Top10.prototype.asc_setPercent = function (val) { this.Percent = val; };
	Top10.prototype.asc_setTop = function (val) { this.Top = val; };
	Top10.prototype.asc_setVal = function (val) { this.Val = val; };
	Top10.prototype.readAttributes = function(attr, uq) {
		if (attr()) {
			var vals = attr();
			var val;
			val = vals["top"];
			if (undefined !== val) {
				this.Top = AscCommon.getBoolFromXml(val);
			}
			val = vals["percent"];
			if (undefined !== val) {
				this.Percent = AscCommon.getBoolFromXml(val);
			}
			val = vals["val"];
			if (undefined !== val) {
				this.Val = val - 0;
			}
			val = vals["filterVal"];
			if (undefined !== val) {
				this.FilterVal = val - 0;
			}
		}
	};
	Top10.prototype.toXml = function (writer, name, ns) {
		if (!ns) {
			ns = "";
		}
		writer.WriteXmlNodeStart(ns + name);
		if (true !== this.Top) {
			writer.WriteXmlAttributeBool("top", this.Top);
		}
		if (false !== this.Percent) {
			writer.WriteXmlAttributeBool("percent", this.Percent);
		}
		if (null !== this.Val) {
			writer.WriteXmlAttributeNumber("val", this.Val);
		}
		if (null !== this.FilterVal) {
			writer.WriteXmlAttributeNumber("filterVal", this.FilterVal);
		}
		writer.WriteXmlAttributesEnd(true);
	};
	Top10.prototype.Write_ToBinary2 = function(w) {
		if (null !== this.FilterVal) {
			w.WriteBool(true);
			w.WriteLong(this.FilterVal);
		} else {
			w.WriteBool(false);
		}

		if (null !== this.Percent) {
			w.WriteBool(true);
			w.WriteLong(this.Percent);
		} else {
			w.WriteBool(false);
		}

		w.WriteBool(this.Top);

		if (null != this.Val) {
			w.WriteBool(true);
			w.WriteLong(this.Val);
		} else {
			w.WriteBool(false);
		}
	};
	Top10.prototype.Read_FromBinary2 = function(reader) {
		if (reader.GetBool()) {
			this.FilterVal = reader.GetLong();
		}
		if (reader.GetBool()) {
			this.Percent = reader.GetLong();
		}
		this.Top = reader.GetBool();
		if (reader.GetBool()) {
			this.Val = reader.GetLong();
		}
	};


	/** @constructor */
	function SortCondition() {
		this.Ref = null;
		this.ConditionSortBy = null;
		this.ConditionDescending = null;
		this.dxf = null;

		this._hasHeaders = null;
	}
	SortCondition.prototype.clone = function() {
		var res = new SortCondition();
		res.Ref = this.Ref ? this.Ref.clone() : null;
		res.ConditionSortBy = this.ConditionSortBy;
		res.ConditionDescending = this.ConditionDescending;
		if (this.dxf)
			res.dxf = this.dxf.clone();
		return res;
	};
	SortCondition.prototype.Read_FromBinary2 = function(r) {
		if (r.GetBool()) {
			var r1 = r.GetLong();
			var c1 = r.GetLong();
			var r2 = r.GetLong();
			var c2 = r.GetLong();

			this.Ref = new Asc.Range(c1, r1, c2, r2);
		}
		if (r.GetBool()) {
			this.ConditionSortBy = r.GetLong();
		}
		if (r.GetBool()) {
			this.ConditionDescending = r.GetBool();
		}

		if (r.GetBool()) {
			var api_sheet = Asc['editor'];
			var wb = api_sheet.wbModel;
			var bsr = new AscCommonExcel.Binary_StylesTableReader(r, wb);
			var bcr = new AscCommon.Binary_CommonReader(r);
			var oDxf = new AscCommonExcel.CellXfs();
			r.GetUChar();
			var length = r.GetULongLE();
			bcr.Read1(length, function(t,l){
				return bsr.ReadDxf(t,l,oDxf);
			});
			this.dxf = oDxf;
		}
	};
	SortCondition.prototype.Write_ToBinary2 = function(w) {
		if (null != this.Ref) {
			w.WriteBool(true);
			w.WriteLong(this.Ref.r1);
			w.WriteLong(this.Ref.c1);
			w.WriteLong(this.Ref.r2);
			w.WriteLong(this.Ref.c2);
		} else {
			w.WriteBool(false);
		}
		if (null != this.ConditionSortBy) {
			w.WriteBool(true);
			w.WriteLong(this.ConditionSortBy);
		} else {
			w.WriteBool(false);
		}
		if (null != this.ConditionDescending) {
			w.WriteBool(true);
			w.WriteBool(this.ConditionDescending);
		} else {
			w.WriteBool(false);
		}

		if(null != this.dxf) {
			var dxf = this.dxf;
			w.WriteBool(true);
			var oBinaryStylesTableWriter = new AscCommonExcel.BinaryStylesTableWriter(w);
			oBinaryStylesTableWriter.bs.WriteItem(0, function(){oBinaryStylesTableWriter.WriteDxf(dxf);});
		}else {
			w.WriteBool(false);
		}
	};
	SortCondition.prototype.moveRef = function(col, row) {
		var ref = this.Ref.clone();
		ref.setOffset(new AscCommon.CellBase(row || 0, col || 0));

		this.Ref = ref;
	};
	SortCondition.prototype.changeColumns = function(activeRange, isDelete) {
		var bIsDeleteCurSortCondition = false;
		var ref = this.Ref.clone();
		var offsetCol = null;

		if(isDelete)
		{
			if(activeRange.c1 <= ref.c1 && activeRange.c2 >= ref.c1)
			{
				bIsDeleteCurSortCondition = true;
			}
			else if(activeRange.c1 < ref.c1)
			{
				offsetCol = -(activeRange.c2 - activeRange.c1 + 1);
			}
		}
		else
		{
			if(activeRange.c1 <= ref.c1)
			{
				offsetCol = activeRange.c2 - activeRange.c1 + 1;
			}
		}

		if(null !== offsetCol)
		{
			ref.setOffset(new AscCommon.CellBase(0, offsetCol));
			this.Ref = ref;
		}

		return bIsDeleteCurSortCondition;
	};

	SortCondition.prototype.setOffset = function(offset) {
		var ref = this.Ref.clone();
		ref.setOffset(offset);
		this.Ref = ref;
	};

	SortCondition.prototype.getSortType = function() {
		var res = null;

		if(true === this.ConditionDescending) {
			res = Asc.c_oAscSortOptions.Ascending;
		} else if(false === this.ConditionDescending) {
			res = Asc.c_oAscSortOptions.Descending;
		} else if(Asc.ESortBy.sortbyCellColor === this.ConditionSortBy) {
			res = Asc.c_oAscSortOptions.ByColorFill;
		} else if(Asc.ESortBy.sortbyCellColor === this.sortbyFontColor) {
			res = Asc.c_oAscSortOptions.ByColorFont;
		}

		return res;
	};

	SortCondition.prototype.getSortColor = function() {
		var res = null;

		if(this.dxf) {
			if(this.dxf.fill && this.dxf.fill.notEmpty()) {
				res = this.dxf.fill.bg();
			} else if(this.dxf.font && this.dxf.font.c) {
				res = this.dxf.font.c;
			}
		}

		return res;
	};

	SortCondition.prototype.applySort = function(type, ref, color) {
		this.Ref = ref;

		if(type === Asc.c_oAscSortOptions.ByColorFill || type === Asc.c_oAscSortOptions.ByColorFont) {
			var newDxf;
			if (type === Asc.c_oAscSortOptions.ByColorFill) {
				newDxf = new AscCommonExcel.CellXfs();
				newDxf.fill = new AscCommonExcel.Fill();
				newDxf.fill.fromColor(color);
				this.ConditionSortBy = Asc.ESortBy.sortbyCellColor;
			} else {
				newDxf.font = new AscCommonExcel.Font();
				newDxf.font.setColor(color);
				this.ConditionSortBy = Asc.ESortBy.sortbyFontColor;
			}

			this.dxf = AscCommonExcel.g_StyleCache.addXf(newDxf);
		} else if(type === Asc.c_oAscSortOptions.Ascending || type === Asc.c_oAscSortOptions.Descending) {
			this.ConditionDescending = type !== Asc.c_oAscSortOptions.Ascending;
		}

	};

	SortCondition.prototype.shift = function(range, offset, bColumnSort) {
		var from = this.Ref;
		var to = null;
		var bAdd = offset.row > 0 || offset.col > 0;
		var bHor = 0 != offset.col;
		var nTemp1, nTemp2;
		var diff = bHor ? range.c1 + offset.col - 1 : range.r1 + offset.row;
		if (bHor && bColumnSort) {
			if (from.c1 < range.c1 && range.r1 <= from.r1 && from.r2 <= range.r2) {
				if (bAdd) {
					to = from.clone();
					to.setOffsetLast(new AscCommon.CellBase(0, range.c2 - range.c1 + 1));
				} else {
					to = from.clone();
					nTemp1 = from.c2 - range.c1 + 1;
					nTemp2 = range.c2 - range.c1 + 1;
					to.setOffsetLast(new AscCommon.CellBase(0, -Math.min(nTemp1, nTemp2)));
				}
			}
		} else if(!bColumnSort) {
			if (from.r1 < range.r1 && range.c1 <= from.c1 && from.c2 <= range.c2) {
				if (bAdd) {
					to = from.clone();
					to.setOffsetLast(new AscCommon.CellBase(range.r2 - range.r1 + 1, 0));
				} else {
					to = from.clone();
					nTemp1 = from.r2 - range.r1 + 1;
					nTemp2 = range.r2 - range.r1 + 1;
					to.setOffsetLast(new AscCommon.CellBase(-Math.min(nTemp1, nTemp2), 0));
				}
			}
		}
		if(null != to) {
			this.Ref = to;
		}
	};

	function AutoFilterDateElem(start, end, dateTimeGrouping) {
		this.start = start;
		this.end = end;
		this.dateTimeGrouping = dateTimeGrouping;
	}
	AutoFilterDateElem.prototype.Write_ToBinary2 = function(w) {
		w.WriteLong(this.start);
		w.WriteLong(this.end);
		w.WriteLong(this.dateTimeGrouping);
	};
	AutoFilterDateElem.prototype.Read_FromBinary2 = function(r) {
		this.start = r.GetLong();
		this.end = r.GetLong();
		this.dateTimeGrouping = r.GetLong();
	};
	AutoFilterDateElem.prototype.clone = function() {
		var res = new AutoFilterDateElem();
		res.start = this.start;
		res.end = this.end;
		res.dateTimeGrouping = this.dateTimeGrouping;

		return res;
	};
	AutoFilterDateElem.prototype.convertDateGroupItemToRange = function(oDateGroupItem) {
		var startDate, endDate, date;

		switch(oDateGroupItem.DateTimeGrouping)
		{
			case Asc.EDateTimeGroup.datetimegroupDay://day
			{
				date = new Asc.cDate(Date.UTC( oDateGroupItem.Year, oDateGroupItem.Month - 1, oDateGroupItem.Day));
				startDate = date.getExcelDateWithTime();
				date.addDays(1);
				endDate = date.getExcelDateWithTime();
				break;
			}
			case Asc.EDateTimeGroup.datetimegroupHour://hour
			{
				startDate = new Asc.cDate(Date.UTC( oDateGroupItem.Year, oDateGroupItem.Month - 1, oDateGroupItem.Day, oDateGroupItem.Hour, 1)).getExcelDateWithTime();
				endDate = new Asc.cDate(Date.UTC( oDateGroupItem.Year, oDateGroupItem.Month - 1, oDateGroupItem.Day, oDateGroupItem.Hour, 59)).getExcelDateWithTime();
				break;
			}
			case Asc.EDateTimeGroup.datetimegroupMinute://minute
			{
				startDate = new Asc.cDate(Date.UTC( oDateGroupItem.Year, oDateGroupItem.Month - 1, oDateGroupItem.Day, oDateGroupItem.Hour, oDateGroupItem.Minute, 0)).getExcelDateWithTime();
				endDate = new Asc.cDate(Date.UTC( oDateGroupItem.Year, oDateGroupItem.Month - 1, oDateGroupItem.Day, oDateGroupItem.Hour, oDateGroupItem.Minute, 59)).getExcelDateWithTime();
				break;
			}
			case Asc.EDateTimeGroup.datetimegroupMonth://month
			{
				date = new Asc.cDate(Date.UTC( oDateGroupItem.Year, oDateGroupItem.Month - 1, 1));
				startDate = date.getExcelDateWithTime();
				date.addMonths(1);
				endDate = date.getExcelDateWithTime();
				break;
			}
			case Asc.EDateTimeGroup.datetimegroupSecond://second
			{
				startDate = new Asc.cDate(Date.UTC( oDateGroupItem.Year, oDateGroupItem.Month - 1, oDateGroupItem.Day, oDateGroupItem.Hour, oDateGroupItem.Minute, oDateGroupItem.Second, 0)).getExcelDateWithTime();
				endDate = new Asc.cDate(Date.UTC( oDateGroupItem.Year, oDateGroupItem.Month - 1, oDateGroupItem.Day, oDateGroupItem.Hour, oDateGroupItem.Minute, oDateGroupItem.Second,  1000)).getExcelDateWithTime();
				break;
			}
			case Asc.EDateTimeGroup.datetimegroupYear://year
			{
				date = new Asc.cDate(Date.UTC( oDateGroupItem.Year, 0));
				startDate = date.getExcelDateWithTime();
				date.addYears(1);
				endDate = date.getExcelDateWithTime();
				break;
			}
		}

		this.start = startDate;
		this.end = endDate;
		this.dateTimeGrouping = oDateGroupItem.DateTimeGrouping;
	};

	/** @constructor */
	function QueryTable() {
		this.queryTableRefresh = null;

		this.adjustColumnWidth = null;
		this.applyAlignmentFormats = null;
		this.applyBorderFormats = null;
		this.applyFontFormats = null;
		this.applyNumberFormats = null;
		this.applyPatternFormats = null;
		this.applyWidthHeightFormats = null;
		this.autoFormatId = null;
		this.backgroundRefresh = null;

		this.connectionId = null;
		this.disableEdit = null;
		this.disableRefresh = null;
		this.fillFormulas = null;
		this.firstBackgroundRefresh = null;
		this.growShrinkType = null;
		this.headers = null;
		this.intermediate = null;
		this.name = null;
		this.preserveFormatting = null;
		this.refreshOnLoad = null;
		this.removeDataOnSave = null;
		this.rowNumbers = null;
	}
	QueryTable.prototype.clone = function() {
		var res = new QueryTable();

		res.queryTableRefresh = this.queryTableRefresh ? this.queryTableRefresh : null;

		res.adjustColumnWidth = this.adjustColumnWidth;
		res.applyAlignmentFormats = this.applyAlignmentFormats;
		res.applyBorderFormats = this.applyBorderFormats;
		res.applyFontFormats = this.applyFontFormats;
		res.applyNumberFormats = this.applyNumberFormats;
		res.applyPatternFormats = this.applyPatternFormats;
		res.applyWidthHeightFormats = this.applyWidthHeightFormats;
		res.autoFormatId = this.autoFormatId;
		res.backgroundRefresh = this.backgroundRefresh;

		res.connectionId = this.connectionId;
		res.disableEdit = this.disableEdit;
		res.disableRefresh = this.disableRefresh;
		res.fillFormulas = this.fillFormulas;
		res.firstBackgroundRefresh = this.firstBackgroundRefresh;
		res.growShrinkType = this.growShrinkType;
		res.headers = this.headers;
		res.intermediate = this.intermediate;
		res.name = this.name;
		res.preserveFormatting = this.preserveFormatting;
		res.refreshOnLoad = this.refreshOnLoad;
		res.removeDataOnSave = this.removeDataOnSave;
		res.rowNumbers = this.rowNumbers;

		return res;
	};

	QueryTable.prototype.deleteTableColumns = function(deletedTableColumns) {
		if (!deletedTableColumns) {
			return;
		}

		if (this.queryTableRefresh) {
			this.queryTableRefresh.deleteTableColumns(deletedTableColumns);
		}
	};
	QueryTable.prototype.syncIndexes = function(tableColumns) {
		//при удалении приходится меняться tableColumnId, поскольку id  у колонок таблиц у нас формируются на сохранение
		//а соотсетствие именно по id tableColumn

		if (this.queryTableRefresh) {
			this.queryTableRefresh.syncIndexes(tableColumns);
		}
	};



	/** @constructor */
	function QueryTableRefresh() {
		this.queryTableDeletedFields = null;
		this.queryTableFields = null;
		this.sortState = null;

		this.fieldIdWrapped = null;
		this.headersInLastRefr = null;
		this.minimumVersion = null;
		this.nextId = null;
		this.preserveSortFilterLayout = null;
		this.unboundColumnsLeft = null;
		this.unboundColumnsRight = null;
	}
	QueryTableRefresh.prototype.clone = function() {
		var res = new QueryTableRefresh();

		res.queryTableDeletedFields = this.queryTableDeletedFields ? this.queryTableDeletedFields.clone() : null;
		res.queryTableFields = this.queryTableFields ? this.queryTableFields.clone() : null;
		res.sortState = this.sortState ? this.sortState.clone() : null;

		res.fieldIdWrapped = this.fieldIdWrapped;
		res.headersInLastRefr = this.headersInLastRefr;
		res.minimumVersion = this.minimumVersion;
		res.nextId = this.nextId;
		res.preserveSortFilterLayout = this.preserveSortFilterLayout;
		res.unboundColumnsLeft = this.unboundColumnsLeft;
		res.unboundColumnsRight = this.unboundColumnsRight;

		return res;
	};
	QueryTableRefresh.prototype.deleteTableColumns = function(deletedTableColumns) {
		if (!deletedTableColumns) {
			return;
		}

		if (this.queryTableFields) {
			for (var i = 0; i < deletedTableColumns.length; i++) {
				for (var j = 0; j < this.queryTableFields.length; j++) {
					if (deletedTableColumns[i].queryTableFieldId === this.queryTableFields[j].tableColumnId) {
						var deletedField = this.queryTableFields.splice(j, 1);
						this.addDeletedField(deletedField[0]);
					}
				}
			}
		}
	};
	QueryTableRefresh.prototype.addDeletedField = function(deletedField) {
		if (!this.queryTableDeletedFields) {
			this.queryTableDeletedFields = [];
		}
		var newDeletedField = new QueryTableDeletedField();
		newDeletedField.name = deletedField.name;
		this.queryTableDeletedFields.push(newDeletedField);
	};
	QueryTableRefresh.prototype.syncIndexes = function(tableColumns) {
		//при удалении приходится меняться tableColumnId, поскольку id  у колонок таблиц у нас формируются на сохранение
		//а соотсетствие именно по id tableColumn

		if (this.queryTableFields) {
			for (var i = 0; i < tableColumns.length; i++) {
				for (var j = 0; j < this.queryTableFields.length; j++) {
					if (tableColumns[i].queryTableFieldId === this.queryTableFields[j].tableColumnId) {
						this.queryTableFields[j].tableColumnId = i + 1;
					}
				}
			}
		}
	};
	QueryTableRefresh.prototype.syncIndexes2 = function(tableColumns) {
		//при удалении приходится меняться tableColumnId, поскольку id  у колонок таблиц у нас формируются на сохранение
		//а соотсетствие именно по id tableColumn

		if (this.queryTableFields) {
			var i;
			var changeQueryTableFieldId = [];
			for (var j = 0; j < this.queryTableFields.length; j++) {
				for (i = 0; i < tableColumns.length; i++) {
					if (tableColumns[i].queryTableFieldId === this.queryTableFields[j].tableColumnId) {
						this.queryTableFields[j].tableColumnId = i + 1;
						this.queryTableFields[j].id = i + 1;
						changeQueryTableFieldId[i] =  i + 1;
						break;
					}
				}
			}
			for (i = 0; i < changeQueryTableFieldId.length; i++) {
				tableColumns[i].queryTableFieldId = changeQueryTableFieldId[i];
				if (null !== tableColumns[i].uniqueName) {
					tableColumns[i].uniqueName = changeQueryTableFieldId[i];
				}
			}
		}

	};


	/** @constructor */
	function QueryTableField() {
		this.name = null;
		this.id = null;
		this.tableColumnId = null;

		this.rowNumbers = null;
		this.fillFormulas = null;
		this.dataBound = null;
		this.clipped = null;
	}
	QueryTableField.prototype.clone = function() {
		var res = new QueryTableField();

		res.name = this.name;
		res.id = this.id;
		res.tableColumnId = this.tableColumnId;
		res.rowNumbers = this.rowNumbers;
		res.fillFormulas = this.fillFormulas;
		res.dataBound = this.dataBound;
		res.clipped = this.clipped;

		return res;
	};

	/** @constructor */
	function QueryTableDeletedField() {
		this.name = null;
	}
	QueryTableField.prototype.clone = function() {
		var res = new QueryTableDeletedField();
		res.name = this.name;
		return res;
	};

	/**
	 * @constructor
	 * @memberOf AscCommonExcel
	 */
	function CSharedStrings () {
		this.all = [];
		this.text = new Map();
		this.multiTextMap = new Map();
	}

	CSharedStrings.prototype.addText = function(text) {
		var index = this.text.get(text);
		if (undefined === index) {
			this.all.push(text);
			index = this.all.length;
			this.text.set(text, index);
			if (AscFonts.IsCheckSymbols) {
				AscFonts.FontPickerByCharacter.getFontsByString(text);
			}
		}
		return index;
	};
	CSharedStrings.prototype.addMultiText = function(multiText) {
		var index, i;
		var text = multiText.reduce(function(accumulator, currentValue) {
			return accumulator + currentValue.text;
		}, '');
		var mapElem = this.multiTextMap.get(text);
		if (!mapElem) {
			mapElem = [];
			this.multiTextMap.set(text, mapElem);
		}
		for (i = 0; i < mapElem.length; ++i) {
			if (AscCommonExcel.isEqualMultiText(multiText, this.all[mapElem[i] - 1])) {
				index = mapElem[i];
				break;
			}
		}
		if (undefined === index) {
			this.all.push(multiText);
			index = this.all.length;
			mapElem.push(index);
			if (AscFonts.IsCheckSymbols) {
				for (i = 0; i < multiText.length; ++i) {
					AscFonts.FontPickerByCharacter.getFontsByString(multiText[i].text);
				}
			}
		}
		return index;
	};
	CSharedStrings.prototype.get = function(index) {
		return 1 <= index && index <= this.all.length ? this.all[index - 1] : null;
	};
	CSharedStrings.prototype.getCount = function() {
		return this.all.length;
	};
	CSharedStrings.prototype.generateFontMap = function(oFontMap) {
		this.multiTextMap.forEach(function(mapElem) {
			for (var i = 0; i < mapElem.length; ++i) {
				var multiText = this.all[mapElem[i] - 1];
				for (var j = 0; j < multiText.length; ++j) {
					var part = multiText[j];
					if (part && part.format) {
						oFontMap[part.format.getName()] = 1;
					}
				}
			}
		}, this);
	};

	/**
	 * @constructor
	 * @memberOf AscCommonExcel
	 */
	function CWorkbookFormulas () {
		this.all = [];
	}

	CWorkbookFormulas.prototype.add = function(formula) {
		var index = formula.getIndexNumber();
		if (undefined === index) {
			this.all.push(formula);
			index = this.all.length;
			formula.setIndexNumber(index);
		}
		return formula;
	};
	CWorkbookFormulas.prototype.get = function(index) {
		return 1 <= index && index <= this.all.length ? this.all[index - 1] : null;
	};
	CWorkbookFormulas.prototype.getCount = function() {
		return this.all.length;
	};


	/** @constructor */
	function asc_CPageMargins (ws) {
		this.left = null;
		this.right = null;
		this.top = null;
		this.bottom = null;

		//TODO в историю нужно будет записывать эти параметры
		this.header = null;
		this.footer = null;

		this.ws = ws;

		return this;
	}
	asc_CPageMargins.prototype.init = function () {
		if (null == this.left)
			this.left = c_oAscPrintDefaultSettings.PageLeftField;
		if (null == this.top)
			this.top = c_oAscPrintDefaultSettings.PageTopField;
		if (null == this.right)
			this.right = c_oAscPrintDefaultSettings.PageRightField;
		if (null == this.bottom)
			this.bottom = c_oAscPrintDefaultSettings.PageBottomField;
		if (null == this.header)
			this.header = c_oAscPrintDefaultSettings.PageHeaderField;
		if (null == this.footer)
			this.footer = c_oAscPrintDefaultSettings.PageFooterField;
	};
	asc_CPageMargins.prototype.clone = function (ws) {
		var res = new asc_CPageMargins(ws);

		res.left = this.left;
		res.right = this.right;
		res.top = this.top;
		res.bottom = this.bottom;

		res.header = this.header;
		res.footer = this.footer;

		return res;
	};
	asc_CPageMargins.prototype.asc_getLeft = function () { return this.left; };
	asc_CPageMargins.prototype.asc_getRight = function () { return this.right; };
	asc_CPageMargins.prototype.asc_getTop = function () { return this.top; };
	asc_CPageMargins.prototype.asc_getBottom = function () { return this.bottom; };
	asc_CPageMargins.prototype.asc_getHeader = function () { return this.header; };
	asc_CPageMargins.prototype.asc_getFooter = function () { return this.footer; };

	asc_CPageMargins.prototype.asc_setLeft = function (newVal) {
		var oldVal = this.left;
		this.left = newVal;
		if (this.ws && AscCommon.History.Is_On() && oldVal !== this.left) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_Left, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}
		let wb = this.ws && this.ws.workbook && this.ws.workbook;
		wb && wb.handlers && wb.handlers.trigger("onChangePageSetupProps", this.ws.getId());
	};
	asc_CPageMargins.prototype.asc_setRight = function (newVal) {
		var oldVal = this.right;
		this.right = newVal;
		if (this.ws && AscCommon.History.Is_On() && oldVal !== this.right) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_Right, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}
		let wb = this.ws && this.ws.workbook && this.ws.workbook;
		wb && wb.handlers && wb.handlers.trigger("onChangePageSetupProps", this.ws.getId());
	};
	asc_CPageMargins.prototype.asc_setTop = function (newVal) {
		var oldVal = this.top;
		this.top = newVal;
		if (this.ws && AscCommon.History.Is_On() && oldVal !== this.top) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_Top, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}
		let wb = this.ws && this.ws.workbook && this.ws.workbook;
		wb && wb.handlers && wb.handlers.trigger("onChangePageSetupProps", this.ws.getId());
	};
	asc_CPageMargins.prototype.asc_setBottom = function (newVal) {
		var oldVal = this.bottom;
		this.bottom = newVal;
		if (this.ws && AscCommon.History.Is_On() && oldVal !== this.bottom) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_Bottom, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}
		let wb = this.ws && this.ws.workbook && this.ws.workbook;
		wb && wb.handlers && wb.handlers.trigger("onChangePageSetupProps", this.ws.getId());
	};

	asc_CPageMargins.prototype.asc_setHeader = function (newVal) {
		var oldVal = this.header;
		this.header = newVal;
		/*if (this.ws && AscCommon.History.Is_On() && oldVal !== this.top) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_Top, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}*/
	};
	asc_CPageMargins.prototype.asc_setFooter = function (newVal) {
		var oldVal = this.footer;
		this.footer = newVal;
		/*if (this.ws && AscCommon.History.Is_On() && oldVal !== this.bottom) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_Bottom, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}*/
	};
	asc_CPageMargins.prototype.asc_setOptions = function (obj) {
		var prop;
		prop = obj.asc_getLeft();
		if(prop !== this.asc_getLeft()) {
			this.asc_setLeft(prop);
		}
		prop = obj.asc_getRight();
		if(prop !== this.asc_getRight()) {
			this.asc_setRight(prop);
		}
		prop = obj.asc_getBottom();
		if(prop !== this.asc_getBottom()) {
			this.asc_setBottom(prop);
		}
		prop = obj.asc_getTop();
		if(prop !== this.asc_getTop()) {
			this.asc_setTop(prop);
		}
	};

	/** @constructor */
	function asc_CPageSetup(ws) {
		this.orientation = c_oAscPrintDefaultSettings.PageOrientation;
		this.width = c_oAscPrintDefaultSettings.PageWidth;
		this.height = c_oAscPrintDefaultSettings.PageHeight;

		this.fitToWidth = 1; //default -> 1, 0 -> automatic
		this.fitToHeight = 1; //default -> 1, 0 -> automatic

		// ToDo
		this.blackAndWhite = false;
		this.cellComments = 0; // none ST_CellComments
		this.copies = 1;
		this.draft = false;
		this.errors = 0; // displayed ST_PrintError
		this.firstPageNumber = null;//default 1
		this.pageOrder = 0; // downThenOver ST_PageOrder
		this.scale = 100;
		this.useFirstPageNumber = false;
		this.usePrinterDefaults = true;
		this.horizontalDpi = 600;
		this.verticalDpi = 600;
		this.paperUnits = 0;

		//для превью передаём из интерфейса
		this.headerFooter = null;
		this.printArea = null;

		this.ws = ws;

		return this;
	}
	asc_CPageSetup.prototype.clone = function (ws) {
		var res = new asc_CPageSetup(ws);

		res.orientation = this.orientation;
		res.width = this.width;
		res.height = this.height;

		res.fitToWidth = this.fitToWidth; //default -> 1, 0 -> automatic
		res.fitToHeight = this.fitToHeight; //default -> 1, 0 -> automatic

		res.blackAndWhite = this.blackAndWhite;
		res.cellComments = this.cellComments; // none ST_CellComments
		res.copies = this.copies;
		res.draft = this.draft;
		res.errors = this.errors; // displayed ST_PrintError
		res.firstPageNumber = this.firstPageNumber;
		res.pageOrder = this.pageOrder; // downThenOver ST_PageOrder
		res.scale = this.scale;
		res.useFirstPageNumber = this.useFirstPageNumber;
		res.usePrinterDefaults = this.usePrinterDefaults;
		res.horizontalDpi = this.horizontalDpi;
		res.verticalDpi = this.verticalDpi;
		res.paperUnits = this.paperUnits;

		return res;
	};
	asc_CPageSetup.prototype.asc_getOrientation = function () { return this.orientation; };
	asc_CPageSetup.prototype.asc_getWidth = function () { return this.width; };
	asc_CPageSetup.prototype.asc_getHeight = function () { return this.height; };

	asc_CPageSetup.prototype.asc_setOrientation = function (newVal) {
		var oldVal = this.orientation;
		this.orientation = newVal;
		if (this.ws && AscCommon.History.Is_On() && oldVal !== this.orientation) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_Orientation, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}
		let wb = this.ws && this.ws.workbook && this.ws.workbook;
		wb && wb.handlers && wb.handlers.trigger("onChangePageSetupProps", this.ws.getId());
	};
	asc_CPageSetup.prototype.asc_setWidth = function (newVal) {
		var oldVal = this.width;
		this.width = newVal;
		if (this.ws && AscCommon.History.Is_On() && oldVal !== this.width) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_Width, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}
		let wb = this.ws && this.ws.workbook && this.ws.workbook;
		wb && wb.handlers && wb.handlers.trigger("onChangePageSetupProps", this.ws.getId());
	};
	asc_CPageSetup.prototype.asc_setHeight = function (newVal) {
		var oldVal = this.height;
		this.height = newVal;
		if (this.ws && AscCommon.History.Is_On() && oldVal !== this.height) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_Height, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}
		let wb = this.ws && this.ws.workbook && this.ws.workbook;
		wb && wb.handlers && wb.handlers.trigger("onChangePageSetupProps", this.ws.getId());
	};
	asc_CPageSetup.prototype.asc_getFirstPageNumber = function () {
		return this.useFirstPageNumber ? this.firstPageNumber : null;
	};
	asc_CPageSetup.prototype.asc_setFirstPageNumber = function (newVal) {
		let oldVal = this.firstPageNumber;
		let oldValUse = this.useFirstPageNumber;
		if (newVal) {
			newVal = parseInt(newVal);
		}
		this.firstPageNumber = newVal;
		this.useFirstPageNumber = true;
		if (this.ws && AscCommon.History.Is_On() && (oldVal !== this.firstPageNumber || oldValUse !== this.useFirstPageNumber)) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_FirstPageNumber, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}
	};

	asc_CPageSetup.prototype.asc_getFitToWidth = function () {
		if(!this.ws) {
			return this.fitToWidth;
		}
		var fitToPage = this.ws && this.ws.sheetPr && this.ws.sheetPr.FitToPage;
		return fitToPage ? this.fitToWidth : 0;
	};
	asc_CPageSetup.prototype.asc_getFitToHeight = function () {
		if(!this.ws) {
			return this.fitToHeight;
		}
		var fitToPage = this.ws && this.ws.sheetPr && this.ws.sheetPr.FitToPage;
		return fitToPage ? this.fitToHeight : 0;
	};

	asc_CPageSetup.prototype.asc_getScale = function () { return this.scale; };

	asc_CPageSetup.prototype.asc_setFitToWidth = function (newVal) {
		//TODO заглушка! потому что из меню проставляется булево значение, а должно быть число
		if(newVal === true) {
			newVal = 1;
		} else if(newVal === false) {
			newVal = 0;
		}

		var oldVal = this.fitToWidth;
		this.fitToWidth = newVal;
		if (this.ws && AscCommon.History.Is_On() && oldVal !== this.fitToWidth) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_FitToWidth, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}
		let wb = this.ws && this.ws.workbook && this.ws.workbook;
		wb && wb.handlers && wb.handlers.trigger("onChangePageSetupProps", this.ws.getId());
	};
	asc_CPageSetup.prototype.asc_setFitToHeight = function (newVal) {
		//TODO заглушка! потому что из меню проставляется булево значение, а должно быть число
		if(newVal === true) {
			newVal = 1;
		} else if(newVal === false) {
			newVal = 0;
		}

		var oldVal = this.fitToHeight;
		this.fitToHeight = newVal;
		if (this.ws && AscCommon.History.Is_On() && oldVal !== this.fitToHeight) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_FitToHeight, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}
		let wb = this.ws && this.ws.workbook && this.ws.workbook;
		wb && wb.handlers && wb.handlers.trigger("onChangePageSetupProps", this.ws.getId());
	};
	asc_CPageSetup.prototype.asc_setOptions = function (obj) {

		var prop;
		prop = obj.asc_getOrientation();
		if(prop !== this.asc_getOrientation()) {
			this.asc_setOrientation(prop);
		}
		prop = obj.asc_getWidth();
		if(prop !== this.asc_getWidth()) {
			this.asc_setWidth(prop);
		}
		prop = obj.asc_getHeight();
		if(prop !== this.asc_getHeight()) {
			this.asc_setHeight(prop);
		}
		prop = obj.asc_getHeight();
		if(prop !== this.asc_getHeight()) {
			this.asc_setHeight(prop);
		}
		prop = obj.asc_getFitToWidth();
		if(prop !== this.asc_getFitToWidth()) {
			this.asc_setFitToWidth(prop);
		}
		prop = obj.asc_getFitToHeight();
		if(prop !== this.asc_getFitToHeight()) {
			this.asc_setFitToHeight(prop);
		}
		prop = obj.asc_getFirstPageNumber();
		if(prop !== this.asc_getFirstPageNumber()) {
			this.asc_setFirstPageNumber(prop);
		}
	};

	asc_CPageSetup.prototype.asc_setScale = function (newVal) {
		var oldVal = this.scale;
		this.scale = newVal;
		if (this.ws && AscCommon.History.Is_On() && oldVal !== this.scale) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_Scale, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}
		let wb = this.ws && this.ws.workbook && this.ws.workbook;
		wb && wb.handlers && wb.handlers.trigger("onChangePageSetupProps", this.ws.getId());
	};

	asc_CPageSetup.prototype.getPreviewHeaderFooter = function () {
		var res = null;
		if (this.headerFooter) {
			res = new CHeaderFooter();
			var tempEditor = new AscCommonExcel.CHeaderFooterEditor();
			tempEditor.setPropsFromInterface(this.headerFooter, true);
			tempEditor._saveToModel(this, res);
		}
		return res;
	};

	/** @constructor */
	//этот объект используется как в модели, так и в меню для передачи измененных опций page layout
	//если определена ws - это означает, что этот объект лежит в модели и при изменении его свойств идёт запись в историю
	//в противном случае запись в историю не происходит
	function asc_CPageOptions(ws) {
		this.pageMargins = new asc_CPageMargins(ws);
		this.pageSetup = new asc_CPageSetup(ws);
		this.gridLines = null;
		this.headings = null;
		this.gridLinesSet = null;
		this.horizontalCentered = null;
		this.verticalCentered = null;

		this.ws = ws;

		//только для передачи из интефейса
		this.printTitlesWidth = null;
		this.printTitlesHeight = null;

		return this;
	}
	asc_CPageOptions.prototype.init = function () {
		this.pageMargins.init();

		if (null == this.gridLines)
			this.gridLines = c_oAscPrintDefaultSettings.PageGridLines;
		if (null == this.headings)
			this.headings = c_oAscPrintDefaultSettings.PageHeadings;
	};
	asc_CPageOptions.prototype.getJson = function (ws) {
		let res = {};

		res["gridLines"] = this.gridLines;
		res["headings"] = this.headings;
		res["gridLinesSet"] = this.gridLinesSet;
		res["horizontalCentered"] = this.horizontalCentered;
		res["verticalCentered"] = this.verticalCentered;
		res["pageMargins"] = {};
		res["pageMargins"]["bottom"] = this.pageMargins.bottom;
		res["pageMargins"]["footer"] = this.pageMargins.footer;
		res["pageMargins"]["header"] = this.pageMargins.header;
		res["pageMargins"]["left"] = this.pageMargins.left;
		res["pageMargins"]["right"] = this.pageMargins.right;
		res["pageMargins"]["top"] = this.pageMargins.top;


		res["pageSetup"] = {};
		res["pageSetup"]["blackAndWhite"] = this.pageSetup.blackAndWhite;
		res["pageSetup"]["cellComments"] = this.pageSetup.cellComments;
		res["pageSetup"]["copies"] = this.pageSetup.copies;
		res["pageSetup"]["draft"] = this.pageSetup.draft;
		res["pageSetup"]["errors"] = this.pageSetup.errors;
		res["pageSetup"]["firstPageNumber"] = this.pageSetup.firstPageNumber;
		res["pageSetup"]["fitToHeight"] = this.pageSetup.fitToHeight;
		res["pageSetup"]["fitToWidth"] = this.pageSetup.fitToWidth;
		res["pageSetup"]["headerFooter"] = this.pageSetup.headerFooter;
		res["pageSetup"]["height"] = this.pageSetup.height;
		res["pageSetup"]["horizontalDpi"] = this.pageSetup.horizontalDpi;
		res["pageSetup"]["orientation"] = this.pageSetup.orientation;
		res["pageSetup"]["pageOrder"] = this.pageSetup.pageOrder;
		res["pageSetup"]["paperUnits"] = this.pageSetup.paperUnits;
		res["pageSetup"]["printArea"] = this.pageSetup.printArea;
		res["pageSetup"]["scale"] = this.pageSetup.scale;
		res["pageSetup"]["useFirstPageNumber"] = this.pageSetup.useFirstPageNumber;
		res["pageSetup"]["usePrinterDefaults"] = this.pageSetup.usePrinterDefaults;
		res["pageSetup"]["verticalDpi"] = this.pageSetup.verticalDpi;
		res["pageSetup"]["width"] = this.pageSetup.width;

		if (ws.headerFooter) {
			res["pageSetup"]["headerFooter"] = {};
			res["pageSetup"]["headerFooter"]["alignWithMargins"] = ws.headerFooter.alignWithMargins;
			res["pageSetup"]["headerFooter"]["differentFirst"] = ws.headerFooter.differentFirst;
			res["pageSetup"]["headerFooter"]["differentOddEven"] = ws.headerFooter.differentOddEven;
			res["pageSetup"]["headerFooter"]["evenFooter"] = ws.headerFooter.evenFooter && ws.headerFooter.evenFooter.getStr();
			res["pageSetup"]["headerFooter"]["evenHeader"] = ws.headerFooter.evenHeader && ws.headerFooter.evenHeader.getStr();
			res["pageSetup"]["headerFooter"]["firstFooter"] = ws.headerFooter.firstFooter && ws.headerFooter.firstFooter.getStr();
			res["pageSetup"]["headerFooter"]["firstHeader"] = ws.headerFooter.firstHeader && ws.headerFooter.firstHeader.getStr();
			res["pageSetup"]["headerFooter"]["oddFooter"] = ws.headerFooter.oddFooter && ws.headerFooter.oddFooter.getStr();
			res["pageSetup"]["headerFooter"]["oddHeader"] = ws.headerFooter.oddHeader && ws.headerFooter.oddHeader.getStr();
			res["pageSetup"]["headerFooter"]["scaleWithDoc"] = ws.headerFooter.scaleWithDoc;
		}

		res["printTitlesHeight"] = this.printTitlesHeight;
		res["printTitlesWidth"] = this.printTitlesWidth;

		return res;
	};
	asc_CPageOptions.prototype.setJson = function (props) {
		let checkOnNull = function (to, from) {
			return to != null ? to : from;
		};

		this.gridLines = checkOnNull(props["gridLines"], this.gridLines);
		this.headings = checkOnNull(props["headings"], this.headings);
		this.gridLinesSet = checkOnNull(props["gridLinesSet"], this.gridLinesSet);
		this.horizontalCentered = checkOnNull(props["horizontalCentered"], this.horizontalCentered);
		this.verticalCentered = checkOnNull(props["verticalCentered"], this.verticalCentered);

		this.pageMargins.bottom = checkOnNull(props["pageMargins"]["bottom"], this.pageMargins.bottom);
		this.pageMargins.footer = checkOnNull(props["pageMargins"]["footer"], this.pageMargins.footer);
		this.pageMargins.header = checkOnNull(props["pageMargins"]["header"], this.pageMargins.header);
		this.pageMargins.left = checkOnNull(props["pageMargins"]["left"], this.pageMargins.left);
		this.pageMargins.right = checkOnNull(props["pageMargins"]["right"], this.pageMargins.right);
		this.pageMargins.top = checkOnNull(props["pageMargins"]["top"], this.pageMargins.top);
		this.pageMargins.bottom = checkOnNull(props["pageMargins"]["bottom"], this.pageMargins.bottom);

		this.pageSetup.blackAndWhite = checkOnNull(props["pageSetup"]["blackAndWhite"], this.pageSetup.blackAndWhite);
		this.pageSetup.cellComments = checkOnNull(props["pageSetup"]["cellComments"], this.pageSetup.cellComments);
		this.pageSetup.copies = checkOnNull(props["pageSetup"]["copies"],this.pageSetup.copies);
		this.pageSetup.draft = checkOnNull(props["pageSetup"]["draft"], this.pageSetup.draft);
		this.pageSetup.errors = checkOnNull(props["pageSetup"]["errors"],this.pageSetup.errors);
		this.pageSetup.firstPageNumber = checkOnNull(props["pageSetup"]["firstPageNumber"]< this.pageSetup.firstPageNumber);
		this.pageSetup.fitToHeight = checkOnNull(props["pageSetup"]["fitToHeight"], this.pageSetup.fitToHeight);
		this.pageSetup.fitToWidth = checkOnNull(props["pageSetup"]["fitToWidth"], this.pageSetup.fitToWidth);
		this.pageSetup.height = checkOnNull(props["pageSetup"]["height"], this.pageSetup.height);
		this.pageSetup.horizontalDpi = checkOnNull(props["pageSetup"]["horizontalDpi"], this.pageSetup.horizontalDpi);
		this.pageSetup.orientation = checkOnNull(props["pageSetup"]["orientation"], this.pageSetup.orientation);
		this.pageSetup.pageOrder = checkOnNull(props["pageSetup"]["pageOrder"], this.pageSetup.pageOrder);
		this.pageSetup.paperUnits = checkOnNull(props["pageSetup"]["paperUnits"], this.pageSetup.paperUnits);
		this.pageSetup.printArea = checkOnNull(props["pageSetup"]["printArea"], this.pageSetup.printArea);
		this.pageSetup.scale = checkOnNull(props["pageSetup"]["scale"], this.pageSetup.scale);
		this.pageSetup.useFirstPageNumber = checkOnNull(props["pageSetup"]["useFirstPageNumber"], this.pageSetup.useFirstPageNumber);
		this.pageSetup.usePrinterDefaults = checkOnNull(props["pageSetup"]["usePrinterDefaults"], this.pageSetup.usePrinterDefaults);
		this.pageSetup.verticalDpi = checkOnNull(props["pageSetup"]["verticalDpi"], this.pageSetup.verticalDpi);
		this.pageSetup.width = checkOnNull(props["pageSetup"]["width"], this.pageSetup.width);

		this.pageSetup.headerFooter = checkOnNull(props["pageSetup"]["headerFooter"], this.pageSetup.headerFooter);
		/*if (ws.headerFooter) {
			res["pageSetup"]["headerFooter"] = {
				"alignWithMargins": ws.headerFooter.alignWithMargins,
				"differentFirst": ws.headerFooter.differentFirst,
				"differentOddEven": ws.headerFooter.differentOddEven,
				"evenFooter": ws.headerFooter.evenFooter && ws.headerFooter.evenFooter.getStr(),
				"evenHeader": ws.headerFooter.evenHeader && ws.headerFooter.evenHeader.getStr(),
				"firstFooter": ws.headerFooter.firstFooter && ws.headerFooter.firstFooter.getStr(),
				"firstHeader": ws.headerFooter.firstHeader && ws.headerFooter.firstHeader.getStr(),
				"oddFooter": ws.headerFooter.oddFooter && ws.headerFooter.oddFooter.getStr(),
				"oddHeader": ws.headerFooter.oddHeader && ws.headerFooter.oddHeader.getStr(),
				"scaleWithDoc": ws.headerFooter.scaleWithDoc
			};
		}*/

		this.printTitlesHeight = checkOnNull(props["printTitlesHeight"], this.printTitlesHeight);
		this.printTitlesWidth = checkOnNull(props["printTitlesWidth"], this.printTitlesWidth);
	};
	asc_CPageOptions.prototype.initPrintTitles = function () {
		//функция добавлена только для того, чтобы в интерфейс передать текущие заголовки печати, которые хранятся как именованный диапазон

		this.printTitlesWidth = null;
		this.printTitlesHeight = null;

		let printTitles = this.ws.workbook.getDefinesNames("Print_Titles", this.ws.getId());
		let c1, c2, r1, r2;
		let t = this;
		if (printTitles) {
			let printTitleRefs;
			AscCommonExcel.executeInR1C1Mode(false, function () {
				printTitleRefs = AscCommonExcel.getRangeByRef(printTitles.ref, t.ws, true, true)
			});
			if (printTitleRefs && printTitleRefs.length) {
				for (let i = 0; i < printTitleRefs.length; i++) {
					let bbox = printTitleRefs[i].bbox;
					if (bbox) {
						if (Asc.c_oAscSelectionType.RangeCol === bbox.getType()) {
							c1 = bbox.c1;
							c2 = bbox.c2;
						} else if(Asc.c_oAscSelectionType.RangeRow === bbox.getType()) {
							r1 = bbox.r1;
							r2 = bbox.r2;
						}
					}
				}
			}
		}
		if (c1 !== undefined) {
			this.printTitlesWidth = new Asc.Range(c1, 0, c2, gc_nMaxRow0).getAbsName();
		}
		if(r1 !== undefined) {
			this.printTitlesHeight = new Asc.Range(0, r1, gc_nMaxCol0, r2).getAbsName()
		}
	};
	asc_CPageOptions.prototype.clone = function (ws) {
		let res = new asc_CPageOptions(ws);

		res.pageMargins = this.pageMargins.clone(ws);
		res.pageSetup = this.pageSetup.clone(ws);
		res.gridLines = this.gridLines;
		res.headings = this.headings;
		res.gridLinesSet = this.gridLinesSet;
		res.horizontalCentered = this.horizontalCentered;
		res.verticalCentered = this.verticalCentered;

		return res;
	};
	asc_CPageOptions.prototype.asc_getPageMargins = function () { return this.pageMargins; };
	asc_CPageOptions.prototype.asc_getPageSetup = function () { return this.pageSetup; };
	asc_CPageOptions.prototype.asc_getGridLines = function () { return this.gridLines; };
	asc_CPageOptions.prototype.asc_getHeadings = function () { return this.headings; };
	asc_CPageOptions.prototype.asc_getGridLinesSet = function () { return this.gridLinesSet; };
	asc_CPageOptions.prototype.asc_getHorizontalCentered = function () { return this.horizontalCentered; };
	asc_CPageOptions.prototype.asc_getVerticalCentered = function () { return this.verticalCentered; };
	//методы только для меню, без добавляем в историю
	asc_CPageOptions.prototype.asc_setPageMargins = function (val) { this.pageMargins = val; };
	asc_CPageOptions.prototype.asc_setPageSetup = function (val) { this.pageSetup = val; };

	asc_CPageOptions.prototype.asc_setGridLines = function (newVal) {
		let oldVal = this.gridLines;
		this.gridLines = newVal;
		if (this.ws && AscCommon.History.Is_On() && oldVal !== this.gridLines) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_GridLines, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}
	};
	asc_CPageOptions.prototype.asc_setHeadings = function (newVal) {
		let oldVal = this.headings;
		this.headings = newVal;
		if (this.ws && AscCommon.History.Is_On() && oldVal !== this.headings) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_Headings, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}
	};
	asc_CPageOptions.prototype.asc_setHorizontalCentered = function (newVal) {
		let oldVal = this.horizontalCentered;
		this.horizontalCentered = newVal;
		if (this.ws && AscCommon.History.Is_On() && oldVal !== this.horizontalCentered) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_HorizontalCentered, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}
	};
	asc_CPageOptions.prototype.asc_setVerticalCentered = function (newVal) {
		let oldVal = this.verticalCentered;
		this.verticalCentered = newVal;
		if (this.ws && AscCommon.History.Is_On() && oldVal !== this.verticalCentered) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_VerticalCentered, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}
	};
	asc_CPageOptions.prototype.asc_setGridLinesSet = function (newVal) {
		/*let oldVal = this.verticalCentered;
		this.verticalCentered = newVal;
		if (this.ws && AscCommon.History.Is_On() && oldVal !== this.verticalCentered) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoLayout, AscCH.historyitem_Layout_VerticalCentered, this.ws.getId(),
				null, new UndoRedoData_Layout(oldVal, newVal));
		}*/
	};
	asc_CPageOptions.prototype.asc_setOptions = function (obj) {
		let gridLines = obj.asc_getGridLines();
		if(gridLines !== this.asc_getGridLines()) {
			this.asc_setGridLines(gridLines);
		}
		let heading = obj.asc_getHeadings();
		if(heading !== this.asc_getHeadings()) {
			this.asc_setHeadings(heading);
		}
		let gridLinesSet = obj.asc_getGridLinesSet();
		if(gridLinesSet !== this.asc_getGridLinesSet()) {
			this.asc_setGridLinesSet(gridLinesSet);
		}
		let horizontalCentered = obj.asc_getHorizontalCentered();
		if(horizontalCentered !== this.asc_getHorizontalCentered()) {
			this.asc_setHorizontalCentered(horizontalCentered);
		}
		let verticalCentered = obj.asc_getVerticalCentered();
		if(verticalCentered !== this.asc_getVerticalCentered()) {
			this.asc_setVerticalCentered(verticalCentered);
		}

		this.asc_getPageMargins().asc_setOptions(obj.asc_getPageMargins());
		this.asc_getPageSetup().asc_setOptions(obj.asc_getPageSetup());
	};
	asc_CPageOptions.prototype.asc_setPrintTitlesWidth = function (newVal) {
		this.printTitlesWidth = newVal;
		let wb = this.ws && this.ws.workbook && this.ws.workbook;
		wb && wb.handlers && wb.handlers.trigger("onChangePageSetupProps", this.ws.getId());
	};
	asc_CPageOptions.prototype.asc_setPrintTitlesHeight = function (newVal) {
		this.printTitlesHeight = newVal;
		let wb = this.ws && this.ws.workbook && this.ws.workbook;
		wb && wb.handlers && wb.handlers.trigger("onChangePageSetupProps", this.ws.getId());
	};
	asc_CPageOptions.prototype.asc_getPrintTitlesWidth = function () {
		return this.printTitlesWidth;
	};
	asc_CPageOptions.prototype.asc_getPrintTitlesHeight = function () {
		return this.printTitlesHeight;
	};


	function CHeaderFooter(ws) {
		this.ws = ws;

		this.alignWithMargins = null;
		this.differentFirst = null;
		this.differentOddEven = null;
		this.scaleWithDoc = true;
		this.evenFooter = null;
		this.evenHeader = null;
		this.firstFooter = null;
		this.firstHeader = null;
		this.oddFooter = null;
		this.oddHeader = null;

		return this;
	}

	CHeaderFooter.prototype.clone = function (ws) {
		var oRes = new CHeaderFooter(ws);

		oRes.alignWithMargins = this.alignWithMargins;
		oRes.differentFirst = this.differentFirst;
		oRes.differentOddEven = this.differentOddEven;
		oRes.scaleWithDoc = this.scaleWithDoc;
		oRes.evenFooter = this.evenFooter ? this.evenFooter.clone() : null;
		oRes.evenHeader = this.evenHeader ? this.evenHeader.clone() : null;
		oRes.firstFooter = this.firstFooter ? this.firstFooter.clone() : null;
		oRes.firstHeader = this.firstHeader ? this.firstHeader.clone() : null;
		oRes.oddFooter = this.oddFooter ? this.oddFooter.clone() : null;
		oRes.oddHeader = this.oddHeader ? this.oddHeader.clone() : null;

		return oRes;
	};

	CHeaderFooter.prototype.getAlignWithMargins = function () { return true === this.alignWithMargins || null === this.alignWithMargins; };
	CHeaderFooter.prototype.getDifferentFirst = function () { return this.differentFirst; };
	CHeaderFooter.prototype.getDifferentOddEven = function () { return this.differentOddEven; };
	CHeaderFooter.prototype.getScaleWithDoc = function () { return this.scaleWithDoc; };
	CHeaderFooter.prototype.getEvenFooter = function () { return this.evenFooter; };
	CHeaderFooter.prototype.getEvenHeader = function () { return this.evenHeader; };
	CHeaderFooter.prototype.getFirstFooter = function () { return this.firstFooter; };
	CHeaderFooter.prototype.getFirstHeader = function () { return this.firstHeader; };
	CHeaderFooter.prototype.getOddFooter = function () { return this.oddFooter; };
	CHeaderFooter.prototype.getOddHeader = function () { return this.oddHeader; };

	CHeaderFooter.prototype.setAlignWithMargins = function (val) { this.alignWithMargins = val; };
	CHeaderFooter.prototype.setDifferentFirst = function (val) { this.differentFirst = val; };
	CHeaderFooter.prototype.setDifferentOddEven = function (val) { this.differentOddEven = val; };
	CHeaderFooter.prototype.setScaleWithDoc = function (val) { this.scaleWithDoc = val; };
	CHeaderFooter.prototype.setEvenFooter = function (newVal) {
		var oldVal = this.evenFooter ? this.evenFooter.str : null;

		if(!compareValues(oldVal, newVal)) {
			if(null === newVal) {
				this.evenFooter = null;
			} else {
				this.evenFooter = new Asc.CHeaderFooterData();
				this.evenFooter.setStr(newVal);
				//TODO separate on read/init in open
				this.evenFooter.setType(Asc.c_oAscPageHFType && Asc.c_oAscPageHFType.evenFooter);
			}

			if (this.ws && AscCommon.History.Is_On()) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoHeaderFooter, AscCH.historyitem_Footer_Even, this.ws.getId(),
					null, new UndoRedoData_Layout(oldVal, newVal));
			}
		}
	};
	CHeaderFooter.prototype.setEvenHeader = function (newVal) {
		var oldVal = this.evenHeader ? this.evenHeader.str : null;

		if(!compareValues(oldVal, newVal)) {
			if(null === newVal) {
				this.evenHeader = null;
			} else {
				this.evenHeader = new Asc.CHeaderFooterData();
				this.evenHeader.setStr(newVal);
				//TODO separate on read/init in open
				this.evenHeader.setType(Asc.c_oAscPageHFType && Asc.c_oAscPageHFType.evenHeader);
			}

			if (this.ws && AscCommon.History.Is_On()) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoHeaderFooter, AscCH.historyitem_Header_Even, this.ws.getId(),
					null, new UndoRedoData_Layout(oldVal, newVal));
			}
		}
	};
	CHeaderFooter.prototype.setFirstFooter = function (newVal) {
		var oldVal = this.firstFooter ? this.firstFooter.str : null;

		if(!compareValues(oldVal, newVal)) {
			if(null === newVal) {
				this.firstFooter = null;
			} else {
				this.firstFooter = new Asc.CHeaderFooterData();
				this.firstFooter.setStr(newVal);
				//TODO separate on read/init in open
				this.firstFooter.setType(Asc.c_oAscPageHFType && Asc.c_oAscPageHFType.firstFooter);
			}

			if (this.ws && AscCommon.History.Is_On()) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoHeaderFooter, AscCH.historyitem_Footer_First, this.ws.getId(),
					null, new UndoRedoData_Layout(oldVal, newVal));
			}
		}
	};
	CHeaderFooter.prototype.setFirstHeader = function (newVal) {
		var oldVal = this.firstHeader ? this.firstHeader.str : null;

		if(!compareValues(oldVal, newVal)) {
			if(null === newVal) {
				this.firstHeader = null;
			} else {
				this.firstHeader = new Asc.CHeaderFooterData();
				this.firstHeader.setStr(newVal);
				//TODO separate on read/init in open
				this.firstHeader.setType(Asc.c_oAscPageHFType && Asc.c_oAscPageHFType.firstHeader);
			}

			if (this.ws && AscCommon.History.Is_On()) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoHeaderFooter, AscCH.historyitem_Header_First, this.ws.getId(),
					null, new UndoRedoData_Layout(oldVal, newVal));
			}
		}
	};
	CHeaderFooter.prototype.setOddFooter = function (newVal) {
		var oldVal = this.oddFooter ? this.oddFooter.str : null;

		if(!compareValues(oldVal, newVal)) {
			if(null === newVal) {
				this.oddFooter = null;
			} else {
				this.oddFooter = new Asc.CHeaderFooterData();
				this.oddFooter.setStr(newVal);
				//TODO separate on read/init in open
				this.oddFooter.setType(Asc.c_oAscPageHFType && Asc.c_oAscPageHFType.oddFooter);
			}

			if (this.ws && AscCommon.History.Is_On()) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoHeaderFooter, AscCH.historyitem_Footer_Odd, this.ws.getId(),
					null, new UndoRedoData_Layout(oldVal, newVal));
			}
		}
	};
	CHeaderFooter.prototype.setOddHeader = function (newVal) {
		var oldVal = this.oddHeader ? this.oddHeader.str : null;

		if(!compareValues(oldVal, newVal)) {
			if(null === newVal) {
				this.oddHeader = null;
			} else {
				this.oddHeader = new Asc.CHeaderFooterData();
				this.oddHeader.setStr(newVal);
				//TODO separate on read/init in open
				this.oddHeader.setType(Asc.c_oAscPageHFType && Asc.c_oAscPageHFType.oddHeader);
			}

			if (this.ws && AscCommon.History.Is_On()) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoHeaderFooter, AscCH.historyitem_Header_Odd, this.ws.getId(),
					null, new UndoRedoData_Layout(oldVal, newVal));
			}
		}
	};

	CHeaderFooter.prototype.setAlignWithMargins = function (newVal) {
		var oldVal = this.alignWithMargins;
		var defaultVal = null === oldVal && (newVal === 1 || newVal === true);

		if(!compareValues(oldVal, newVal) && !defaultVal) {
			this.alignWithMargins = newVal;

			if (this.ws && AscCommon.History.Is_On()) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoHeaderFooter, AscCH.historyitem_Align_With_Margins, this.ws.getId(),
					null, new UndoRedoData_Layout(oldVal, newVal));
			}
		}
	};

	CHeaderFooter.prototype.setScaleWithDoc = function (newVal) {
		var oldVal = this.scaleWithDoc;

		if(!compareValues(oldVal, newVal)) {
			this.scaleWithDoc = newVal;

			if (this.ws && AscCommon.History.Is_On()) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoHeaderFooter, AscCH.historyitem_Scale_With_Doc, this.ws.getId(),
					null, new UndoRedoData_Layout(oldVal, newVal));
			}
		}
	};

	CHeaderFooter.prototype.setDifferentFirst = function (newVal) {
		var oldVal = this.differentFirst;
		var defaultVal = null === oldVal && (newVal === 0 || newVal === false);

		if(!compareValues(oldVal, newVal) && !defaultVal) {
			this.differentFirst = newVal;

			if (this.ws && AscCommon.History.Is_On()) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoHeaderFooter, AscCH.historyitem_Different_First, this.ws.getId(),
					null, new UndoRedoData_Layout(oldVal, newVal));
			}
		}
	};

	CHeaderFooter.prototype.setDifferentOddEven = function (newVal) {
		var oldVal = this.differentOddEven;
		var defaultVal = null === oldVal && (newVal === 0 || newVal === false);

		if(!compareValues(oldVal, newVal) && !defaultVal) {
			this.differentOddEven = newVal;

			if (this.ws && AscCommon.History.Is_On()) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoHeaderFooter, AscCH.historyitem_Different_Odd_Even, this.ws.getId(),
					null, new UndoRedoData_Layout(oldVal, newVal));
			}
		}
	};

	CHeaderFooter.prototype.setHeaderFooterData = function(str, type) {
		switch (type){
			case Asc.c_oAscPageHFType.firstHeader: {
				this.setFirstHeader(str);
				break;
			}
			case Asc.c_oAscPageHFType.oddHeader: {
				this.setOddHeader(str);
				break;
			}
			case Asc.c_oAscPageHFType.evenHeader: {
				this.setEvenHeader(str);
				break;
			}
			case Asc.c_oAscPageHFType.firstFooter: {
				this.setFirstFooter(str);
				break;
			}
			case Asc.c_oAscPageHFType.oddFooter: {
				this.setOddFooter(str);
				break;
			}
			case Asc.c_oAscPageHFType.evenFooter: {
				this.setEvenFooter(str);
				break;
			}
		}
	};

	CHeaderFooter.prototype.clean = function() {
		this.setFirstHeader(null);
		this.setOddHeader(null);
		this.setEvenHeader(null);
		this.setFirstFooter(null);
		this.setOddFooter(null);
		this.setEvenFooter(null);
	};

	CHeaderFooter.prototype.init = function () {
		if(this.evenFooter) {
			this.evenFooter.parse();
		}
		if(this.evenHeader) {
			this.evenHeader.parse();
		}
		if(this.firstFooter) {
			this.firstFooter.parse();
		}
		if(this.firstHeader) {
			this.firstHeader.parse();
		}
		if(this.oddFooter) {
			this.oddFooter.parse();
		}
		if(this.oddHeader) {
			this.oddHeader.parse();
		}
	};
	CHeaderFooter.prototype.getAllFonts = function (oFontMap) {
		if(this.evenFooter) {
			this.evenFooter.getAllFonts(oFontMap);
		}
		if(this.evenHeader) {
			this.evenHeader.getAllFonts(oFontMap);
		}
		if(this.firstFooter) {
			this.firstFooter.getAllFonts(oFontMap);
		}
		if(this.firstHeader) {
			this.firstHeader.getAllFonts(oFontMap);
		}
		if(this.oddFooter) {
			this.oddFooter.getAllFonts(oFontMap);
		}
		if(this.oddHeader) {
			this.oddHeader.getAllFonts(oFontMap);
		}
	};
	CHeaderFooter.prototype.getForInterface = function () {
		var res = null;
		var tempEditor = new AscCommonExcel.CHeaderFooterEditor();
		tempEditor._createAndDrawSections(Asc.c_oAscHeaderFooterType.first, null, this);
		tempEditor._createAndDrawSections(Asc.c_oAscHeaderFooterType.odd, null, this);
		tempEditor._createAndDrawSections(Asc.c_oAscHeaderFooterType.even, null, this);
		tempEditor.alignWithMargins = this.alignWithMargins;
		tempEditor.differentFirst = this.differentFirst;
		tempEditor.differentOddEven = this.differentOddEven;
		tempEditor.scaleWithDoc = this.scaleWithDoc;

		res = tempEditor.getPropsToInterface();

		return res;
	};
	CHeaderFooter.prototype.initFromJson = function (val) {
		this.alignWithMargins = val["alignWithMargins"];
		this.differentFirst = val["differentFirst"];
		this.differentOddEven = val["differentOddEven"];

		this.scaleWithDoc = val["scaleWithDoc"];

		if (val["evenFooter"]) {
			this.evenFooter = new CHeaderFooterData();
			this.evenFooter.setType(Asc.c_oAscPageHFType.evenFooter);
			this.evenFooter.setStr(val["evenFooter"]);
		}
		if (val["evenHeader"]) {
			this.evenHeader = new CHeaderFooterData();
			this.evenHeader.setType(Asc.c_oAscPageHFType.evenHeader);
			this.evenHeader.setStr(val["evenHeader"]);
		}
		if (val["firstFooter"]) {
			this.firstFooter = new CHeaderFooterData();
			this.firstFooter.setType(Asc.c_oAscPageHFType.firstFooter);
			this.firstFooter.setStr(val["firstFooter"]);
		}
		if (val["firstHeader"]) {
			this.firstHeader = new CHeaderFooterData();
			this.firstHeader.setType(Asc.c_oAscPageHFType.firstHeader);
			this.firstHeader.setStr(val["firstHeader"]);
		}
		if (val["oddFooter"]) {
			this.oddFooter = new CHeaderFooterData();
			this.oddFooter.setType(Asc.c_oAscPageHFType.oddFooter);
			this.oddFooter.setStr(val["oddFooter"]);
		}
		if (val["oddHeader"]) {
			this.oddHeader = new CHeaderFooterData();
			this.oddHeader.setType(Asc.c_oAscPageHFType.oddHeader);
			this.oddHeader.setStr(val["oddHeader"]);
		}

		this.init();
	};

	CHeaderFooter.prototype.isEmpty = function () {
		if (this.alignWithMargins === null &&
			this.differentFirst === null &&
			this.differentOddEven === null &&
			this.scaleWithDoc === null &&
			this.evenFooter === null &&
			this.evenHeader === null &&
			this.firstFooter === null &&
			this.firstHeader === null &&
			this.oddFooter === null &&
			this.oddHeader === null) {
			return true;
		}
		return false;
	};


	function CHeaderFooterData(str) {
		this.str = str;
		this.parser = null;
		this.type = null;

		return this;
	}

	CHeaderFooterData.prototype.clone = function () {
		var oRes = new CHeaderFooterData();
		oRes.str = this.str;
		oRes.type = this.type;
		return oRes;
	};
	CHeaderFooterData.prototype.getStr = function () {
		return this.str;
	};
	CHeaderFooterData.prototype.setStr = function (val) {
		this.str = val;
	};
	CHeaderFooterData.prototype.setType = function (val) {
		this.type = val;
	};
	CHeaderFooterData.prototype.parse = function () {
		var parser = new window["AscCommonExcel"].HeaderFooterParser();
		parser.parse(this.str);
		this.parser = parser;
		return parser;
	};
	CHeaderFooterData.prototype.getAllFonts = function (oFontMap) {
		if (this.parser) {
			this.parser.getAllFonts(oFontMap);
		}
	};

	function CSortProperties(ws) {
		this.selection = null;
		this._newSelection = null;
		this.hasHeaders = null;
		this.columnSort = null;
		this.caseSensitive = null;

		this.levels = null;

		this.sortList = null;//массив, порядковый номер - его индекс в levels

		this.lockChangeHeaders = null;
		this.lockChangeOrientation = null;

		this._ws = ws;

		return this;
	}

	CSortProperties.prototype.asc_getHasHeaders = function () {
		return this.hasHeaders;
	};
	CSortProperties.prototype.asc_getColumnSort = function () {
		return this.columnSort;
	};
	CSortProperties.prototype.asc_getCaseSensitive = function () {
		return this.caseSensitive;
	};
	CSortProperties.prototype.asc_getLevels = function () {
		return this.levels;
	};
	CSortProperties.prototype.asc_getSortList = function () {
		return this.sortList;
	};
	CSortProperties.prototype.asc_getLockChangeHeaders = function () {
		return this.lockChangeHeaders;
	};
	CSortProperties.prototype.asc_getLockChangeOrientation = function () {
		return this.lockChangeOrientation;
	};
	CSortProperties.prototype.asc_setHasHeaders = function (val) {
		var oldVal = !!this.hasHeaders;
		if (this._newSelection && oldVal !== val) {
			if (val) {
				this._newSelection.r1++;
			} else {
				this._newSelection.r1--;
			}
			this._ws.setSelection(this._newSelection);
		}
		this.hasHeaders = val;
	};
	CSortProperties.prototype.asc_setColumnSort = function (val) {
		this.columnSort = val;
	};
	CSortProperties.prototype.asc_setCaseSensitive = function (val) {
		this.caseSensitive = val;
	};
	CSortProperties.prototype.asc_setLevels = function (val) {
		this.levels = val;
	};

	CSortProperties.prototype.asc_updateSortList = function (saveIndexes) {
		//TODO change selection
		this.generateSortList(saveIndexes);
	};
	//TODO delete
	CSortProperties.prototype.asc_getFilterInside = function () {
	};
	CSortProperties.prototype.generateSortList = function (saveIndexes) {
		var maxCount = 500;
		var selection = this._newSelection;
		var j;

		if (saveIndexes && this.sortList && this.sortList.length) {
			var newSortList = [];
			for (j in this.sortList) {
				newSortList[j] = this.getNameColumnByIndex(parseInt(j), selection);
			}
			this.sortList = newSortList;
		} else {
			this.sortList = [];
			if (this.columnSort) {
				for (j = selection.c1; j <= selection.c2; j++) {
					if (j - selection.c1 >= maxCount) {
						break;
					}
					this.sortList.push(this.getNameColumnByIndex(j - selection.c1, selection));
				}
			} else {
				for (j = selection.r1; j <= selection.r2; j++) {
					if (j - selection.r1 >= maxCount) {
						break;
					}
					this.sortList.push(this.getNameColumnByIndex(j - selection.r1, selection));
				}
			}
		}

		if (this.levels) {
			for (var i = 0; i < this.levels.length; i++) {
				if (!this.sortList[this.levels[i].index]) {
					this.sortList[this.levels[i].index] = this.getNameColumnByIndex(this.levels[i].index, selection);
				}
			}
		}
	};
	CSortProperties.prototype.asc_addBySortList = function (sRange) {
		//при количестве строк/столбцов более 500, добавляем по одной строке/одному столбцу
		var selection = this._newSelection;
		var range = AscCommonExcel.g_oRangeCache.getAscRange(sRange);
		var index = this.columnSort ? range.c1 - selection.c1 : range.r1 - selection.r1;
		this.sortList[index] = this.getNameColumnByIndex(index, selection);

		return index;
	};
	CSortProperties.prototype.getNameColumnByIndex = function (index, parentRef) {
		var t = this;
		var _generateName = function (index) {
			var base = t.columnSort ? AscCommon.translateManager.getValue("Column") : AscCommon.translateManager.getValue("Row");
			var text = t.columnSort ? t._ws._getColumnTitle(index) : t._ws._getRowTitle(index);
			text = base + " " + text;
			if (t.hasHeaders) {
				text = "(" + text + ")";
			}
			return text;
		};

		//columnSort; dataHasHeaders
		var row = this.columnSort ? parentRef.r1 : index + parentRef.r1;
		var col = !this.columnSort ? parentRef.c1 : index + parentRef.c1;
		//TODO проверить в 1 строке как должно работать
		if (this.hasHeaders) {
			if (this.columnSort) {
				row--;
			} else {
				col--;
			}
		}

		if (!this.hasHeaders) {
			return _generateName(this.columnSort ? col : row);
		} else {
			var cell = t._ws.model.getCell3(row, col);
			var value = cell.getValueWithFormat();
			return value !== "" ? value : _generateName(this.columnSort ? col : row);
		}
	};

	CSortProperties.prototype.asc_getLevelProps = function (index) {
		var t = this;

		var selection = t._newSelection;
		var r1 = this.columnSort ? selection.r1 : selection.r1 + index;
		var r2 = this.columnSort ? selection.r2 : selection.r1 + index;
		var c1 = this.columnSort ? selection.c1 + index : selection.c1;
		var c2 = this.columnSort ? selection.c1 + index : selection.c2;
		var range = new Asc.Range(c1, r1, c2, r2);

		var levelInfo;
		var rangeInfo = t._ws.model.getRowColColors(range, !this.columnSort, true);
		if (rangeInfo) {
			levelInfo = new CSortLevelInfo();
			levelInfo.colorsFill = rangeInfo.colors;
			levelInfo.colorsFont = rangeInfo.fontColors;

			levelInfo.isText = rangeInfo.text;
			levelInfo.isDate = rangeInfo.date;
		}

		return levelInfo;
	};

	CSortProperties.prototype.asc_getRangeStr = function () {
		return this._newSelection.getAbsName();
	};

	CSortProperties.prototype.asc_getSelection = function () {
		return this.selection;
	};

	CSortProperties.prototype.asc_setSelection = function (val) {
		this.selection = val;
	};

	function CSortPropertiesLevel() {
		this.index = null;
		this.name = null;

		this.sortBy = null;
		this.descending = null;
		this.color = null;

		return this;
	}

	CSortPropertiesLevel.prototype.asc_getIndex = function () {
		return this.index;
	};
	CSortPropertiesLevel.prototype.asc_getName = function () {
		return this.name;
	};
	CSortPropertiesLevel.prototype.asc_getSortBy = function () {
		return this.sortBy;
	};
	CSortPropertiesLevel.prototype.asc_getDescending = function () {
		return this.descending;
	};
	CSortPropertiesLevel.prototype.asc_getColor = function () {
		return this.color;
	};
	CSortPropertiesLevel.prototype.asc_setIndex = function (val) {
		this.index = val;
	};
	CSortPropertiesLevel.prototype.asc_setName = function (val) {
		this.name = val;
	};
	CSortPropertiesLevel.prototype.asc_setSortBy = function (val) {
		this.sortBy = val;
	};
	CSortPropertiesLevel.prototype.asc_setDescending = function (val) {
		this.descending = val;
	};
	CSortPropertiesLevel.prototype.asc_setColor = function (val) {
		this.color = val;
	};


	function CSortLevelInfo() {
		this.colorsFill = null;
		this.colorsFont = null;

		this.isText = null;
		this.isDate = null;

		return this;
	}

	CSortLevelInfo.prototype.asc_getColorsFill = function () {
		return this.colorsFill;
	};
	CSortLevelInfo.prototype.asc_getColorsFont = function () {
		return this.colorsFont;
	};
	CSortLevelInfo.prototype.asc_getIsTextData = function () {
		return this.isText;
	};
	CSortLevelInfo.prototype.asc_getIsDate = function () {
		return this.isDate;
	};

	function CRemoveDuplicatesProps(ws) {
		this.selection = null;
		this._newSelection = null;
		this.hasHeaders = null;

		this.columnList = null;

		this._ws = ws;

		this.duplicateValues = null;
		this.uniqueValues = null;

		return this;
	}

	CRemoveDuplicatesProps.prototype.asc_getHasHeaders = function () {
		return this.hasHeaders;
	};
	CRemoveDuplicatesProps.prototype.asc_getColumnList = function () {
		return this.columnList;
	};
	CRemoveDuplicatesProps.prototype.asc_setHasHeaders = function (val) {
		var oldVal = !!this.hasHeaders;
		if (this._newSelection && oldVal !== val) {
			if (val) {
				this._newSelection.r1++;
			} else {
				this._newSelection.r1--;
			}
			this._ws.setSelection(this._newSelection);
		}
		this.hasHeaders = val;
	};
	CRemoveDuplicatesProps.prototype.asc_updateColumnList = function () {
		//TODO change selection
		this.generateColumnList();
	};
	CRemoveDuplicatesProps.prototype.generateColumnList = function () {
		var maxCount = 500;
		var selection = this._newSelection;
		var j, elem;

		if (this.columnList && this.columnList.length) {
			for (j in this.columnList) {
				this.columnList[j].asc_setVal(this.getNameColumnByIndex(parseInt(j)));
			}
		} else {
			this.columnList = [];
			for (j = selection.c1; j <= selection.c2; j++) {
				if (j - selection.c1 >= maxCount) {
					break;
				}
				elem = new window["AscCommonExcel"].AutoFiltersOptionsElements();
				elem.asc_setVisible(true);
				elem.asc_setVal(this.getNameColumnByIndex(j - selection.c1));
				this.columnList.push(elem);
			}
		}
	};
	CRemoveDuplicatesProps.prototype.getNameColumnByIndex = function (index) {
		var t = this;
		var _generateName = function (index) {
			var base = AscCommon.translateManager.getValue("Column");
			var text = t._ws._getColumnTitle(index);
			text = base + " " + text;
			return text;
		};

		var row = this._newSelection.r1;
		var col = index + this._newSelection.c1;

		if (!this.hasHeaders) {
			return _generateName(col);
		} else {
			var cell = t._ws.model.getCell3(row - 1, col);
			var value = cell.getValueWithFormat();
			return value !== "" ? value : _generateName(col);
		}
	};
	CRemoveDuplicatesProps.prototype.setDuplicateValues = function (val) {
		this.duplicateValues = val;
	};
	CRemoveDuplicatesProps.prototype.setUniqueValues = function (val) {
		this.uniqueValues = val;
	};
	CRemoveDuplicatesProps.prototype.asc_getDuplicateValues = function (val) {
		return this.duplicateValues;
	};
	CRemoveDuplicatesProps.prototype.asc_getUniqueValues = function (val) {
		return this.uniqueValues;
	};

	function CFunctionInfo(name) {
		this.name = null;
		this.argumentsMin = null;
		this.argumentsMax = null;

		this.argumentsValue = null;
		this.argumentsType = null;
		this.argumentsResult = null;
		this.formulaResult = null;
		this.functionResult = null;

		this.arguments = null;

		this._init(name);

		return this;
	}

	CFunctionInfo.prototype._init = function (name) {
		var f = AscCommonExcel.cFormulaFunctionLocalized ? AscCommonExcel.cFormulaFunctionLocalized[name] : AscCommonExcel.cFormulaFunction[name];
		if (f) {
			this.name = name;
			this.argumentsMin = f.prototype.argumentsMin;
			this.argumentsMax = f.prototype.argumentsMax;
			this.argumentsType = f.prototype.argumentsType;
		}
	};
	CFunctionInfo.prototype.asc_getArgumentMin = function () {
		return this.argumentsMin;
	};
	CFunctionInfo.prototype.asc_getArgumentMax = function () {
		return this.argumentsMax;
	};
	CFunctionInfo.prototype.asc_getArgumentsValue = function () {
		return this.argumentsValue;
	};
	CFunctionInfo.prototype.asc_getArgumentsType = function () {
		return this.argumentsType;
	};
	CFunctionInfo.prototype.asc_getArgumentsResult = function () {
		return this.argumentsResult;
	};
	CFunctionInfo.prototype.asc_getFormulaResult = function () {
		return this.formulaResult;
	};
	CFunctionInfo.prototype.asc_getFunctionResult = function () {
		return this.functionResult;
	};
	CFunctionInfo.prototype.asc_getName = function () {
		return this.name;
	};
	CFunctionInfo.prototype.asc_getArguments = function () {
		return this.arguments;
	};
	CFunctionInfo.prototype.asc_setArguments = function (val) {
		this.arguments = val;
	};


	function CPrintPreviewState(wb) {
		this.ctx = null;
		this.pages = null;

		//отслеживаем активную страницу и активный лист
		this.activePage = null;
		this.activeSheet = null;
		this.sheetsProps = null;

		this.start = null;

		//зум для масштабирования страницы и зум печати
		this.pageZoom = null;
		this.printZoom = null;

		this.wb = wb;
		//после закрытия окна предварительной печати возвращаем зум и актиный лист
		this.realZoom = null;
		this.realActiveSheet = null;

		this.pixelRatio = null;

		this.advancedOptions = null;

		//избегаем повторных вызовов, пересмотреть
		this.isDrawPrintPreview = null;
		//избегаем повторного отображения ошибки(максимальное количество страниц)
		this.maxPagesCount = null;

		//опции измененного листа пока не мержу с теми, что в модели. перезатираю полностью. если будет необходимость - _pageOptionsMap использовать и при сохранении проверять что изменилось
		//при закрытии окна с сохранением вычисляем только измененные настройки, для этого храним те настройки, которые были до открытия
		//this._pageOptionsMap = null;

		return this;
	}

	CPrintPreviewState.prototype.init = function () {
		this.start = true;
		/*var api = window["Asc"]["editor"];
		this._pageOptionsMap = {};
		for (var i = 0, length = this.wb.model.getWorksheetCount(); i < length; ++i) {
			var wsModel = this.wb.model.getWorksheet(i);
			var wsIndex = wsModel.getIndex();
			var pageOptions = api.asc_getPageOptions(wsIndex, true);
			if (pageOptions) {
				this._pageOptionsMap[wsIndex] = pageOptions.clone();
			}
		}*/
	};
	CPrintPreviewState.prototype.isStart = function () {
		return this.start;
	};
	CPrintPreviewState.prototype.setCtx = function (val) {
		this.ctx = val;
	};
	CPrintPreviewState.prototype.setPages = function (val) {
		this.pages = val;
	};
	CPrintPreviewState.prototype.getCtx = function () {
		return this.ctx;
	};
	CPrintPreviewState.prototype.getPages = function () {
		return this.pages;
	};
	CPrintPreviewState.prototype.getPage = function (index) {
		return this.pages && this.pages.arrPages[index];
	};
	CPrintPreviewState.prototype.isNeedShowError = function (bMoreThenMax) {
		var res = bMoreThenMax;

		if (this.isStart()) {
			if (bMoreThenMax) {
				if (!this.maxPagesCount) {
					this.maxPagesCount = true;
				} else {
					res = false;
				}
			} else {
				this.maxPagesCount = null;
			}
		}

		return res;
	};
	CPrintPreviewState.prototype.clean = function (revertZoom) {
		//this.ctx = null;
		this.pages = null;
		this.activePage = null;
		this.activeSheet = null;
		this.start = null;

		if (revertZoom) {
			if (null != this.realActiveSheet) {
				this.wb.model.setActive(this.realActiveSheet);
			}
			if (null != this.realZoom) {
				this.wb.changeZoom(this.realZoom, true);
			}
		}
		this.realActiveSheet = null;
		this.realZoom = null;
		this._pageOptionsMap = null;

		this.advancedOptions = null;
		this.maxPagesCount = null;
	};
	CPrintPreviewState.prototype.getPagesLength = function () {
		return this.pages && this.pages.arrPages.length;
	};
	CPrintPreviewState.prototype.getIndexPageByIndexSheet = function (indexSheet) {
		if (this.pages && this.pages.arrPages.length) {
			for (var i = 0; i < this.pages.arrPages.length; i++) {
				if (indexSheet === this.pages.arrPages[i].indexWorksheet) {
					return i;
				}
			}
		}
		return null;
	};
	CPrintPreviewState.prototype.setPage = function (index, checkZoom) {
		this.activePage = index;
		var page = this.getPage(this.activePage);
		var newIndexSheet = page && page.indexWorksheet;
		var needUpdateActiveSheet;
		if (newIndexSheet !== this.activeSheet) {
			this.activeSheet = newIndexSheet;
			needUpdateActiveSheet = true;
		}
		if (checkZoom) {
			this.checkZoom(needUpdateActiveSheet);
		}
	};
	CPrintPreviewState.prototype.checkZoom = function (needUpdateActiveSheet) {
		if (null === this.realZoom) {
			this.realZoom = this.wb.getZoom();
		}
		if (null === this.realActiveSheet) {
			this.realActiveSheet = this.wb.model.getActive();
		}

		var isChangeSystemZoom = this.pixelRatio !== null && this.pixelRatio !== AscCommon.AscBrowser.retinaPixelRatio;

		var page = this.getPage(this.activePage);
		var pageWidth = page && page.pageWidth ? page.pageWidth : AscCommon.c_oAscPrintDefaultSettings.PageWidth;
		var pageHeight = page && page.pageHeight ? page.pageHeight : AscCommon.c_oAscPrintDefaultSettings.PageHeight;

		var canvasTopPadding = 24;
		var ppiX = AscCommon.AscBrowser.convertToRetinaValue(96, true);
		var height = Math.floor(pageHeight * Asc.getCvtRatio(3/*mm*/, 0/*px*/, ppiX));
		var width = Math.floor(pageWidth * Asc.getCvtRatio(3/*mm*/, 0/*px*/, ppiX));
		var canvasHeight = this.ctx.canvas.parentElement.clientHeight;
		var canvasWidth = this.ctx.canvas.parentElement.clientWidth;

		var kF = 1;
		if (height > canvasHeight) {
			kF = canvasHeight / height;
		}
		if (width * kF > canvasWidth) {
			kF = canvasWidth / width;
		}

		kF *= (height * kF) / (height * kF + canvasTopPadding);

		var isChangeForZoom;
		var trueZoom = kF * AscCommon.AscBrowser.retinaPixelRatio;
		var _height = Math.floor(height * kF);
		var _width = Math.floor(width * kF);
		var borderWidth = 1;
		if (trueZoom !== this.pageZoom) {
			this.pageZoom = trueZoom;
			this.ctx.canvas.style.height = _height + "px";
			this.ctx.canvas.style.width = _width + borderWidth * 2 + "px";
			this.ctx.canvas.height = AscCommon.AscBrowser.convertToRetinaValue(_height, true);
			this.ctx.canvas.width = AscCommon.AscBrowser.convertToRetinaValue(_width, true);
			isChangeForZoom = true;
		}
		this.ctx.canvas.style.marginLeft = Math.floor(canvasWidth/2 - _width / 2) + "px";
		let topMargin = Math.floor(canvasHeight/2 - _height / 2 + canvasTopPadding * kF);
		if (topMargin + _height > canvasHeight) {
			topMargin = canvasHeight - _height;
		}
		this.ctx.canvas.style.marginTop = topMargin + "px";


		kF = trueZoom;
		if (!page || page.scale !== this.printZoom) {
			this.printZoom = page ? page.scale : 1;
			this.pageZoom = kF;
			isChangeForZoom = true;
		}

		if (isChangeForZoom || needUpdateActiveSheet) {
			//change zoom on default
			if (needUpdateActiveSheet) {
				this.wb.model.setActive(this.activeSheet);
			}
			this.wb.changeZoom(this.pageZoom * this.printZoom, true);
			if (isChangeSystemZoom) {
				this.ctx.changeZoom(null);
			}
			this.ctx.changeZoom(this.pageZoom* this.printZoom);
		}
		var oGraphics = new AscCommon.CGraphics();
		var nWidth = this.ctx.getWidth(0);
		var nHeight = this.ctx.getHeight(0);
		var dWidth = this.ctx.getWidth(3);
		var dHeight = this.ctx.getHeight(3);
		oGraphics.init(this.ctx.canvas.getContext('2d'), nWidth, nHeight, dWidth, dHeight);
		oGraphics.m_oFontManager = AscCommon.g_fontManager;
		this.ctx.DocumentRenderer = oGraphics;
		this.pixelRatio = AscCommon.AscBrowser.retinaPixelRatio;

		if (isChangeForZoom && this.advancedOptions) {
			this.recalculatePageClipRect();
		}
	};
	CPrintPreviewState.prototype.recalculatePageClipRect = function () {
		var pages = this.wb.calcPagesPrint(this.advancedOptions);
		if (pages && pages.arrPages && this.pages.arrPages && this.pages.arrPages.length === pages.arrPages.length) {
			for (var i in this.pages.arrPages) {
				this.pages.arrPages[i].pageClipRectHeight = pages.arrPages[i].pageClipRectHeight;
				this.pages.arrPages[i].pageClipRectLeft = pages.arrPages[i].pageClipRectLeft;
				this.pages.arrPages[i].pageClipRectTop = pages.arrPages[i].pageClipRectTop;
				this.pages.arrPages[i].pageClipRectWidth = pages.arrPages[i].pageClipRectWidth;
			}
		}
	};
	CPrintPreviewState.prototype.setAdvancedOptions = function (val) {
		this.advancedOptions = val;
	};
	CPrintPreviewState.prototype.isNeedUpdate = function (ws, data) {
		if (!ws) {
			return false;
		}

		var maxCol = data ? data.col : ws.nColsCount;
		var maxRow = data ? data.row : ws.nRowsCount;
		var res = true;
		if (this.sheetsProps && this.sheetsProps[ws.index]) {
			if (this.sheetsProps[ws.index].col === maxCol && this.sheetsProps[ws.index].row === maxRow) {
				res = false;
			}
		}
		if (!this.sheetsProps) {
			this.sheetsProps = [];
		}
		if (!this.sheetsProps[ws.index]) {
			this.sheetsProps[ws.index] = {};
		}
		this.sheetsProps[ws.index].col = maxCol;
		this.sheetsProps[ws.index].row = maxRow;

		return res;
	};
	CPrintPreviewState.prototype.getActivePageSetup = function () {
		if (!this.isStart()) {
			return;
		}
		let activeSheet = this.activeSheet;
		let advancedOptions = this.advancedOptions;
		if (advancedOptions && advancedOptions.pageOptionsMap) {
			return advancedOptions.pageOptionsMap[activeSheet] && advancedOptions.pageOptionsMap[activeSheet].pageSetup;
		}
	};


	function CT_Connection() {
		this.dbPr = null;//CT_DbPr
		this.olapPr = null;//CT_OlapPr
		this.webPr = null;//CT_WebPr
		this.textPr = null;//CT_TextPr
		this.parameters = null;//CT_Parameters
		this.extLst = null;//CT_ExtensionList

		this.id = null;//xsd:unsignedInt
		this.sourceFile = null;//s:ST_Xstring
		this.odcFile = null;//s:ST_Xstring
		this.keepAlive = null;//xsd:boolean
		this.interval = null;//xsd:unsignedInt
		this.name = null;//s:ST_Xstring
		this.description = null;//s:ST_Xstring
		this.type = null;//xsd:unsignedInt
		this.reconnectionMethod = null;//xsd:unsignedInt
		this.refreshedVersion = null;//xsd:unsignedByte
		this.minRefreshableVersion = null;//xsd:unsignedByte
		this.savePassword = null;//xsd:boolean
		this.new = null;//xsd:boolean
		this.deleted = null;//xsd:boolean
		this.onlyUseConnectionFile = null;//xsd:boolean
		this.background = null;//xsd:boolean
		this.refreshOnLoad = null;//xsd:boolean
		this.saveData = null;//xsd:boolean
		this.credentials = null;//ST_CredMethod
		this.singleSignOnId = null;//s:ST_Xstring

		return this;
	}

	AscDFH.changesFactory[AscDFH.historyitem_OleSizeSelectionSetRange] = AscDFH.CChangesDrawingsObjectNoId;
	AscDFH.drawingsChangesMap[AscDFH.historyitem_OleSizeSelectionSetRange] = function (oClass, value) {
		oClass.resetHistory();
		oClass.applyRange(value);
		oClass.addPointToLocalHistory(true);
	};

	AscDFH.changesFactory[AscDFH.historyitem_Sparkline_Worksheet] = AscDFH.CChangesDrawingsString;
	AscDFH.drawingsChangesMap[AscDFH.historyitem_Sparkline_Worksheet] = function (oClass, value) {
		let oWB = Asc.editor.wbModel;
		if(!oWB) return;
		let oWS = oWB.getWorksheetById(value);
		oClass.worksheet = oWS;
	};
	/**
	 *
	 * @param ws
	 * @param range
	 * @constructor
	 * @extends {SelectionRange}
	 */
	function OleSizeSelectionRange(ws, range) {
		AscCommonExcel.SelectionRange.call(this, ws);
		this.Id = AscCommon.g_oIdCounter.Get_NewId();

		this.classType = AscDFH.historyitem_type_OleSizeSelection;
		this.init = false;
		if (range) {
			this.ranges = [range];
			this.activeCell = new AscCommon.CellBase(range.r1, range.c1);
		}
		this.localHistory = [];
		this.currentHistoryIndex = -1;
		AscCommon.g_oTableId.Add( this, this.Id );
		this.api = Asc.editor || editor;
	}
	OleSizeSelectionRange.prototype = Object.create(AscCommonExcel.SelectionRange.prototype);
	OleSizeSelectionRange.prototype.constructor = OleSizeSelectionRange;
	OleSizeSelectionRange.prototype.GetId = AscFormat.CBaseObject.prototype.Get_Id;
	OleSizeSelectionRange.prototype.Get_Id = AscFormat.CBaseObject.prototype.Get_Id;
	OleSizeSelectionRange.prototype.Write_ToBinary2 = AscFormat.CBaseNoIdObject.prototype.Write_ToBinary2;
	OleSizeSelectionRange.prototype.Read_FromBinary2 = AscFormat.CBaseNoIdObject.prototype.Read_FromBinary2;
	OleSizeSelectionRange.prototype.getObjectType = AscFormat.CBaseNoIdObject.prototype.getObjectType;

	OleSizeSelectionRange.prototype.addToGlobalHistory = function () {
		const oOldRange = this.getFirstFromLocalHistory();
		const oNewRange = this.getLastFromLocalHistory();

		this.resetHistory();
		this.addPointToLocalHistory(true);

		AscCommon.History.Create_NewPoint();
		AscCommon.History.Add(new AscDFH.CChangesDrawingsObjectNoId(this, AscDFH.historyitem_OleSizeSelectionSetRange, oOldRange, oNewRange));
	};

	OleSizeSelectionRange.prototype.validActiveCell = function () {
		return true;
	};
	OleSizeSelectionRange.prototype.addPointToLocalHistory = function (bSkipUpdate) {
		this.localHistory.length = this.currentHistoryIndex + 1;
		this.localHistory.push(this.getLast().clone());
		this.currentHistoryIndex = this.localHistory.length - 1;

		if (!bSkipUpdate) {
			this.updateUndoRedoChanged();
		}
	};

	OleSizeSelectionRange.prototype.updateUndoRedoChanged = function () {
		this.api.wb.cellEditor.handlers.trigger( "updateUndoRedoChanged", this.currentHistoryIndex > 0, this.currentHistoryIndex < this.localHistory.length - 1);
	};


	OleSizeSelectionRange.prototype.undo = function () {
		const ws = this.api.wb.getWorksheet();
		ws.cleanSelection();
		this.currentHistoryIndex = this.currentHistoryIndex > 0 ? this.currentHistoryIndex - 1 : 0;
		this.applyRange(this.localHistory[this.currentHistoryIndex]);
		ws._drawSelection();
		this.updateUndoRedoChanged();
	};

	OleSizeSelectionRange.prototype.getFirstFromLocalHistory = function () {
		return this.localHistory[0].clone();
	};

	OleSizeSelectionRange.prototype.getLastFromLocalHistory = function () {
		return this.localHistory[this.localHistory.length - 1].clone();
	};

	OleSizeSelectionRange.prototype.resetHistory = function () {
		this.localHistory = [];
		this.currentHistoryIndex = -1;
	};

	OleSizeSelectionRange.prototype.redo = function () {
		const ws = this.api.wb.getWorksheet();
		ws.cleanSelection();
		this.currentHistoryIndex = this.currentHistoryIndex < this.localHistory.length - 1 ? this.currentHistoryIndex + 1 : this.localHistory.length - 1;
		this.applyRange(this.localHistory[this.currentHistoryIndex]);
		ws._drawSelection();
		this.updateUndoRedoChanged();
	};

	OleSizeSelectionRange.prototype.applyRange = function (oRange) {
		this.ranges = [oRange.clone()];
		this.activeCellId = 0;
		this.activeCell = new AscCommon.CellBase(oRange.r1, oRange.c1);
	};

	OleSizeSelectionRange.prototype.clean = function () {
		this.ranges = [new Asc.Range(0, 0, 10, 10)];
		this.activeCellId = 0;
		this.activeCell.clean();
	};
	OleSizeSelectionRange.prototype.getName = function () {
		var range = this.getLast();
		return range.getName();
	};

	//external reference
	function ExternalReference() {
		this.DefinedNames = [];
		this.Id = null;
		this.SheetDataSet = [];
		this.SheetNames = [];
		this.Type = 0;

		//дополнительная информация, которая приходит при copy/paste
		//необходимо её добавлять в ooxml
		//fileId
		//portalName
		this.referenceData = null;

		//temp for update
		this.sKey = null;

		this.worksheets = {};
	}

	ExternalReference.prototype.getType = function() {
		return AscCommonExcel.UndoRedoDataTypes.externalReference;
	};

	ExternalReference.prototype.Read_FromBinary2 = function(r) {
		//TODO DefinedName
		var i;
		var length = r.GetLong();
		for (i = 0; i < length; ++i) {
			var definedName = new ExternalDefinedName(this);
			definedName.Read_FromBinary2(r);
			if(!this.DefinedNames) {
				this.DefinedNames = [];
			}
			this.DefinedNames.push(definedName);
		}

		if (r.GetBool()) {
			this.Id = r.GetString2();
		}

		length = r.GetLong();
		for (i = 0; i < length; ++i) {
			var sheetDataSet = new ExternalSheetDataSet();
			sheetDataSet.Read_FromBinary2(r);
			if(!this.SheetDataSet) {
				this.SheetDataSet = [];
			}
			this.SheetDataSet.push(sheetDataSet);
		}

		//TODO SheetName
		length = r.GetLong();
		for (i = 0; i < length; ++i) {
			var sheetName = r.GetString2();
			if(!this.SheetNames) {
				this.SheetNames = [];
			}
			this.SheetNames.push(sheetName);
		}

		if (r.GetBool()) {
			this.Type = r.GetLong();
		}

		if (r.GetBool()) {
			this.referenceData = {};
			if (r.GetBool()) {
				this.referenceData["fileKey"] = r.GetString2();
			}
			if (r.GetBool()) {
				this.referenceData["instanceId"] = r.GetString2();
			}
		}
	};
	ExternalReference.prototype.Write_ToBinary2 = function(w) {
		var i;
		w.WriteLong(this.DefinedNames ? this.DefinedNames.length : 0);
		if (this.DefinedNames) {
			for (i = 0; i < this.DefinedNames.length; ++i) {
				this.DefinedNames[i].Write_ToBinary2(w);
			}
		}

		if (null != this.Id) {
			w.WriteBool(true);
			w.WriteString2(this.Id);
		} else {
			w.WriteBool(false);
		}

		w.WriteLong(this.SheetDataSet ? this.SheetDataSet.length : 0);
		if (this.SheetDataSet) {
			for (i = 0; i < this.SheetDataSet.length; ++i) {
				this.SheetDataSet[i].Write_ToBinary2(w);
			}
		}

		w.WriteLong(this.SheetNames ? this.SheetNames.length : 0);
		if (this.SheetNames) {
			for (i = 0; i < this.SheetNames.length; ++i) {
				w.WriteString2(this.SheetNames[i]);
			}
		}

		if (null != this.Type) {
			w.WriteBool(true);
			w.WriteLong(this.Type);
		} else {
			w.WriteBool(false);
		}

		if (null != this.referenceData) {
			w.WriteBool(true);
			if (null != this.referenceData["fileKey"]) {
				w.WriteBool(true);
				w.WriteString2(this.referenceData["fileKey"]);
			} else {
				w.WriteBool(false);
			}
			if (null != this.referenceData["instanceId"]) {
				w.WriteBool(true);
				w.WriteString2(this.referenceData["instanceId"]);
			} else {
				w.WriteBool(false);
			}
		} else {
			w.WriteBool(false);
		}
	};

	ExternalReference.prototype.clone = function (needCloneSheets) {
		var newObj = new ExternalReference();

		if (this.DefinedNames) {
			for (var i = 0; i < this.DefinedNames.length; i++) {
				newObj.DefinedNames.push(this.DefinedNames[i].clone());
			}
		}

		newObj.Id = this.Id;

		if (this.SheetDataSet) {
			for (i = 0; i < this.SheetDataSet.length; ++i) {
				newObj.SheetDataSet.push(this.SheetDataSet[i].clone());
			}
		}

		newObj.Type = this.Type;

		if (null != this.referenceData) {

			newObj.referenceData = {};
			newObj.referenceData["fileKey"] = this.referenceData["fileKey"];
			newObj.referenceData["instanceId"] = this.referenceData["instanceId"];
		}

		if (needCloneSheets) {
			for (let i = 0; i < this.SheetNames.length; i++) {
				newObj.SheetNames.push(this.SheetNames[i]);
			}
			for (let i in this.worksheets) {
				newObj.worksheets[i] = this.worksheets[i];
			}
		}

		return newObj;
	};

	ExternalReference.prototype.initPostOpen = function () {
		let res = null;
		if (this.DefinedNames && this.DefinedNames.length) {
			for (let i = 0; i < this.DefinedNames.length; i++) {
				this.DefinedNames[i].parent = this;
			}
		}

		let api = Asc.editor || editor;
		let originalWb = api.wbModel;
		originalWb && originalWb.dependencyFormulas.lockRecal();

		this.initWorksheetsFromSheetDataSet();
		this.initWorkbook();
		this.prepareDefNames();

		originalWb && originalWb.dependencyFormulas.unlockRecal();

		return res;
	};

	ExternalReference.prototype.getDefinedNamesBySheetIndex = function (index, wb) {
		let res = null;
		if (this.DefinedNames && this.DefinedNames.length) {
			for (let i = 0; i < this.DefinedNames.length; i++) {
				let defnameInERName = this.DefinedNames[i].Name;

				// we get a workbook scope defname and compare the sheet where it was originally written with the current sheet name
				// if it's the same sheet, add defname to array for further dataset update
				let wbDefName = wb && wb.getDefinesNames(defnameInERName);
				let defnameArea3D = wbDefName && wbDefName.parsedRef && wbDefName.parsedRef.outStack && wbDefName.parsedRef.outStack[0];
				let defnameWorksheet = defnameArea3D && defnameArea3D.getWS && defnameArea3D.getWS();
				let wsFromName = defnameWorksheet && defnameWorksheet.getName();

				if (wb && wsFromName === this.SheetNames[index]) {
					if (!res) {
						res = [];
					}
					res.push(this.DefinedNames[i]);
					continue;
				}

				if (this.DefinedNames[i].SheetId === index) {
					if (!res) {
						res = [];
					}
					res.push(this.DefinedNames[i]);
				}
			}
		}
		return res;
	};

	ExternalReference.prototype.removeSheetByName = function (sheetName, workbook) {
		if (sheetName && workbook) {
			let index = this.getSheetByName(sheetName);
			let idInWb = this.worksheets[sheetName].getId();
			if (index != null) {
				let wb = this.getWb();
				/*
					when deleting a sheet from ER, we also need to check all defnames and their listeners, since when creating a short link,
					we create a non-existent ws referenced in cName3D before receiving the promise with real data
				*/
				for (let i = 0; i < this.DefinedNames.length; i++) {
					let defname = this.DefinedNames[i];
					let defnameFromWorkbook = wb && wb.getDefinesNames(defname.Name);
					if (defnameFromWorkbook) {
						let defnameArea3D = defnameFromWorkbook.parsedRef && defnameFromWorkbook.parsedRef.outStack && defnameFromWorkbook.parsedRef.outStack[0];
						let defnameWorksheet = defnameArea3D && defnameArea3D.getWS && defnameArea3D.getWS();

						let wsToSwitch = this.worksheets[defnameWorksheet.sName];

						// we get all the listeners of this defname and rewrite the sheet to the one on which the listener is located
						let defNameDepInfo = workbook && workbook.dependencyFormulas && workbook.dependencyFormulas.defNameListeners && workbook.dependencyFormulas.defNameListeners[defname.Name];
						if (wsToSwitch && defNameDepInfo && defNameDepInfo.listeners) {
							let defNameListeners = defNameDepInfo.listeners;
							for (let id in defNameListeners) {
								let formula = defNameListeners[id];
								// let parent = formula.parent;

								let outStack = formula.outStack;
								if (outStack) {
									for (let j = 0; j < outStack.length; j++) {
										let elem = outStack[j];
										if (elem.type === AscCommonExcel.cElementType.name3D && elem.ws.sName === sheetName) {
											// change ws in outstack
											outStack[j] = new AscCommonExcel.cName3D(elem.value, wsToSwitch ? wsToSwitch : elem.ws, elem.externalLink, elem.shortLink);
										}
									}
								}

								// add to build dependencies according to new data in outstack
								formula.isInDependencies = false;
								formula.buildDependencies();
							}
						}

						// if the range has a workbook scope, delete sheetId and write a new refInfo(RefersTo)
						this.DefinedNames[i].SheetId = null;
						this.DefinedNames[i].RefersTo = defnameFromWorkbook.getRef();
					}
				}

				// delete SheetListener from wb
				delete workbook.dependencyFormulas.sheetListeners[idInWb];

				// remove all data from er associated with this sheet
				this.SheetNames.splice(index, 1);
				this.SheetDataSet.splice(index, 1);
				delete this.worksheets[sheetName];

				// shift all dataset indexes 
				this.shiftData();
			}
		}
	};
	
	ExternalReference.prototype.updateData = function (arr, oPortalData, noData, workbook) {
		var t = this;
		var isChanged = false;
		var cloneER = this.clone();
		
		let existedWsArray = [];
		for (var i = 0; i < arr.length; i++) {
			//если есть this.worksheets, если нет - проверить и обработать
			var sheetName = arr[i].sName;
			if (this.worksheets[null]) {
				this.changeSheetName(null, sheetName);
			}
			if (!this.worksheets[sheetName]) {
				this.addSheetName(sheetName, true, true);
			}
			if (this.worksheets && this.worksheets[sheetName]) {
				let wsTo = this.worksheets[sheetName];
				//меняем лист
				AscFormat.ExecuteNoHistory(function(){
					AscCommonExcel.executeInR1C1Mode(false, function () {

						let defNames = wsTo.workbook && wsTo.workbook.dependencyFormulas && wsTo.workbook.dependencyFormulas.defNames;
						wsTo.workbook.dependencyFormulas._foreachDefName(function (defName) {
							var api_sheet = Asc['editor'];
							var wb = api_sheet.wbModel;
							let realWb = defName.wb;
							defName.wb = wb;
							defName.onFormulaEvent(AscCommon.c_oNotifyParentType.Change);
							defName.wb = realWb;
						});

						var oAllRange = wsTo.getRange3(0, 0, wsTo.getRowsCount(), wsTo.getColsCount());
						oAllRange.cleanAll();
						wsTo.copyFrom(arr[i], wsTo.sName);
					});
				});
				//this.worksheets[sheetName] = arr[i];

				//обновляем данные в SheetDataSet
				var index = this.getSheetByName(sheetName);
				if (index != null) {
					var externalSheetDataSet = this.SheetDataSet[index];
					if (externalSheetDataSet) {
						if (externalSheetDataSet.updateFromSheet(t.worksheets[sheetName], noData)) {
							isChanged = true;
						}
					}
					const eWB = t.getWb();
					let externalDefName = this.getDefinedNamesBySheetIndex(index, eWB);
					if (externalDefName) {
						for (let i = 0; i < externalDefName.length; i++) {
							if (externalDefName[i].updateFromSheet(t.worksheets[sheetName], noData)) {
								isChanged = true;
							}
						}
					}
				}
			}
			existedWsArray.push(sheetName);
		}

		// delete all non-existent sheets in ExternalReference
		for (let wsName in this.worksheets) {
			if (!existedWsArray.includes(wsName)) {
				// throw an error if we referenced to one of the deleted sheets?
				this.removeSheetByName(wsName, workbook);
			}
		}

		var oReferenceData = oPortalData && oPortalData["referenceData"];
		//data from portal, need update reference data
		if (oReferenceData && (!this.referenceData || (this.referenceData["instanceId"] !== oReferenceData["instanceId"] || this.referenceData["fileKey"] !== oReferenceData["fileKey"]))) {
			this.setReferenceData(oReferenceData["fileKey"], oReferenceData["instanceId"]);
			isChanged = true;
		}
		//path also can changed
		var path = oPortalData && oPortalData["path"];
		if (path && this.Id !== path) {
			let isNotUpdate = (AscCommonExcel.importRangeLinksState && AscCommonExcel.importRangeLinksState.notUpdateIdMap && AscCommonExcel.importRangeLinksState.notUpdateIdMap[this.Id]) || this.notUpdateId;
			!isNotUpdate && this.setId(path);
			isChanged = true;
		}

		if (isChanged && AscCommon.History.Is_On()) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_ChangeExternalReference,
				null, null, new AscCommonExcel.UndoRedoData_FromTo(cloneER, this));
		}
	};

	ExternalReference.prototype.changeSheetName = function (from, to) {
		for (let i in this.SheetNames) {
			if (this.SheetNames[i] === from) {
				this.SheetNames[i] = to;
			}
		}

		for (let i in this.worksheets) {
			if (i === from + "") {
				let _val = this.worksheets[i];
				_val.sName = to;
				this.worksheets[to] = _val;
				delete this.worksheets[from];
			}
		}
	};

	ExternalReference.prototype.putToChangedCells = function () {
		for (let i in this.SheetDataSet) {
			if (this.SheetDataSet.hasOwnProperty(i)) {
				let sheetName = this.SheetNames && this.SheetNames[this.SheetDataSet[i].SheetId];
				this.SheetDataSet[i].putToChangedCells(this.worksheets && this.worksheets[sheetName]);
			}
		}
	};

	//TODO внешние источники данных, как в файле из бага https://bugzilla.onlyoffice.com/show_bug.cgi?id=38646

	ExternalReference.prototype.getAscLink = function () {

		// вот так, если это из файла прилетело, в т.ч. из буфера
		// onRequestReferenceData({data:{referenceData:config.document.referenceData}})
		//
		//
		// вот так, если это будет ссылка на редактор файла в тестовом как в onedrive
		// onRequestReferenceData({data:{link:"http://192.168.1.1/editor?fileName=new.docx"}})
		//
		// вот так, если б это было просто путь к файлу как в MS:
		// 	onRequestReferenceData({data:{path: "new.docx"}})


		var res = new asc_CExternalReference();

		if (this.referenceData) {
			res.type = Asc.c_oAscExternalReferenceType.referenceData;
			res.data = this.referenceData;
		} else if (this.isExternalLink()) {
			res.type = Asc.c_oAscExternalReferenceType.link;
			res.data = this.Id;
		} else {
			res.type = Asc.c_oAscExternalReferenceType.path;
			res.data = this.Id;
		}

		res.externalReference = this;

		return res;
	};

	ExternalReference.prototype.isExternalLink = function() {
		var p = /^(?:http:\/\/|https:\/\/)/;
		return this.Id.match(p);
	};

	ExternalReference.prototype.isXlsx = function() {
		var p = /^.*\.(xlsx)$/i;
		return this.Id.match(p);
	};

	ExternalReference.prototype.addSheetName = function (name, generateDefaultStructure, addSheetObj) {
		this.SheetNames.push(name);
		if (generateDefaultStructure) {
			var externalSheetDataSet = new ExternalSheetDataSet();
			externalSheetDataSet.SheetId = this.SheetNames.length - 1;
			this.SheetDataSet.push(externalSheetDataSet);
		}
		if (addSheetObj) {
			let wb = this.getWb();
			if (!wb) {
				wb = new AscCommonExcel.Workbook(null, window["Asc"]["editor"], false);
			}
			let ws = new AscCommonExcel.Worksheet(wb);
			ws.sName = name;
			this.worksheets[name] = ws;
		}
	};

	ExternalReference.prototype.getSheetDataSetByName = function (name) {
		if (this.SheetNames) {
			for (var i = 0; i < this.SheetNames.length; i++) {
				if (this.SheetNames[i] === name) {
					return this.SheetDataSet[i];
				}
			}
		}
		return null;
	};

	ExternalReference.prototype.addSheet = function (sheet, ranges) {
		//добавляем нужные данные из вставляемого листа в SheetDataSet
		var sheetName = sheet && sheet.sName;
		if (sheetName && null === this.getSheetByName(sheetName)) {
			//делаем init и формируем на основании данных из SheetDataSet
			//добавляем в this.worksheets
			this.addSheetName(sheetName, true);

			//TODO при копипасте есть нюанас - те данные, которые не были скопированы из исходного документа, подхватить не сможем
			//либо вставку менять - и делать не так как мс(в случае ссылки вставляеть всегда диапазон, равный скопированному). сейчас - если выделили больше, то подхватиться этот диапазон.

			this.updateSheetData(sheetName, sheet, ranges);
		}
	};

	ExternalReference.prototype.updateSheetData = function (sheetName, sheet, ranges) {
		var index = this.getSheetByName(sheetName);
		var externalSheetDataSet = this.SheetDataSet[index];
		if (!externalSheetDataSet) {
			this.addSheetName(sheetName, true);
			externalSheetDataSet = this.getSheetDataSetByName(sheetName);
		}
		externalSheetDataSet.initFromSheet(sheet, ranges);

		//из SheetDataSet данные добавляем во временный лист, который хранится в worksheets
		this.initWorksheetFromSheetDataSet(sheetName);
	};


	ExternalReference.prototype.initWorksheetsFromSheetDataSet = function () {
		if (this.SheetNames) {
			for (var i = 0; i < this.SheetNames.length; i++) {
				this.initWorksheetFromSheetDataSet(this.SheetNames[i]);
			}
		}
	};

	ExternalReference.prototype.prepareDefNames = function () {
		let wb = this.getWb();
		if (wb && wb.dependencyFormulas && wb.dependencyFormulas.defNames && wb.dependencyFormulas.defNames.wb) {
			for (let i in wb.dependencyFormulas.defNames.wb) {
				let defName = wb.dependencyFormulas.defNames.wb[i];
				defName.parsedRef.parse();
			}
		}
	};

	ExternalReference.prototype.initWorksheets = function () {
		if (this.SheetNames) {
			for (var i = 0; i < this.SheetNames.length; i++) {
				this.initWorksheet(this.SheetNames[i]);
			}
		}
	};
	
	ExternalReference.prototype.initWorksheet = function (sheetName) {
		var ws = this.worksheets[sheetName];
		if (!this.worksheets[sheetName]) {
			var wb = this.getWb();
			if (!wb) {
				wb = new AscCommonExcel.Workbook(null, window["Asc"]["editor"], false);
			}
			ws = new AscCommonExcel.Worksheet(wb);
			ws.sName = sheetName;

			this.worksheets[sheetName] = ws;
		}
	};

	ExternalReference.prototype.initWorksheetFromSheetDataSet = function (sheetName) {
		var sheetDataSetIndex = this.getSheetByName(sheetName);
		if (null !== sheetDataSetIndex) {

			var sheetDataSet = this.SheetDataSet[sheetDataSetIndex];
			var ws = this.worksheets[sheetName];
			if (!this.worksheets[sheetName]) {
				var wb = this.getWb();
				if (!wb) {
					wb = new AscCommonExcel.Workbook(null, window["Asc"]["editor"], false);
				}
				ws = new AscCommonExcel.Worksheet(wb, wb.aWorksheets.length);
				ws.sName = sheetName;

				this.worksheets[sheetName] = ws;
				wb.aWorksheets.push(ws);
			}


			//клонируем все данные из SheetDataSet в данный темповый Worksheet
			if (!sheetDataSet || !sheetDataSet.Row) {
				return;
			}
			let api = Asc.editor || editor;
			let originalWb = api.wbModel;
			let isLockRecalc = false;
			for (var i = 0; i < sheetDataSet.Row.length; i++) {
				if (!sheetDataSet.Row[i] || !sheetDataSet.Row[i].Cell) {
					continue;
				}
				for (var j = 0; j < sheetDataSet.Row[i].Cell.length; j++) {
					if (! sheetDataSet.Row[i].Cell[j]) {
						continue;
					}
					// this.Ref = null;
					// this.CellType = null;
					// this.CellValue = null;
					AscFormat.ExecuteNoHistory(function(){
						AscCommonExcel.executeInR1C1Mode(false, function () {
							if (!isLockRecalc) {
								originalWb && originalWb.dependencyFormulas.lockRecal();
								isLockRecalc = true;
							}
							var range = ws.getRange2(sheetDataSet.Row[i].Cell[j].Ref);
							range.setValue(sheetDataSet.Row[i].Cell[j].CellValue);
						});
					});
				}
			}
			if (isLockRecalc) {
				originalWb && originalWb.dependencyFormulas.unlockRecal();
			}
		}
	};

	ExternalReference.prototype.initWorkbook = function () {
		if (this.DefinedNames) {
			let wb = this.getWb();
			for (let i = 0; i < this.DefinedNames.length; i++) {
				let defName = this.DefinedNames[i];
				let ws = this.getSheetByIndex(defName.SheetId);
				if (!ws && defName.RefersTo) {
					// try to find sheetname by RefersTo string
					let exclamationMarkIndex = defName.RefersTo.lastIndexOf("!");
					if (exclamationMarkIndex !== -1) {
						let sheetNamePart = defName.RefersTo.slice(0, exclamationMarkIndex);
						// remove equal sign
						if (sheetNamePart[0] === "=") {
							sheetNamePart = sheetNamePart.substring(1);
						}

						// regex to find string enclosed in single qoutes
						let regex = /^'(.*)'$/;
						let match = regex.exec(sheetNamePart);
						if (match && match[1]) {
							sheetNamePart = match[1];
						}

						ws = this.worksheets[sheetNamePart];
					}
				}

				if (ws != null) {
					//on parse name3d use g_DefNameWorksheet
					let RealDefNameWorksheet = AscCommonExcel.g_DefNameWorksheet;
					AscCommonExcel.g_DefNameWorksheet = ws;
					let stringToParse;
					if (defName && defName.RefersTo && defName.RefersTo[0] === "=") {
						stringToParse = defName.RefersTo.substring(1);
					}
					let oDefName = new Asc.asc_CDefName(defName.Name, stringToParse ? stringToParse : defName.RefersTo);
					wb && wb.editDefinesNames(null, oDefName);
					AscCommonExcel.g_DefNameWorksheet = RealDefNameWorksheet;	
				}
			}
		}
	};

	ExternalReference.prototype.getSheetByName = function (val) {
		for (var i = 0; i < this.SheetNames.length; i++) {
			//если есть this.worksheets, если нет - проверить и обработать
			if (this.SheetNames[i] === val) {
				return i;
			}
		}
		return null;
	};

	ExternalReference.prototype.getSheetByIndex = function (val) {
		let sheetName = this.SheetNames[val];
		if (sheetName != null) {
			return this.worksheets && this.worksheets[sheetName];
		}
		return null;
	};

	ExternalReference.prototype.getSheetBySheetId = function (sheetId) {
		if (this.worksheets) {
			for (var i in this.worksheets) {
				//если есть this.worksheets, если нет - проверить и обработать
				if (this.worksheets[i].Id === sheetId) {
					return i;
				}
			}
		}

		return null;
	};

	ExternalReference.prototype.getWb = function () {
		if (this.worksheets) {
			for (let i in this.worksheets) {
				//если есть this.worksheets, если нет - проверить и обработать
				if (this.worksheets[i]) {
					return this.worksheets[i].workbook;
				}
			}
		}

		return null;
	};

	ExternalReference.prototype.initRows = function (range) {

		var sheetName = range && range.worksheet && range.worksheet.sName;
		if (sheetName !== undefined) {
			var index = this.getSheetByName(sheetName);
			if (index != null) {
				var externalSheetDataSet = this.SheetDataSet[index];
				if (!externalSheetDataSet) {
					externalSheetDataSet = new ExternalSheetDataSet();
					externalSheetDataSet.SheetId = this.SheetNames.length - 1;
					this.SheetDataSet.push(externalSheetDataSet);
				}

				for (var i = range.bbox.r1; i <= range.bbox.r2; i++) {
					var row = externalSheetDataSet.getRow(i + 1, true);
					for (var j = range.bbox.c1; j <= range.bbox.c2; j++) {
						row.getCell(j, true);
					}
				}
			}
		}
	};

	ExternalReference.prototype.initDefinedName = function (val) {
		if (!val) {
			return;
		}

		const index = this.getSheetByName(val.ws.sName);
		const name = val.value;

		//check on exist
		// if (this.getDefName(name, index)) {
		// 	return;
		// }
		if (this.getDefNameByName(name)) {
			return;
		}

		let defName = new ExternalDefinedName(this);
		defName.Name = name;
		defName.SheetId = val.shortLink ? null : index;
		this.addDefName(defName);
	};

	ExternalReference.prototype.addDefName = function (defName) {
		this.DefinedNames.push(defName);
	};

	ExternalReference.prototype.getDefNameByName = function (name) {
		for (let i in this.DefinedNames) {
			if (this.DefinedNames[i] && this.DefinedNames[i].Name === name) {
				return this.DefinedNames[i];
			}
		}
		return null;
	};

	ExternalReference.prototype.getDefName = function (name, sheetId) {
		for (let i in this.DefinedNames) {
			if (this.DefinedNames[i] && this.DefinedNames[i].SheetId === sheetId && this.DefinedNames[i].Name === name) {
				return this.DefinedNames[i];
			}
		}
		return null;
	};

	ExternalReference.prototype.removeSheetById = function (sheetId) {
		if (sheetId != null) {
			var sheetName = this.getSheetBySheetId(sheetId);
			var index = this.getSheetByName(sheetName);
			if (index != null) {
				this.SheetNames.splice(index, 1);
				this.SheetDataSet.splice(index, 1);
				delete this.worksheets[sheetName];
			}
		}
	};

	ExternalReference.prototype.setReferenceData = function (fileId, portalName) {
		if (!fileId || !portalName) {
			return;
		}
		if (!this.referenceData) {
			this.referenceData = {};
		}
		this.referenceData["instanceId"] = portalName;
		this.referenceData["fileKey"] = fileId + "";
	};

	ExternalReference.prototype.setId = function (id) {
		if (!id) {
			return;
		}

		this.Id = id;
	};

	ExternalReference.prototype.initFromObj = function (obj) {
		//directUrl:
		//fileType:
		//token:
		//url
		//path
		//referenceData
		if (obj["path"] !== this.Id) {
			this.setId(this._checkAndCorrectPath(obj["path"], obj["filePath"]));
		}

		if (obj["referenceData"] && (!this.referenceData || this.referenceData["instanceId"] !== obj["referenceData"]["instanceId"] ||
			this.referenceData["instanceId"] !== obj["referenceData"]["fileKey"])) {
			this.setReferenceData(obj["referenceData"]["fileKey"], obj["referenceData"]["instanceId"]);
		}
	};

	ExternalReference.prototype._checkAndCorrectPath = function (sPath, sAbsolutePath) {
		if (!sPath || 1 === sPath.indexOf("../")) {
			// sPath -> ../../from.xlsx
			//sAbsolutePath - > C:\root\from.xlsx
			// need -> /root/from.xlsx
			if (sAbsolutePath) {
				sPath = sAbsolutePath.substring(sAbsolutePath.indexOf("\\"))
				sPath = sPath.replace(/\\/g,"/")
			}
		} else if (sPath && -1 !== sPath.indexOf(":/")) {
			// sPath -> C:/root/from1.xlsx
			//need -> file:///C:\root\from1.xlsx
			sPath = sPath.replace(/\//g,"\\");
			sPath = "file:///" + sPath;
		}

		return sPath;
	};

	ExternalReference.prototype.addDataSetFrom = function (eR) {
		if (!eR.SheetDataSet) {
			return;
		}
		for (let i = 0; i < eR.SheetDataSet.length; i++) {
			let _sheetId = eR.SheetDataSet[i].SheetId;
			let sheetName = eR.SheetNames[_sheetId];
			if (sheetName) {
				let sheetDataSet = this.getSheetDataSetByName(sheetName);
				if (sheetDataSet) {
					//add new from eR to this
					sheetDataSet.addFrom(eR.SheetDataSet[i])
				} else {
					//add new structure

				}
			}
		}
	};

	ExternalReference.prototype.getKey = function() {
		return this.sKey;
	};
	ExternalReference.prototype.shiftData = function () {
		/* shift data to 1 position left */
		for (let i in this.SheetDataSet) {
			let dataSet = this.SheetDataSet[i];
			if (dataSet.SheetId !== 0) {
				dataSet.SheetId--;
			}
		}
	};

	ExternalReference.prototype.setKey = function(val) {
		this.sKey = val;
	};

	function asc_CExternalReference() {
		this.type = null;
		this.data = null;

		this.externalReference = null;
	}
	asc_CExternalReference.prototype.asc_getType = function () {
		return this.type;
	};
	asc_CExternalReference.prototype.asc_getData = function () {
		return this.data;
	};
	asc_CExternalReference.prototype.asc_getLink = function () {
		let res = null;
		//if link on portal(example - importRange function)
		if (this.externalReference && this.externalReference.isExternalLink()) {
			res = this.externalReference && this.externalReference.Id;
		}
		return res;
	};
	asc_CExternalReference.prototype.asc_getSource = function () {
		let id = this.externalReference && this.externalReference.Id;
		if (id) {
			let diskRegex = /^[a-zA-Z]\:/gi; 
			let lastIndex = 0 === id.indexOf("file:///") || id.match(diskRegex) ? id.lastIndexOf('\\') : id.lastIndexOf('/');

			if (lastIndex === -1) {
				lastIndex = id.lastIndexOf('/\/');
			}
			return lastIndex === -1 ? id : id.substr(lastIndex + 1);
		}
		return null;
	};
	asc_CExternalReference.prototype.asc_getPath = function () {
		return this.externalReference && this.externalReference.Id;
	};
	asc_CExternalReference.prototype.asc_getLocation = function () {

	};
	asc_CExternalReference.prototype.isExternalLink = function () {
		return this.type === Asc.c_oAscExternalReferenceType.link;
	};
	asc_CExternalReference.prototype.asc_getId = function () {
		return this.externalReference && this.externalReference.Id;
	};


	function ExternalSheetDataSet() {
		this.SheetId = null;
		this.RefreshError = null;
		this.Row = [];
	}

	ExternalSheetDataSet.prototype.Read_FromBinary2 = function(r) {
		if (r.GetBool()) {
			this.SheetId = r.GetLong();
		}
		if (r.GetBool()) {
			this.RefreshError = r.GetLong();
		}

		var length = r.GetLong();
		for (var i = 0; i < length; ++i) {
			var row = new ExternalRow();
			row.Read_FromBinary2(r);
			if(!this.Row) {
				this.Row = [];
			}
			this.Row.push(row);
		}
	};
	ExternalSheetDataSet.prototype.Write_ToBinary2 = function(w) {
		if (null != this.SheetId) {
			w.WriteBool(true);
			w.WriteLong(this.SheetId);
		} else {
			w.WriteBool(false);
		}

		if (null != this.RefreshError) {
			w.WriteBool(true);
			w.WriteLong(this.RefreshError);
		} else {
			w.WriteBool(false);
		}

		w.WriteLong(this.Row ? this.Row.length : 0);
		if (this.Row) {
			for (var i = 0; i < this.Row.length; ++i) {
				this.Row[i].Write_ToBinary2(w);
			}
		}
	};

	ExternalSheetDataSet.prototype.clone = function () {
		var newObj = new ExternalSheetDataSet();

		newObj.SheetId = this.SheetId;
		newObj.RefreshError = this.RefreshError;

		if (this.Row) {
			for (var i = 0; i < this.Row.length; ++i) {
				newObj.Row.push(this.Row[i].clone());
			}
		}

		return newObj;
	};

	ExternalSheetDataSet.prototype.initFromSheet = function(sheet, ranges) {
		if (sheet && ranges) {
			var t = this;

			var addedRowMap = [];
			for (var i = 0; i < ranges.length; i++) {
				var range = sheet.getRange3(ranges[i].r1, ranges[i].c1, ranges[i].r2, ranges[i].c2);
				range._foreach(function (cell) {
					if (!addedRowMap[cell.nRow]) {
						var row = new ExternalRow();
						row.R = cell.nRow + 1;
						t.Row.push(row);
						addedRowMap[cell.nRow] = {index: t.Row.length - 1, cells: []};
					}
					if (addedRowMap[cell.nRow].cells && !addedRowMap[cell.nRow].cells[cell.nCol]) {
						var externalCell = new ExternalCell();
						externalCell.initFromCell(cell);

						var index = addedRowMap[cell.nRow].index;
						t.Row[index].Cell.push(externalCell);
						addedRowMap[cell.nRow].cells[cell.nCol] = 1;
					}
				});
			}
		}
	};

	ExternalSheetDataSet.prototype.updateFromSheet = function(sheet, noData) {
		var isChanged = false;
		if (sheet) {
			var t = this;

			//TODO пока обновлю ячейки по одной, в дальнейшем нужно объединить ячейки в диапазоны
			for (var i = 0; i < this.Row.length; i++) {
				var row = this.Row[i];
				if (!row) {
					continue;
				}
				for (var j = 0; j < this.Row[i].Cell.length; j++) {
					var externalCell = this.Row[i].Cell[j];
					if (!externalCell) {
						continue;
					}
					var range = sheet.getRange2(externalCell.Ref);
					range._foreach(function (cell) {

						let changedCell = externalCell.initFromCell(cell, true, noData);
						if (!isChanged) {
							isChanged = changedCell;
						}

						var api_sheet = Asc['editor'];
						var wb = api_sheet.wbModel;
						
						/* if we haven't received data from an external source, put #REF error for all cells */
						if (noData) {
							cell._setValue("#REF!");
						}

						wb.dependencyFormulas.addToChangedCell(cell);
					});
				}
			}
		}
		return isChanged;
	};

	ExternalSheetDataSet.prototype.putToChangedCells = function(sheet) {
		if (!sheet) {
			return;
		}
		for (var i = 0; i < this.Row.length; i++) {
			var row = this.Row[i];
			if (!row) {
				continue;
			}
			for (var j = 0; j < this.Row[i].Cell.length; j++) {
				var externalCell = this.Row[i].Cell[j];
				if (!externalCell) {
					continue;
				}
				var range = sheet.getRange2(externalCell.Ref);
				range._foreach(function (cell) {
					var api_sheet = Asc['editor'];
					var wb = api_sheet.wbModel;
					wb.dependencyFormulas.addToChangedCell(cell);
				});
			}
		}
	};

	ExternalSheetDataSet.prototype.getRow = function(index, needGenerateRow) {
		var row = null;

		for (var i = 0; i < this.Row.length; i++) {
			if (index === this.Row[i].R) {
				return this.Row[i];
			}
		}

		if (needGenerateRow) {
			row = new ExternalRow();
			row.R = index;
			this.Row.push(row);
		}

		return row;
	};

	ExternalSheetDataSet.prototype.addFrom = function(oSheetDataSet) {
		var row = null;

		for (var i = 0; i < oSheetDataSet.Row.length; i++) {
			if (!this.Row[i]) {
				this.Row[i] = oSheetDataSet.Row[i].clone();
			}
		}

		return row;
	};



	function ExternalRow() {
		this.R = null;
		this.Cell = [];
	}

	ExternalRow.prototype.Read_FromBinary2 = function(r) {
		if (r.GetBool()) {
			this.R = r.GetLong();
		}

		var length = r.GetLong();
		for (var i = 0; i < length; ++i) {
			var row = new ExternalCell();
			row.Read_FromBinary2(r);
			if(!this.Cell) {
				this.Cell = [];
			}
			this.Cell.push(row);
		}
	};
	ExternalRow.prototype.Write_ToBinary2 = function(w) {
		if (null != this.R) {
			w.WriteBool(true);
			w.WriteLong(this.R);
		} else {
			w.WriteBool(false);
		}

		w.WriteLong(this.Cell ? this.Cell.length : 0);
		if (this.Cell) {
			for (var i = 0; i < this.Cell.length; ++i) {
				this.Cell[i].Write_ToBinary2(w);
			}
		}
	};

	ExternalRow.prototype.clone = function () {
		var newObj = new ExternalRow();

		newObj.R = this.R;

		if (this.Cell) {
			for (var i = 0; i < this.Cell.length; ++i) {
				newObj.Cell.push(this.Cell[i].clone());
			}
		}

		return newObj;
	};

	ExternalRow.prototype.getCell = function(index, needGenerateRow) {
		var t = this;
		var cell = null;

		for (var i = 0; i < this.Cell.length; i++) {
			var cellRange = this.Cell[i].getRange();
			if (cellRange && index === cellRange.c1) {
				return this.Cell[i];
			}
		}

		if (needGenerateRow) {
			cell = new ExternalCell();
			AscCommonExcel.executeInR1C1Mode(false, function () {
				cell.Ref = new Asc.Range(index, t.R - 1, index, t.R - 1).getName();
			});
			this.Cell.push(cell);
		}

		return cell;
	};


	function ExternalCell() {
		this.Ref = null;//храним в строке, в будущем перевести в row/col
		this.CellType = null;
		this.CellValue = null;
	}

	ExternalCell.prototype.Read_FromBinary2 = function(r) {
		if (r.GetBool()) {
			this.Ref = r.GetString2();
		}
		if (r.GetBool()) {
			this.CellType = r.GetLong();
		}
		if (r.GetBool()) {
			this.CellValue = r.GetString2();
		}
	};
	ExternalCell.prototype.Write_ToBinary2 = function(w) {
		if (null != this.Ref) {
			w.WriteBool(true);
			w.WriteString2(this.Ref);
		} else {
			w.WriteBool(false);
		}

		if (null != this.CellType) {
			w.WriteBool(true);
			w.WriteLong(this.CellType);
		} else {
			w.WriteBool(false);
		}

		if (null != this.CellValue) {
			w.WriteBool(true);
			w.WriteString2(this.CellValue);
		} else {
			w.WriteBool(false);
		}
	};
	ExternalCell.prototype.clone = function () {
		var newObj = new ExternalCell();

		newObj.Ref = this.Ref;
		newObj.CellType = this.CellType;
		newObj.CellValue = this.CellValue;

		return newObj;
	};
	ExternalCell.prototype.initFromCell = function(cell, bUpdate, noData) {
		var isChanged = false;
		if (cell) {
			var t = this;
			if (!bUpdate) {
				AscCommonExcel.executeInR1C1Mode(false, function () {
					t.Ref = cell.getName();
				});
			}

			let newVal = noData ? "#REF" : cell.getValue();
			if (this.CellValue !== newVal) {
				isChanged = true;
				this.CellValue = newVal;
			}


			var cellValueType = null;
			switch (cell.getType()) {
				case CellValueType.String:
					cellValueType = Asc.ECellTypeType.celltypeStr;
					break;
				case CellValueType.Bool:
					cellValueType = Asc.ECellTypeType.celltypeBool;
					break;
				case CellValueType.Error:
					cellValueType = Asc.ECellTypeType.celltypeError;
					break;
			}

			if (noData) {
				cellValueType = Asc.ECellTypeType.celltypeError;
			}

			if (this.CellType !== cellValueType) {
				this.CellType = cellValueType;
				isChanged = true;
			}
		}
		return isChanged;
	};

	ExternalCell.prototype.getRange = function() {
		var res;
		var t = this;
		AscCommonExcel.executeInR1C1Mode(false, function () {
			res = AscCommonExcel.g_oRangeCache.getAscRange(t.Ref);
		});
		return res;
	};

	ExternalCell.prototype.getFormulaValue = function () {
		let res;

		switch (this.CellType) {
			case Asc.ECellTypeType.celltypeStr:
			case Asc.ECellTypeType.celltypeSharedString:
			case Asc.ECellTypeType.celltypeInlineStr:
			case Asc.ECellTypeType.celltypeDate:
				res = new AscCommonExcel.cString(this.CellValue);
				break;
			case Asc.ECellTypeType.celltypeBool:
				res = new AscCommonExcel.cBool(this.CellValue);
				break;
			case Asc.ECellTypeType.celltypeError:
				res = new AscCommonExcel.cError(this.CellValue);
				break;
			case Asc.ECellTypeType.celltypeNumber:
			default:
				res = new AscCommonExcel.cNumber(this.CellValue - 0);
				break;
		}

		return res;
	};

	function ExternalDefinedName(parent) {
		this.Name = null;
		this.RefersTo = null;
		this.SheetId = null;

		this.parent = parent;
	}

	ExternalDefinedName.prototype.Read_FromBinary2 = function(r) {
		if (r.GetBool()) {
			this.Name = r.GetString2();
		}
		if (r.GetBool()) {
			this.RefersTo = r.GetString2();
		}
		if (r.GetBool()) {
			this.SheetId = r.GetString2();
		}
	};
	ExternalDefinedName.prototype.Write_ToBinary2 = function(w) {
		if (null != this.Name) {
			w.WriteBool(true);
			w.WriteString2(this.Name);
		} else {
			w.WriteBool(false);
		}

		if (null != this.RefersTo) {
			w.WriteBool(true);
			w.WriteString2(this.RefersTo);
		} else {
			w.WriteBool(false);
		}

		if (null != this.SheetId) {
			w.WriteBool(true);
			w.WriteString2(this.SheetId);
		} else {
			w.WriteBool(false);
		}
	};
	ExternalDefinedName.prototype.clone = function () {
		var newObj = new ExternalDefinedName(this);

		newObj.Name = this.Name;
		newObj.RefersTo = this.RefersTo;
		newObj.SheetId = this.SheetId;

		newObj.parent = this.parent;

		return newObj;
	};
	ExternalDefinedName.prototype.updateFromSheet = function(sheet) {
		let isChanged = false;
		if (sheet) {
			//check on sheet name and def name
			let defNames = sheet.workbook && sheet.workbook.dependencyFormulas && sheet.workbook.dependencyFormulas.defNames;
			let sheetName = sheet.getName();
			if (defNames) {
				if (defNames.sheet[sheetName]) {
					if (defNames.sheet[sheetName][this.Name]) {
						isChanged = true;
					}
				}

				if (!isChanged) {
					if (defNames.wb[this.Name]) {
						this.RefersTo = defNames.wb[this.Name].getRef();
						//need init from range + updateFromSheet from data set
						if (this.RefersTo) {
							let exclamationMarkIndex = this.RefersTo.lastIndexOf("!");
							this.parent.updateSheetData(sheetName, sheet, [AscCommonExcel.g_oRangeCache.getAscRange(this.RefersTo.slice(exclamationMarkIndex + 1))]);
						}

						isChanged = true;
					}
				}
			}
		}
		return isChanged;
	};

	//CellWatch
	function CCellWatch(ws) {
		this.r = null;

		this._ws = ws;
		this._workbook = null;
		this._sheet = null;
		this._name = null;
		this._cell = null;
		this._value = null;
		this._formula = null;

		this.needRecalc = null;

		return this;
	}
	CCellWatch.prototype.clone = function () {
		var res = new CCellWatch();
		res.r = this.r.clone();

		res._ws = this._ws;
		res._workbook = this._workbook;
		res._sheet = this._sheet;
		res._name = this._name;
		res._cell = this._cell;
		res._value = this._value;
		res._formula = this._formula;

		return res;
	};
	CCellWatch.prototype.setNeedRecalc = function () {
		this.needRecalc = true;
	};
	CCellWatch.prototype.setRef = function (ref) {
		this.r = ref;
	};
	CCellWatch.prototype.asc_getWorkbook = function () {
		return this._workbook;
	};
	CCellWatch.prototype.asc_getSheet = function () {
		return this._sheet;
	};
	CCellWatch.prototype.asc_getName = function () {
		return this._name;
	};
	CCellWatch.prototype.asc_getCell = function () {
		return this._cell;
	};
	CCellWatch.prototype.asc_getValue = function () {
		return this._value;
	};
	CCellWatch.prototype.asc_getFormula = function () {
		return this._formula;
	};
	CCellWatch.prototype.recalculate = function (ignoreNeedRecalc) {
		var isChanged = false;
		if (this.needRecalc || ignoreNeedRecalc) {
			if (this._ws) {
				var docInfo = window["Asc"]["editor"].DocInfo;

				var dN = new Asc.Range(this.r.c1, this.r.r1, this.r.c2, this.r.r2, true);
				var defName = parserHelp.get3DRef(this._ws.getName(), dN.getAbsName());
				defName = this._ws.workbook.findDefinesNames(defName, this._ws.getId(), true);

				var formula = "";
				var value = "";
				if (this.r && this._ws) {
					this._ws.getCell3(this.r.r1, this.r.c1)._foreachNoEmpty(function (cell) {
						if (cell.isFormula()) {
							formula = cell.getValueForEdit();
						}
						value = cell.getValue();
					});
				}

				if (this._name != defName) {
					this._name = defName ? defName : null;
					isChanged = true;
				}

				this._cell = this.r ? this.r.getName() : this.r;

				if (!this._workbook) {
					this._workbook = docInfo ? docInfo.get_Title() : "";
					isChanged = true;
				}
				if (this._sheet !== this._ws.sName) {
					this._sheet = this._ws.sName;
					isChanged = true;
				}
				if (this._value !== value) {
					this._value = value;
					isChanged = true;
				}
				if (this._formula !== formula) {
					this._formula = formula;
					isChanged = true;
				}

				this.needRecalc = null;
			}
		}
		return isChanged;
	};
	CCellWatch.prototype.initPostOpen = function (ws) {
		this._ws = ws;
	};
	CCellWatch.prototype.setOffset = function (row, col) {
		if (this.r) {
			this.r.setOffset({row: row ? row : 0, col: col ? col : 0});
		}
	};

	function CRowColBreaks(/*ws*/) {
		this.count = null;
		this.manualBreakCount = null;
		this.breaks = [];

		return this;
	}
	CRowColBreaks.prototype.clone = function () {
		var res = new CRowColBreaks();
		res.count = this.count;
		res.manualBreakCount = this.count;

		for (let i = 0 ; i < this.breaks.length; i++) {
			res.breaks.push(this.breaks[i].clone());
		}

		return res;
	};
	CRowColBreaks.prototype.setCount = function (val) {
		this.count = val;
	};
	CRowColBreaks.prototype.setManualBreakCount = function (val) {
		this.manualBreakCount = val;
	};
	CRowColBreaks.prototype.pushBreak = function (val) {
		this.breaks.push(val);
	};
	CRowColBreaks.prototype.getCount = function () {
		return this.count;
	};
	CRowColBreaks.prototype.getManualBreakCount = function () {
		return this.manualBreakCount;
	};
	CRowColBreaks.prototype.getBreaks = function () {
		return this.breaks;
	};
	CRowColBreaks.prototype.isBreak = function (index, opt_min, opt_max) {
		if (this.count > 0) {
			for (let i = 0; i < this.breaks.length; i++) {
				if (this.breaks[i].isBreak(index, opt_min, opt_max)) {
					return true;
				}
			}
		}
		return false;
	};
	CRowColBreaks.prototype.forEach = function (callback) {
		for (var i = 0, l = this.breaks.length; i < l; ++i) {
			if (callback(this.breaks[i], i)) {
				break;
			}
		}
	};
	CRowColBreaks.prototype.containsBreak = function (id, min, max) {
		let res = false;
		this.forEach(function (breaksElem) {
			if (breaksElem.id === id) {
				res = true;
				return true;
			}
		});
		return res;
	};
	CRowColBreaks.prototype.getBreak = function (id) {
		let res = null;
		this.forEach(function (breaksElem) {
			if (breaksElem.id === id) {
				res = breaksElem;
				return true;
			}
		});
		return res;
	};

	CRowColBreaks.prototype.changeBreak = function (idFrom, idTo, min, max, man, pt) {
		let breakElem = this.getBreak(idFrom);
		if (breakElem) {
			let res = breakElem.set(idTo, min, max, man, pt);
			this.reSortingBreaks();
			return res;
		}
	};

	CRowColBreaks.prototype.addBreak = function (id, min, max, man, pt) {
		if (!id) {
			return false;
		}
		let newBreak = new CBreak();
		newBreak.set(id, min, max, man, pt);

		let res = this._addBreak(newBreak);
		this.reSortingBreaks();
		return res;
	};

	CRowColBreaks.prototype._addBreak = function (newBreak) {
		this.breaks.push(newBreak);
		if (this.manualBreakCount === null) {
			this.manualBreakCount = 0;
		}
		if (this.count === null) {
			this.count = 0;
		}
		this.manualBreakCount++;
		this.count++;
		return true;
	};

	CRowColBreaks.prototype.removeBreak = function (id, min, max) {
		let isDeleted = false;
		for (let i = 0; i < this.breaks.length; i++) {
			if (this.breaks[i].isBreak(id, min, max)) {
				this.breaks.splice(i, 1);
				isDeleted = true;
				this.manualBreakCount--;
				this.count--;
				break;
			}
		}
		this.reSortingBreaks();
		return isDeleted;
	};

	CRowColBreaks.prototype.reSortingBreaks = function () {
		this.breaks.sort(function sortArr(a, b) {
			return a.id - b.id;
		});
	};


	function CBreak(/*ws*/) {
		this.id = null;
		this.min = null;
		this.max = null;

		this.man = null;
		this.pt = null;

		return this;
	}
	CBreak.prototype.clone = function () {
		var res = new CBreak();
		res.id = this.id;

		res.min = this.min;
		res.max = this.max;

		res.man = this.man;
		res.pt = this.pt;

		return res;
	};
	CBreak.prototype.setId = function (val) {
		this.id = val;
	};
	CBreak.prototype.setMan = function (val) {
		this.man = val;
	};
	CBreak.prototype.setMax = function (val) {
		this.max = val;
	};
	CBreak.prototype.setMin = function (val) {
		this.min = val;
	};
	CBreak.prototype.setPt = function (val) {
		this.pt = val;
	};
	CBreak.prototype.getId = function () {
		return this.id;
	};
	CBreak.prototype.getMan = function () {
		return this.man;
	};
	CBreak.prototype.getMax = function () {
		return this.max;
	};
	CBreak.prototype.getMin = function () {
		return this.min;
	};
	CBreak.prototype.getPt = function () {
		return this.pt;
	};
	CBreak.prototype.isBreak = function (index, opt_min, opt_max) {
		let res = this.id === index;
		if (res) {
			if (opt_min != null && opt_max != null) {
				if (this.min == null && this.max == null) {
					res = false;
				} else {
					if (this.min > opt_max || this.max < opt_min) {
						res = false;
					}
				}
			}
		}
		return res;
	};
	CBreak.prototype.set = function (id, min, max, man, pt) {
		let isChanged = false;
		if (id !== this.getId()) {
			this.setId(id);
			isChanged = true;
		}
		if (min !== this.getMin()) {
			this.setMin(min);
			isChanged = true;
		}
		if (max !== this.getMax()) {
			this.setMax(max);
			isChanged = true;
		}
		if (man !== this.getMan()) {
			this.setMan(man);
			isChanged = true;
		}
		if (pt !== this.getPt()) {
			this.setPt(pt);
			isChanged = true;
		}
		return isChanged ? this : null;
	};


	/**
	 * Class representing a Series settings for fills data of context menu and dialog window - "Series"
	 * @property {c_oAscSeriesInType} seriesIn - Series in. Contains: Rows, Columns
	 * @property {c_oAscSeriesType} type - Type. Contains:  Linear, Growth, Date, AutoFill
	 * @property {c_oAscDateUnitType} dateUnit - Date Unit. Contains: Day, Weekday, Month, Year
	 * @property {number|null} stepValue - Step value
	 * @property {number|null} stopValue - Stop value
	 * @property {boolean|null} trend - Trend
	 * @property {object} contextMenuAllowedProps - Allowed properties
	 * @property {c_oAscFillType} contextMenuChosenProperty - Chosen property of context menu
	 * @constructor
	 * @returns {asc_CSeriesSettings}
	 */
	function asc_CSeriesSettings() {
		this.seriesIn = null;
		this.type = null;
		this.dateUnit = null;

		this.trend = null;

		this.stepValue = null;
		this.stopValue = null;

		//transfer right click fill information
		this.contextMenuAllowedProps = null;
		this.contextMenuChosenProperty = null;
		this.toolbarMenuAllowedProps = null;

		return this;
	}

	/**
	 * Method clones SeriesSettings object
	 * @memberof asc_CSeriesSettings
	 * @returns {asc_CSeriesSettings}
	 */
	asc_CSeriesSettings.prototype.clone = function () {
		let res = new asc_CSeriesSettings();

		res.seriesIn = this.seriesIn;
		res.type = this.type;
		res.dateUnit = this.dateUnit;

		res.trend = this.trend;

		res.stepValue = this.stepValue;
		res.stopValue = this.stopValue;

		res.contextMenuAllowedProps = this.contextMenuAllowedProps;
		res.contextMenuChosenProperty = this.contextMenuChosenProperty;
		res.toolbarMenuAllowedProps = this.toolbarMenuAllowedProps;

		return res;
	};

	/**
	 * Function returns next value from range
	 * @param {Range} oRange - Range object from Workbook model
	 * @param {Cell} oCell
	 * @param {boolean} bVertical
	 * @returns {string} Value without format of next cell
	 * @private
	 */
	function _getNextValueFromRange(oRange, oCell, bVertical) {
		let ws = oRange.worksheet;

		if (bVertical) {
			let nNextItRow = oCell.nRow + 1;
			if (nNextItRow <= oRange.bbox.r2) {
				return ws.getCell3(nNextItRow, oCell.nCol).getValueWithoutFormat();
			}
		} else {
			let nNextItCol = oCell.nCol + 1;
			if (nNextItCol <= oRange.bbox.c2) {
				return ws.getCell3(oCell.nRow, nNextItCol).getValueWithoutFormat();
			}
		}

		return '';
	}
	/**
	 * Function returns range with filled cells.
	 * @param {Range} oRange - Range object from Workbook model
	 * @param {boolean} bVertical - Vertical or horizontal direction
	 * @returns  Range object from Workbook model
	 * @private
	 */
	function _getFilledRange(oRange, bVertical) {
		let oFilledRange = oRange.clone();
		let nRow = oFilledRange.bbox.r1;
		let nCol = oFilledRange.bbox.c1;
		let sFirstValue = null;

		oRange._foreach(function (oCell, nCurRow, nCurCol) {
			if (oCell && oCell.getValueWithoutFormat()) {
				nRow = nCurRow;
				nCol = nCurCol;
				if (nCurRow === oRange.bbox.r1 && nCurCol === oRange.bbox.c1) {
					sFirstValue = oCell.getValueWithoutFormat();
				}
			} else {
				let sNextValue = _getNextValueFromRange(oRange, oCell, bVertical);
				if (!sFirstValue && sNextValue) {
					if (bVertical) {
						oFilledRange.bbox.r1 = nCurRow + 1;
						oFilledRange.bbox.c1 = nCurCol;
					} else {
						oFilledRange.bbox.c1 = nCurCol + 1;
						oFilledRange.bbox.r1 = nCurRow;
					}
				}
			}
		});
		oFilledRange.bbox.r2 = nRow;
		oFilledRange.bbox.c2 = nCol;

		return oFilledRange;
	}
	/**
	 * Rounds numbers with 0 after a point. E.g. Numbers like 0.500000001
	 * @param {number} nValue Number to need round
	 * @returns {number} Rounded number
	 */
	function _roundResult(nValue) {
		if (!Number.isInteger(nValue)) {
			let sValue = nValue.toString();
			let secondPartOfNumber = sValue.split('.')[1];

			if (secondPartOfNumber.includes('00000') && !sValue.includes('.00000')) {
				let nRoundIndex = secondPartOfNumber.indexOf('00000');
				return parseFloat(nValue.toFixed(nRoundIndex));
			}
		}

		// Return value without changes
		return nValue;
	}
	/**
	 * Method fills data of SeriesSettings object for context menu and dialog window
	 * @memberof asc_CSeriesSettings
	 * @param {WorksheetView} ws
	 */
	asc_CSeriesSettings.prototype.prepare = function (ws) {
		if (!ws) {
			return;
		}

		function calcAvg(seriesInType) {
			let filledRangeLength = 0;

			if (seriesInType === Asc.c_oAscSeriesInType.rows) {
				filledRange._foreach2(function (cell, curRow, curCol, rowStart, colStart) {
					if ((cell && cell.getNumberValue() != null) && curRow === rowStart) {
						// If columns range not starts with 0, define indexCell by difference between current column and start column
						let indexCell = curCol - colStart;
						ySum += cell.getNumberValue();
						xSum += indexCell;
						filledRangeLength++;
					}
				});
			} else {
				filledRange._foreach2(function (cell, curRow, curCol, rowStart, colStart) {
					if ((cell && cell.getNumberValue() != null) && curCol === colStart) {
						let indexCell = curRow - rowStart;
						ySum += cell.getNumberValue();
						xSum += indexCell;
						filledRangeLength++;
					}
				});
			}

			xAvg = xSum / filledRangeLength;
			yAvg = ySum / filledRangeLength;
		}
		function actionCell(cell, curRow, curCol, rowStart, colStart) {
			if (cell && cell.getValueWithoutFormat()) {
				// Fill type
				if (seriesSettings.asc_getType() == null) {
					seriesSettings.asc_setType(Asc.c_oAscSeriesType.linear);
					if (cell.xfs != null && cell.xfs.num != null && cell.xfs.num.getFormat() != null) {
						let numFormat = AscCommon.oNumFormatCache.get(cell.xfs.num.getFormat());
						if (numFormat.isDateTimeFormat() && numFormat.getType() === Asc.c_oAscNumFormatType.Date) {
							seriesSettings.asc_setType(Asc.c_oAscSeriesType.date);

							contextMenuAllowedProps[Asc.c_oAscFillType.fillDays] = true;
							contextMenuAllowedProps[Asc.c_oAscFillType.fillWeekdays] = true;
							contextMenuAllowedProps[Asc.c_oAscFillType.fillMonths] = true;
							contextMenuAllowedProps[Asc.c_oAscFillType.fillYears] = true;
						}
					}
				}

				// Fill step value
				let cellNumberValue = cell.getNumberValue();
				let isVertical = seriesSettings.asc_getSeriesIn() === Asc.c_oAscSeriesInType.columns;
				let cellIndex = isVertical ? curRow - rowStart : curCol - colStart;

				if (cell.isFormula()) {
					seriesSettings.asc_setStepValue(1);
				} else if (firstValue != null) {
					if (cellNumberValue != null) {
						numeratorOfSlope += (cellIndex - xAvg) * (cellNumberValue - yAvg);
						denominatorOfSlope += Math.pow((cellIndex - xAvg), 2);
						if (seriesSettings.asc_getType() === Asc.c_oAscSeriesType.date) {
							let firstValueDate = new Asc.cDate().getDateFromExcel(firstValue < 60 ? firstValue + 1 : firstValue);
							let curValueDate = new Asc.cDate().getDateFromExcel(cellNumberValue < 60 ? cellNumberValue + 1 : cellNumberValue);
							let nextValue = _getNextValueFromRange(filledRange, cell, isVertical) - 0;
							let nextValueDate = new Asc.cDate().getDateFromExcel(nextValue < 60 ? nextValue + 1 : nextValue);
							let isSequenceSeries = isVertical ? curRow === (rowStart + 1) : curCol === (colStart + 1);
							// For recognize month
							let daysIsEqual = firstValueDate.getDate() === curValueDate.getDate() && firstValueDate.getDate() === nextValueDate.getDate();
							let monthsIsNotEqual = firstValue!== cellNumberValue && firstValue!== nextValue;
							// For recognize year
							let monthIsNotSequence = firstValueDate.getMonth() === curValueDate.getMonth();
							let yearsIsNotEqual = firstValueDate.getFullYear() !== curValueDate.getFullYear() && firstValueDate.getFullYear() !== nextValueDate.getFullYear();
							if (daysIsEqual && monthsIsNotEqual && !monthIsNotSequence) {
								seriesSettings.asc_setDateUnit(Asc.c_oAscDateUnitType.month);
								seriesSettings.asc_setStepValue(Math.round((cellNumberValue - firstValue) / 30));
							} else if (daysIsEqual && yearsIsNotEqual) {
								seriesSettings.asc_setDateUnit(Asc.c_oAscDateUnitType.year);
								seriesSettings.asc_setStepValue(curValueDate.getFullYear() - firstValueDate.getFullYear());
							}  else if(isSequenceSeries) {
								seriesSettings.asc_setStepValue(cellNumberValue - firstValue);
							}
						}
					} else {
						seriesSettings.asc_setStepValue(1);

						contextMenuAllowedProps[Asc.c_oAscFillType.fillSeries] = true;
						contextMenuAllowedProps[Asc.c_oAscFillType.series] = true;
					}
				} else {
					if (cellNumberValue != null) {
						numeratorOfSlope += (cellIndex - xAvg) * (cellNumberValue - yAvg);
						denominatorOfSlope += Math.pow((cellIndex - xAvg), 2);
						firstValue = cellNumberValue;
					} else {
						seriesSettings.asc_setStepValue(1);

						let typeCell = cell.getType();
						if (typeCell === AscCommon.CellValueType.String) {
							let cellValue = cell.getValueWithoutFormat();
							if (cellValue[cellValue.length - 1] >= '0' && cellValue[cellValue.length - 1] <= '9') {
								contextMenuAllowedProps[Asc.c_oAscFillType.fillSeries] = true;
							}
						}
					}
				}
			} else if (firstValue == null){
				seriesSettings.asc_setStepValue(1);
			}

			if (seriesSettings.asc_getStepValue() != null) {
				if (seriesSettings.asc_getType() == null) {
					seriesSettings.asc_setType(Asc.c_oAscSeriesType.linear);
				}
				if (seriesSettings.asc_getType() === Asc.c_oAscSeriesType.date) {
					contextMenuAllowedProps[Asc.c_oAscFillType.fillSeries] = true;
					contextMenuAllowedProps[Asc.c_oAscFillType.series] = true;
				}
				return true;
			}
		}

		let seriesSettings = this;
		let selection = ws.model && ws.model.getSelection();
		let selectionRanges = selection && selection.ranges;
		let range = selectionRanges && selectionRanges[0];
		if(!range) {
			return;
		}
		let rangeModel = ws.model.getRange3(range.r1, range.c1, range.r2, range.c2);
		let firstValue = null;
		let contextMenuAllowedProps = {};
		contextMenuAllowedProps[Asc.c_oAscFillType.copyCells] = true;
		contextMenuAllowedProps[Asc.c_oAscFillType.fillSeries] = false;
		contextMenuAllowedProps[Asc.c_oAscFillType.fillFormattingOnly] = null;
		contextMenuAllowedProps[Asc.c_oAscFillType.fillWithoutFormatting] = null;
		contextMenuAllowedProps[Asc.c_oAscFillType.fillDays] = false;
		contextMenuAllowedProps[Asc.c_oAscFillType.fillWeekdays] = false;
		contextMenuAllowedProps[Asc.c_oAscFillType.fillMonths] = false;
		contextMenuAllowedProps[Asc.c_oAscFillType.fillYears] = false;
		contextMenuAllowedProps[Asc.c_oAscFillType.linearTrend] = false;
		contextMenuAllowedProps[Asc.c_oAscFillType.growthTrend] = false;
		contextMenuAllowedProps[Asc.c_oAscFillType.flashFill] = null;
		contextMenuAllowedProps[Asc.c_oAscFillType.series] = false;

		let countOfCol = range.c2 - range.c1;
		let countOfRow = range.r2 - range.r1;
		let filledRange = _getFilledRange(rangeModel, countOfCol <= countOfRow);
		let countOfFilledRow = filledRange.bbox.r2 - filledRange.bbox.r1;
		let countOfFilledCol = filledRange.bbox.c2 - filledRange.bbox.c1;
		// Init variables for calc step. x - index, y - value of cell.
		let xSum = 0;
		let ySum = 0;
		let xAvg = 0;
		let yAvg = 0;
		let numeratorOfSlope = 0;
		let denominatorOfSlope = 0;

		let isApplyFilter = false;
		//if sheet contains applyed filter -> only copy cell
		if (ws.model && ws.model.AutoFilter && ws.model.AutoFilter.isApplyAutoFilter()) {
			isApplyFilter = true;
		} else if (ws.model) {
			for (let i = 0; i < selectionRanges.length; i++) {
				let tables =  ws.model.autoFilters.getTablesIntersectionRange(selectionRanges[i]);
				for (let j = 0; j < tables.length; j++) {
					if (tables[i].isApplyAutoFilter()) {
						isApplyFilter = true;
						break;
					}
				}
				if (isApplyFilter) {
					break;
				}
			}
		}
		this.asc_setDateUnit(Asc.c_oAscDateUnitType.day);
		this.asc_setTrend(false);
		// select one cell and use fill handle
		if (!isApplyFilter) {
			if (ws.activeFillHandle != null) {
				if (ws.fillHandleDirection === 0) {
					this.asc_setSeriesIn(Asc.c_oAscSeriesInType.rows);
				} else {
					this.asc_setSeriesIn(Asc.c_oAscSeriesInType.columns);
				}
				// Calculate xAvg, yAvg
				calcAvg(this.asc_getSeriesIn());
				filledRange._foreach2(actionCell);
				if (rangeModel.getType() === AscCommon.CellValueType.Number) {
					contextMenuAllowedProps[Asc.c_oAscFillType.fillSeries] = true;
					contextMenuAllowedProps[Asc.c_oAscFillType.series] = true;
				}
			} else if (countOfCol >= countOfRow) {
				this.asc_setSeriesIn(Asc.c_oAscSeriesInType.rows);
				calcAvg(Asc.c_oAscSeriesInType.rows);
				filledRange._foreach2(function (cell, curRow, curCol, rowStart, colStart) {
					if (curRow === rowStart) {
						return actionCell(cell, curRow, curCol, rowStart, colStart);
					}
				});
			} else {
				this.asc_setSeriesIn(Asc.c_oAscSeriesInType.columns);
				calcAvg(Asc.c_oAscSeriesInType.columns);
				filledRange._foreach2(function (cell, curRow, curCol, rowStart, colStart) {
					if (curCol === colStart) {
						return actionCell(cell, curRow, curCol, rowStart, colStart);
					}
				});
			}
		}

		if (!isApplyFilter && seriesSettings.asc_getStepValue() == null) {
			let isVertical = this.asc_getSeriesIn() === Asc.c_oAscSeriesInType.columns;
			let rangeLen = isVertical ? countOfRow + 1 : countOfCol + 1;
			let filledRangeLen = isVertical ? countOfFilledRow + 1 : countOfFilledCol + 1;
			let isStartPointShifted = rangeModel.bbox.r1 !== filledRange.bbox.r1 || rangeModel.bbox.c1 !== filledRange.bbox.c1;

			// all selected cells are filled or only one cell is filled or first cell is empty
			if (filledRangeLen === rangeLen || firstValue === ySum || firstValue == null || isStartPointShifted) {
				this.asc_setStepValue(1);
			} else {
				let slope = _roundResult(numeratorOfSlope / denominatorOfSlope);
				if (this.asc_getType() === Asc.c_oAscSeriesType.date && !Number.isInteger(slope)) {
					this.asc_setStepValue(1);
				} else {
					this.asc_setStepValue(slope);
				}
			}

			if (firstValue != null) {
				contextMenuAllowedProps[Asc.c_oAscFillType.fillSeries] = true;
				contextMenuAllowedProps[Asc.c_oAscFillType.series] = true;
				if (firstValue !== ySum && seriesSettings.asc_getType() !== Asc.c_oAscSeriesType.date) {
					contextMenuAllowedProps[Asc.c_oAscFillType.linearTrend] = true;
					contextMenuAllowedProps[Asc.c_oAscFillType.growthTrend] = true;
				}
			}
		}

		//2. init for context menu - allowed options
		this.asc_setContextMenuAllowedProps(contextMenuAllowedProps);

		//3. toolbar - allowed options
		let toolbarMenuAllowedProps = {};
		toolbarMenuAllowedProps[Asc.c_oAscFillType.fillDown] = true;
		toolbarMenuAllowedProps[Asc.c_oAscFillType.fillRight] = true;
		toolbarMenuAllowedProps[Asc.c_oAscFillType.fillUp] = true;
		toolbarMenuAllowedProps[Asc.c_oAscFillType.fillLeft] = true;
		toolbarMenuAllowedProps[Asc.c_oAscFillType.series] = !isApplyFilter;

		if (range.isOneCol()) {
			if (range.c1 === 0) {
				toolbarMenuAllowedProps[Asc.c_oAscFillType.fillRight] = false;
			} else if (range.c1 === AscCommon.gc_nMaxCol0) {
				toolbarMenuAllowedProps[Asc.c_oAscFillType.fillLeft] = false;
			}
		}
		if (range.isOneRow()) {
			if (range.r1 === 0) {
				toolbarMenuAllowedProps[Asc.c_oAscFillType.fillDown] = false;
			} else if (range.r1 === AscCommon.gc_nMaxRow0) {
				toolbarMenuAllowedProps[Asc.c_oAscFillType.fillUp] = false;
			}
		}

		this.asc_setToolbarMenuAllowedProps(toolbarMenuAllowedProps);
	};
	/**
	 * Method updates "Type", "Trend" and "Step Value" (for Date type) attributes of SeriesSettings object for chosen context menu property
	 * @param {WorksheetView} wsView
	 * @memberof asc_CSeriesSettings
	 */
	asc_CSeriesSettings.prototype.init = function (wsView) {
		const chosenContextMenuProp = this.asc_getContextMenuChosenProperty();
		const fillType = Asc.c_oAscFillType;
		const seriesSettings = this;
		function calcStepForDate() {
			const arn = wsView.model && wsView.model.getSelection().getLast();
			const activeFillHandle = wsView.activeFillHandle;
			const rangeModel = wsView.model.getRange3(arn.r1, arn.c1, arn.r2, arn.c2);
			const index = wsView.fillHandleDirection === 0 ? activeFillHandle.c2 - arn.c1 : activeFillHandle.r2 - arn.r1;
			const dateUnitType = Asc.c_oAscDateUnitType;

			let prevCellValue = null;

			function actionCell(cell) {
				if (cell && cell.getValueWithoutFormat()) {
					let dateUnit = seriesSettings.asc_getDateUnit();
					let cellValue;

					if (dateUnit === dateUnitType.day || dateUnit === dateUnitType.weekday) {
						cellValue = cell.getNumberValue();
					} else if (dateUnit === dateUnitType.month) {
						let dateVal = new Asc.cDate().getDateFromExcel(cell.getNumberValue());
						cellValue = dateVal.getMonth();
					} else {
						let dateVal = new Asc.cDate().getDateFromExcel(cell.getNumberValue());
						cellValue = dateVal.getFullYear();
					}

					if (prevCellValue != null) {
						seriesSettings.asc_setStepValue(cellValue - prevCellValue);
						return true;
					} else {
						prevCellValue = cellValue;
					}
				} else {
					return true;
				}
			}

			if (arn.isOneCell()) {
				seriesSettings.asc_setStepValue(1);
			} else {
				if (seriesSettings.asc_getSeriesIn() === Asc.c_oAscSeriesInType.rows) {
					rangeModel._foreach2(actionCell);
				} else {
					rangeModel._foreach2(function (cell, curRow, curCol, rowStart, colStart) {
						if (curCol === colStart) {
							return actionCell(cell);
						}
					});
				}
			}

			// Define direction step
			let step = seriesSettings.asc_getStepValue();
			if (index < 0) {
				seriesSettings.asc_setStepValue(step * -1);
			}
		}

		switch (chosenContextMenuProp) {
			case fillType.linearTrend:
				this.asc_setType(Asc.c_oAscSeriesType.linear);
				this.asc_setTrend(true);
				break;
			case fillType.growthTrend:
				this.asc_setType(Asc.c_oAscSeriesType.growth);
				this.asc_setTrend(true);
				break;
			case fillType.fillDays:
				this.asc_setType(Asc.c_oAscSeriesType.date);
				this.asc_setDateUnit(Asc.c_oAscDateUnitType.day);
				calcStepForDate()
				break;
			case fillType.fillWeekdays:
				this.asc_setType(Asc.c_oAscSeriesType.date);
				this.asc_setDateUnit(Asc.c_oAscDateUnitType.weekday);
				calcStepForDate()
				break;
			case fillType.fillMonths:
				this.asc_setType(Asc.c_oAscSeriesType.date);
				this.asc_setDateUnit(Asc.c_oAscDateUnitType.month);
				calcStepForDate()
				break;
			case fillType.fillYears:
				this.asc_setType(Asc.c_oAscSeriesType.date);
				this.asc_setDateUnit(Asc.c_oAscDateUnitType.year);
				calcStepForDate()
				break;
		}
	};
	/**
	 * Method returns "Series In" attribute of SeriesSettings object
	 * @memberof asc_CSeriesSettings
	 * @returns {c_oAscSeriesInType}
	 */
	asc_CSeriesSettings.prototype.asc_getSeriesIn = function () {
		return this.seriesIn;
	};
	/**
	 * Method returns "Type" attribute of SeriesSettings object
	 * @memberof asc_CSeriesSettings
	 * @returns {c_oAscSeriesType}
	 */
	asc_CSeriesSettings.prototype.asc_getType = function () {
		return this.type;
	};
	/**
	 * Method returns "Date Unit" attribute of SeriesSettings object
	 * @memberof asc_CSeriesSettings
	 * @returns {c_oAscDateUnitType}
	 */
	asc_CSeriesSettings.prototype.asc_getDateUnit = function () {
		return this.dateUnit;
	};
	/**
	 * Method returns "Trend" attribute of SeriesSettings object
	 * @memberof asc_CSeriesSettings
	 * @returns {boolean}
	 */
	asc_CSeriesSettings.prototype.asc_getTrend = function () {
		return this.trend;
	};
	/**
	 * Method returns "Step Value" attribute of SeriesSettings object
	 * @memberof asc_CSeriesSettings
	 * @returns {number}
	 */
	asc_CSeriesSettings.prototype.asc_getStepValue = function () {
		return this.stepValue;
	};
	/**
	 * Method returns "Stop Value" attribute of SeriesSettings object
	 * @memberof asc_CSeriesSettings
	 * @returns {number}
	 */
	asc_CSeriesSettings.prototype.asc_getStopValue = function () {
		return this.stopValue;
	};
	/**
	 * Method returns "contextMenuAllowedProps" attribute of SeriesSettings object.
	 * Uses for hide and shade menu items in context menu.
	 * @memberof asc_CSeriesSettings
	 * @returns {object} - object with properties of Asc.c_oAscFillType
	 */
	asc_CSeriesSettings.prototype.asc_getContextMenuAllowedProps = function () {
		return this.contextMenuAllowedProps;
	};
	/**
	 * Method returns "contextMenuChosenProperty" attribute of SeriesSettings object.
	 * Uses for recognize chosen item in context menu.
	 * @memberof asc_CSeriesSettings
	 * @returns {c_oAscFillType}
	 */
	asc_CSeriesSettings.prototype.asc_getContextMenuChosenProperty = function () {
		return this.contextMenuChosenProperty;
	};
	/**
	 * Method returns "toolbarMenuAllowedProps" attribute of SeriesSettings object.
	 * Uses for hide and shade menu items in context menu.
	 * @memberof asc_CSeriesSettings
	 * @returns {object} - object with properties of Asc.c_oAscFillType
	 */
	asc_CSeriesSettings.prototype.asc_getToolbarMenuAllowedProps = function () {
		return this.toolbarMenuAllowedProps;
	};
	/**
	 * Method sets "Series In" attribute of SeriesSettings object
	 * @memberof asc_CSeriesSettings
	 * @param {c_oAscSeriesInType} val - attribute of c_oAscSeriesInType
	 */
	asc_CSeriesSettings.prototype.asc_setSeriesIn = function (val) {
		this.seriesIn = val;
	};
	/**
	 * Method sets "Type" attribute of SeriesSettings object
	 * @memberof asc_CSeriesSettings
	 * @param {c_oAscSeriesType} val - attribute of c_oAscSeriesType
	 */
	asc_CSeriesSettings.prototype.asc_setType = function (val) {
		this.type = val;
	};
	/**
	 * Method sets "Date Unit" attribute of SeriesSettings object
	 * @memberof asc_CSeriesSettings
	 * @param {c_oAscDateUnitType} val - attribute of c_oAscDateUnitType
	 */
	asc_CSeriesSettings.prototype.asc_setDateUnit = function (val) {
		this.dateUnit = val;
	};
	/**
	 * Method sets "Trend" attribute of SeriesSettings object
	 * @memberof asc_CSeriesSettings
	 * @param {boolean} val
	 */
	asc_CSeriesSettings.prototype.asc_setTrend = function (val) {
		this.trend = val;
	};
	/**
	 * Method sets "Step Value" attribute of SeriesSettings object
	 * @param {number} val
	 */
	asc_CSeriesSettings.prototype.asc_setStepValue = function (val) {
		this.stepValue = val;
	};
	/**
	 * Method sets "Stop Value" attribute of SeriesSettings object
	 * @param {number} val
	 */
	asc_CSeriesSettings.prototype.asc_setStopValue = function (val) {
		this.stopValue = val;
	};
	/**
	 * Method checks "Step Value" attribute of SeriesSettings object
	 * @param {number} val
	 */
	asc_CSeriesSettings.prototype.asc_isValidStepValue = function (val) {
		return this.checkValidValue(val);
	};

	/**
	 * Method checks "Stop Value" attribute of SeriesSettings object
	 * @param {number} val
	 */
	asc_CSeriesSettings.prototype.asc_isValidStopValue = function (val) {
		return this.checkValidValue(val);
	};

	/**
	 * Method checks input values
	 * @param {number} val
	 */
	asc_CSeriesSettings.prototype.checkValidValue = function (val) {
		let errCode = Asc.c_oAscError.ID.No;

		let regstr = new RegExp('^\s*[+-]?([0-9]+([.][0-9]*)?|[.][0-9]+)\s*$');
		if (typeof val === 'string') {
			let findComma = val.match(/,/g);
			if (findComma && findComma.length === 1) {
				val = val.replace(',','.');
			}
		}

		if (val !== '' && (!regstr.test(val.trim()) || isNaN(parseFloat(val)))) {
			//check on format
			let parsedRes = AscCommon.g_oFormatParser.parse(val);
			if (parsedRes === null) {
				errCode = Asc.c_oAscError.ID.MustIntegerOrDecimalNumber;
				val = null;
			} else {
				val = parsedRes.value;
			}
		}

		return [errCode, (val != null && val !== "") ? parseFloat(val) : null];
	};

	/**
	 * Method sets "contextMenuAllowedProps" attribute of SeriesSettings object.
	 * Uses for hide and shade menu items in context menu.
	 * * true - unshade menu item
	 * * false - shade menu item
	 * * null - hide menu item
	 * @memberof asc_CSeriesSettings
	 * @param {object} val - object with properties of Asc.c_oAscFillType as attribute and boolean or null as value
	 */
	asc_CSeriesSettings.prototype.asc_setContextMenuAllowedProps = function (val) {
		this.contextMenuAllowedProps = val;
	};
	/**
	 * Method sets "contextMenuChosenProperty" attribute of SeriesSettings object.
	 * Uses for recognize chosen item in context menu.
	 * @param {c_oAscFillType} val - attribute of c_oAscFillType
	 */
	asc_CSeriesSettings.prototype.asc_setContextMenuChosenProperty = function (val) {
		this.contextMenuChosenProperty = val;
	};

	/**
	 * Method sets "toolbarMenuAllowedProps" attribute of SeriesSettings object.
	 * Uses for hide and shade menu items in context menu.
	 * * true - unshade menu item
	 * * false - shade menu item
	 * * null - hide menu item
	 * @memberof asc_CSeriesSettings
	 * @param {object} val - object with properties of Asc.c_oAscFillType as attribute and boolean or null as value
	 */
	asc_CSeriesSettings.prototype.asc_setToolbarMenuAllowedProps = function (val) {
		this.toolbarMenuAllowedProps = val;
	};


	function CTimeline() {
		this.name = null;
		this.cache = null;
		this.caption = null;

		this.scrollPosition = null;
		this.uid = null;

		this.level = null;
		this.selectionLevel = null;
		this.showHeader = null;

		this.showHorizontalScrollbar = null;
		this.showSelectionLabel = null;
		this.showTimeLevel = null;
		this.style = null;

		return this;
	}
	CTimeline.prototype.clone = function () {
		var res = new CTimeline();

		res.name = this.name;
		res.cache = this.cache;
		res.caption = this.caption;

		res.scrollPosition = this.scrollPosition;
		res.uid = this.uid;

		res.level = this.level;
		res.selectionLevel = this.selectionLevel;
		res.showHeader = this.showHeader;

		res.showHorizontalScrollbar = this.showHorizontalScrollbar;
		res.showSelectionLabel = this.showSelectionLabel;
		res.showTimeLevel = this.showTimeLevel;
		res.style = this.style;

		return res;
	};

	function CTimelineCacheDefinition() {
		this.name = null;
		this.sourceName = null;
		this.uid = null;

		this.pivotTables = null;
		this.state = null;
		this.pivotFilter = null;

		return this;
	}
	CTimelineCacheDefinition.prototype.clone = function () {
		var res = new CTimelineCacheDefinition();

		res.name = this.name;
		res.sourceName = this.sourceName;
		res.uid = this.uid;

		res.pivotTables = this.pivotTables;
		res.state = this.state;
		res.pivotFilter = this.pivotFilter;

		return res;
	};
	CTimelineCacheDefinition.prototype.initPostOpen = function (tableIds, sheetIds) {
		if (this.pivotTables) {
			this.pivotTables.forEach(function(elem){
				elem.initPostOpen(tableIds, sheetIds);
			});
		}
	};

	CTimelineCacheDefinition.prototype.isCachePivotTable = function (sheetId, pivotName) {
		if (this.pivotTables) {
			for (let i = 0; i < this.pivotTables.length; i++) {
				if (this.pivotTables[i].isCachePivotTable(sheetId, pivotName)) {
					return true;
				}
			}
		}
		return false;
	};


	function CTimelineCachePivotTable() {
		this.name = null;
		this.tabId = null;
	}
	CTimelineCachePivotTable.prototype.clone = function () {
		var res = new CTimelineCachePivotTable();

		res.name = this.name;
		res.tabId = this.tabId;

		return res;
	};
	CTimelineCachePivotTable.prototype.initPostOpen = function (tableIds, sheetIds) {
		var ws = null;
		if (null != this.tabId) {
			ws = sheetIds[this.tabId];
			if (ws) {
				this.tabId = ws.getId();
			}
		}
		return ws;
	};
	CTimelineCachePivotTable.prototype.isCachePivotTable = function (sheetId, pivotName) {
		return this.tabId === sheetId && this.name === pivotName;
	};


	function CTimelineState() {
		this.name = null;
		this.filterState = null;

		this.pivotCacheId = null;
		this.minimalRefreshVersion = null;

		this.lastRefreshVersion = null;
		this.filterType = null;

		this.selection = null;
		this.bounds = null;
	}
	CTimelineState.prototype.clone = function () {
		var res = new CTimelineState();

		res.name = this.name;
		res.filterState = this.filterState;

		res.pivotCacheId = this.pivotCacheId;
		res.minimalRefreshVersion = this.minimalRefreshVersion;

		res.lastRefreshVersion = this.lastRefreshVersion;
		res.filterType = this.filterType;

		res.selection = this.selection && this.selection.clone();
		res.bounds = this.bounds && this.bounds.clone();

		return res;
	};

	function CTimelineRange() {
		this.startDate = null;
		this.endDate = null;
	}
	CTimelineRange.prototype.clone = function () {
		var res = new CTimelineRange();

		res.startDate = this.startDate;
		res.endDate = this.endDate;

		return res;
	};

	function CTimelineStyles() {
		this.defaultTimelineStyle = null;
		this.timelineStyles = null;
	}
	CTimelineStyles.prototype.clone = function () {
		var res = new CTimelineRange();

		res.defaultTimelineStyle = this.defaultTimelineStyle;
		if (this.timelineStyles) {
			res.timelineStyles = [];
			for (let i = 0; i < this.timelineStyles.length; i++) {
				res.timelineStyles.push(this.timelineStyles[i].clone());
			}
		}

		return res;
	};

	function CTimelineStyle() {
		this.name = null;
		this.timelineStyleElements = null;
	}
	CTimelineStyle.prototype.clone = function () {
		var res = new CTimelineRange();

		res.name = this.name;
		if (this.timelineStyleElements) {
			res.timelineStyleElements = [];
			for (let i = 0; i < this.timelineStyleElements.length; i++) {
				res.timelineStyleElements.push(this.timelineStyleElements[i].clone());
			}
		}

		return res;
	};

	function CTimelineStyleElement() {
		this.type = null;
		this.dxfId = null;
	}
	CTimelineStyleElement.prototype.clone = function () {
		var res = new CTimelineRange();

		res.type = this.type;
		res.dxfId = this.dxfId;

		return res;
	};

	function CTimelinePivotFilter() {
		this.name = null;
		this.description = null;
		this.useWholeDay = null;
		this.id = null;
		this.fld = null;

		this.autoFilter = null;
	}
	CTimelinePivotFilter.prototype.clone = function () {
		var res = new CTimelinePivotFilter();

		res.name = this.name;
		res.description = this.description;
		res.useWholeDay = this.name;
		res.id = this.description;
		res.fld = this.name;

		res.autoFilter = this.autoFilter && this.autoFilter.clone();

		return res;
	};

	/**
	 * Class representing "Formulas" properties specifies calculation options
	 * @constructor
	 */
	function CCalcPr() {
		this.calcId = null;
		this.calcMode = null;
		this.fullCalcOnLoad = null;
		this.refMode = null;
		this.iterate = null;
		this.iterateCount = null;
		this.iterateDelta = null;
		this.fullPrecision = null;
		this.calcCompleted = null;
		this.calcOnSave = null;
		this.concurrentCalc = null;
		this.concurrentManualCount = null;
		this.forceFullCalc = null;
	}

	/**
	 * Method clones calculation options
	 * @memberof CCalcPr
	 * @returns {CCalcPr}
	 */
	CCalcPr.prototype.clone = function () {
		var res = new CCalcPr();

		res.calcId = this.calcId;
		res.calcMode = this.calcMode;
		res.fullCalcOnLoad = this.fullCalcOnLoad;
		res.refMode = this.refMode;
		res.iterate = this.iterate;
		res.iterateCount = this.iterateCount;
		res.iterateDelta = this.iterateDelta;
		res.fullPrecision = this.fullPrecision;
		res.calcCompleted = this.calcCompleted;
		res.calcOnSave = this.calcOnSave;
		res.concurrentCalc = this.concurrentCalc;
		res.concurrentManualCount = this.concurrentManualCount;
		res.forceFullCalc = this.forceFullCalc;

		return res;
	};
	/**
	 * Method returns "iterate" attribute specifies whether the application should attempt to calculate formulas
	 * that contain circular references.
	 * @memberof CCalcPr
	 * @returns {boolean}
	 */
	CCalcPr.prototype.getIterate = function () {
		return this.iterate;
	};
	/**
	 * Method returns "iterateCount" attribute specifies the number of iterations attempts when calculating a
	 * workbook with circular references, when the "iterate" attribute is true.
	 * @memberof CCalcPr
	 * @returns {number}
	 */
	CCalcPr.prototype.getIterateCount = function () {
		return this.iterateCount;
	};
	/**
	 * Method returns "iterateDelta" attribute specifies the maximum change for iterative calculations.
	 * @memberof CCalcPr
	 * @returns {number}
	 */
	CCalcPr.prototype.getIterateDelta = function () {
		return this.iterateDelta;
	};
	/**
	 * Method returns "calcMode" attribute specifies when the application should calculate formulas in the workbook.
	 * @memberof CCalcPr
	 * @returns {Asc.c_oAscCalcMode}
	 */
	CCalcPr.prototype.getCalcMode = function () {
		return this.calcMode;
	};
	/**
	 * Method sets "iterate" attribute specifies whether the application should attempt to calculate formulas
	 * that contain circular references.
	 * @memberof CCalcPr
	 * @param {boolean} bIterate
	 */
	CCalcPr.prototype.setIterate = function (bIterate) {
		const DEFAULT_ITERATE = false;
		this.iterate = DEFAULT_ITERATE !== bIterate ? bIterate : null;
	};
	/**
	 * Method sets "iterateCount" attribute specifies the number of iterations attempts when calculating a
	 * workbook with circular references, when the "iterate" attribute is true.
	 * @memberof CCalcPr
	 * @param {number} nIterateCount
	 */
	CCalcPr.prototype.setIterateCount = function (nIterateCount) {
		const DEFAULT_ITER_COUNT = 100;
		this.iterateCount = DEFAULT_ITER_COUNT !== nIterateCount ? nIterateCount : null;
	};
	/**
	 * Method sets "iterateDelta" attribute specifies the maximum change for iterative calculations.
	 * @memberof CCalcPr
	 * @param {number} nIterateDelta
	 */
	CCalcPr.prototype.setIterateDelta = function (nIterateDelta) {
		const DEFAULT_ITER_DELTA = 1e-3;
		this.iterateDelta = DEFAULT_ITER_DELTA !== nIterateDelta ? nIterateDelta : null;
	};
	/**
	 * Method updates calcPr attributes
	 * @memberof CCalcPr
	 * @param {asc_CCalcSettings} oCalcSettings
	 * @param {Workbook} oWbModel
	 */
	CCalcPr.prototype.updateCalcProperties = function (oCalcSettings, oWbModel) {
		let bIterativeCalc = oCalcSettings.asc_getIterativeCalc();
		let nMaxIterations = oCalcSettings.asc_getMaxIterations();
		let nMaxChange = oCalcSettings.asc_getMaxChange();

		let bOldIterate = this.getIterate();
		let nOldIterateCount = this.getIterateCount();
		let nOldIterateDelta = this.getIterateDelta();

		this.setIterate(bIterativeCalc);
		this.setIterateCount(nMaxIterations);
		this.setIterateDelta(nMaxChange);

		if (History.Is_On()) {
			let oUpdateSheet = oWbModel.getActiveWs();
			let oUpdateRange = new Asc.Range(0, 0, oUpdateSheet.getColsCount(), oUpdateSheet.getRowsCount());

			if (this.getIterate() !== bOldIterate) {
				History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_CalcPr_iterate,
					oUpdateSheet.getId(), oUpdateRange, new UndoRedoData_FromTo(bOldIterate, this.getIterate()));
			}
			if (this.getIterateCount() !== nOldIterateCount) {
				History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_CalcPr_iterateCount,
					oUpdateSheet.getId(), oUpdateRange, new UndoRedoData_FromTo(nOldIterateCount, this.getIterateCount()));
			}
			if (this.getIterateDelta() !== nOldIterateDelta) {
				History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_CalcPr_iterateDelta,
					oUpdateSheet.getId(), oUpdateRange, new UndoRedoData_FromTo(nOldIterateDelta, this.getIterateDelta()));
			}
		}
	};

	/**
	 * This element defines a collection of workbook properties
	 * @constructor
	 */
	function CWorkbookPr() {
		this.Date1904 = null;
		this.DateCompatibility = null;
		this.HidePivotFieldList = null;
		this.ShowPivotChartFilter = null;
		this.UpdateLinks = null;
	}
	/**
	 * Method clones calculation options
	 * @memberof CWorkbookPr
	 * @returns {CWorkbookPr}
	 */
	CWorkbookPr.prototype.clone = function () {
		let res = new CWorkbookPr();

		res.Date1904 = this.Date1904;
		res.DateCompatibility = this.DateCompatibility;
		res.HidePivotFieldList = this.HidePivotFieldList;
		res.ShowPivotChartFilter = this.ShowPivotChartFilter;
		res.UpdateLinks = this.UpdateLinks;

		return res;
	};
	/**
	 * Method returns "1904" flag
	 * @memberof CWorkbookPr
	 * @returns {boolean}
	 */
	CWorkbookPr.prototype.getDate1904 = function () {
		return this.Date1904;
	};
	/**
	 * Method set "1904" flag
	 * @memberof CWorkbookPr
	 * @returns {boolean}
	 */
	CWorkbookPr.prototype.setDate1904 = function (val) {
		this.Date1904 = val;
	};
	/**
	 * Method returns "DateCompatibility" flag
	 * @memberof CWorkbookPr
	 * @returns {boolean}
	 */
	CWorkbookPr.prototype.getDateCompatibility = function () {
		return this.DateCompatibility;
	};
	/**
	 * Method set "DateCompatibility" flag
	 * @memberof CWorkbookPr
	 * @returns {boolean}
	 */
	CWorkbookPr.prototype.setDateCompatibility = function (val) {
		this.DateCompatibility = val;
	};
	/**
	 * Method returns "HidePivotFieldList" flag
	 * @memberof CWorkbookPr
	 * @returns {boolean}
	 */
	CWorkbookPr.prototype.getHidePivotFieldList = function () {
		return this.HidePivotFieldList;
	};
	/**
	 * Method set "HidePivotFieldList" flag
	 * @memberof CWorkbookPr
	 * @returns {boolean}
	 */
	CWorkbookPr.prototype.setHidePivotFieldList = function (val) {
		this.HidePivotFieldList = val;
	};
	/**
	 * Method returns "ShowPivotChartFilter" flag
	 * @memberof CWorkbookPr
	 * @returns {boolean}
	 */
	CWorkbookPr.prototype.getShowPivotChartFilter = function () {
		return this.ShowPivotChartFilter;
	};
	/**
	 * Method set "ShowPivotChartFilter" flag
	 * @memberof CWorkbookPr
	 * @returns {boolean}
	 */
	CWorkbookPr.prototype.setShowPivotChartFilter = function (val) {
		this.ShowPivotChartFilter = val;
	};
	/**
	 * Method returns "UpdateLinks" flag
	 * @memberof CWorkbookPr
	 * @returns {boolean}
	 */
	CWorkbookPr.prototype.getUpdateLinks = function () {
		return (this.UpdateLinks === Asc.EUpdateLinksType.updatelinksAlways || this.UpdateLinks === Asc.EUpdateLinksType.updatelinksUserSet);
	};
	/**
	 * Method set "UpdateLinks" flag
	 * @memberof CWorkbookPr
	 * @returns {boolean}
	 */
	CWorkbookPr.prototype.setUpdateLinks = function (val) {
		//bool from interface
		if (val === true) {
			val = Asc.EUpdateLinksType.updatelinksAlways;
		}
		if (val === false) {
			val = Asc.EUpdateLinksType.updatelinksNever;
		}
		this.UpdateLinks = val;
	};


	/**
	 * Class representing calculation settings for UI interface
	 * @constructor
	 */
	function asc_CCalcSettings() {
		// Default values if calcPr is empty
		this.bIterativeCalc = false;
		this.nMaxIterations = 100;
		this.nMaxChange = 1e-3;
	}

	/**
	 * Method sets "Enable iterative calculation" setting
	 * @memberof asc_CCalcSettings
	 * @param {boolean} bIterativeCalc
	 */
	asc_CCalcSettings.prototype.asc_setIterativeCalc = function (bIterativeCalc) {
		this.bIterativeCalc = bIterativeCalc
	};
	/**
	 * Method returns "Enable iterative calculation" setting
	 * @memberof asc_CCalcSettings
	 * @returns {boolean}
	 */
	asc_CCalcSettings.prototype.asc_getIterativeCalc = function () {
		return this.bIterativeCalc;
	};
	/**
	 * Method sets "Maximum iterations" setting
	 * @memberof asc_CCalcSettings
	 * @param {number} nMaxIterations
	 */
	asc_CCalcSettings.prototype.asc_setMaxIterations = function (nMaxIterations) {
		this.nMaxIterations = nMaxIterations;
	};
	/**
	 * Method returns "Maximum iterations" setting
	 * @memberof asc_CCalcSettings
	 * @returns {number}
	 */
	asc_CCalcSettings.prototype.asc_getMaxIterations = function () {
		return this.nMaxIterations;
	};
	/**
	 * Method sets "Maximum change" setting
	 * @memberof asc_CCalcSettings
	 * @param {number} nMaxChange
	 */
	asc_CCalcSettings.prototype.asc_setMaxChange = function (nMaxChange) {
		this.nMaxChange = nMaxChange;
	};
	/**
	 * Method returns "Maximum change" setting
	 * @memberof asc_CCalcSettings
	 * @returns {number}
	 */
	asc_CCalcSettings.prototype.asc_getMaxChange = function () {
		return this.nMaxChange;
	};
	/**
	 * Method initializes settings according CalcPr
	 * @memberof asc_CCalcSettings
	 * @param {Workbook} oWbModel
	 */
	asc_CCalcSettings.prototype.asc_initSettings = function (oWbModel) {
		const oCalcPr = oWbModel.calcPr;
		if (oCalcPr.getIterate() != null) {
			this.asc_setIterativeCalc(oCalcPr.getIterate());
		}
		if (oCalcPr.getIterateCount() != null) {
			const MAX_ITERATE_COUNT = 32767;
			this.asc_setMaxIterations(oCalcPr.getIterateCount() <= MAX_ITERATE_COUNT ? oCalcPr.getIterateCount() : MAX_ITERATE_COUNT);
		}
		if (oCalcPr.getIterateDelta() != null) {
			this.asc_setMaxChange(oCalcPr.getIterateDelta());
		}
	};
	/**
	 * Method checks if the given CalcPr object is equal to the current object.
	 * @memberof asc_CCalcSettings
	 * @param {CCalcPr} oCalcPr - The CalcPr to compare with.
	 * @returns {boolean} - True if the objects are equal, false otherwise.
	 */
	asc_CCalcSettings.prototype.asc_isEqual = function (oCalcPr) {
		const DEFAULT_ITERATE = false;
		const DEFAULT_MAX_ITERATE = 100;
		const DEFAULT_MAX_CHANGE = 1e-3;

		const bIterativeCalcEqual = this.asc_getIterativeCalc() === oCalcPr.getIterate();
		const bMaxIterationEqual = this.asc_getMaxIterations() === oCalcPr.getIterateCount();
		const bMaxChangeEqual = this.asc_getMaxChange() === oCalcPr.getIterateDelta();
		const bIterCalcDefaultVal = this.asc_getIterativeCalc() === DEFAULT_ITERATE && oCalcPr.getIterate() == null;
		const bMaxIterDefaultVal = this.asc_getMaxIterations() === DEFAULT_MAX_ITERATE && oCalcPr.getIterateCount() == null;
		const bMaxChangeDefaultVal = this.asc_getMaxChange() === DEFAULT_MAX_CHANGE && oCalcPr.getIterateDelta() == null;

		return (bIterativeCalcEqual && bMaxIterationEqual && bMaxChangeEqual) ||
			(bIterCalcDefaultVal && bMaxIterDefaultVal && bMaxChangeDefaultVal);
	};

	//****metadata****
	function CMetadata() {
		this.metadataTypes = null;
		this.metadataStrings = null;
		this.mdxMetadata = null;
		this.cellMetadata = null;
		this.valueMetadata = null;

		this.aFutureMetadata = null;
	}
	CMetadata.prototype.clone = function () {
		let res = new CMetadata();

		res.metadataTypes = this.metadataTypes && this.metadataTypes.clone();
		res.metadataStrings = this.metadataStrings && this.metadataStrings.clone();
		res.mdxMetadata = this.mdxMetadata && this.mdxMetadata.clone();
		res.cellMetadata = this.cellMetadata && this.cellMetadata.clone();	// CMetadataRecord.clone
		res.valueMetadata = this.valueMetadata && this.valueMetadata.clone();	// CMetadataRecord.clone

		res.aFutureMetadata = this.aFutureMetadata && this.aFutureMetadata.clone();

		return res;
	};

	function CFutureMetadata() {
		this.name = null;
		this.futureMetadataBlocks = null;
	}
	CFutureMetadata.prototype.clone = function () {
		let res = new CFutureMetadata();

		res.name = this.name;
		res.futureMetadataBlocks = this.futureMetadataBlocks && this.futureMetadataBlocks.clone();

		return res;
	};

	function CMetadataType() {
		this.minSupportedVersion = null;
		this.ghostRow = null;
		this.ghostCol = null;
		this.edit = null;
		this.delete = null;
		this.copy = null;
		this.pasteAll = null;
		this.pasteFormulas = null;
		this.pasteValues = null;
		this.pasteFormats = null;
		this.pasteComments = null;
		this.pasteDataValidation = null;
		this.pasteBorders = null;
		this.pasteColWidths = null;
		this.pasteNumberFormats = null;
		this.merge = null;
		this.splitFirst = null;
		this.splitAll = null;
		this.rowColShift = null;
		this.clearAll = null;
		this.clearFormats = null;
		this.clearContents = null;
		this.clearComments = null;
		this.assign = null;
		this.coerce = null;
		this.cellMeta = null;
	}
	CMetadataType.prototype.clone = function () {
		let res = new CMetadataType();

		res.minSupportedVersion = this.minSupportedVersion;
		res.ghostRow = this.ghostRow;
		res.ghostCol = this.ghostCol;
		res.edit = this.edit;
		res.delete = this.delete;
		res.copy = this.copy;
		res.pasteAll = this.pasteAll;
		res.pasteFormulas = this.pasteFormulas;
		res.pasteValues = this.pasteValues;
		res.pasteFormats = this.pasteFormats;
		res.pasteComments = this.pasteComments;
		res.pasteDataValidation = this.pasteDataValidation;
		res.pasteBorders = this.pasteBorders;
		res.pasteColWidths = this.pasteColWidths;
		res.pasteNumberFormats = this.pasteNumberFormats;
		res.merge = this.merge;
		res.splitFirst = this.splitFirst;
		res.splitAll = this.splitAll;
		res.rowColShift = this.rowColShift;
		res.clearAll = this.clearAll;
		res.clearFormats = this.clearFormats;
		res.clearContents = this.clearContents;
		res.clearComments = this.clearComments;
		res.assign = this.assign;
		res.coerce = this.coerce;
		res.cellMeta = this.cellMeta;

		return res;
	};

	function CMetadataString() {
		this.v = null;
	}
	CMetadataString.prototype.clone = function () {
		let res = new CMetadataString();

		res.v = this.v;

		return res;
	};

	function CMdx() {
		this.mdxTuple = null;
		this.mdxSet = null;
		this.mdxKPI = null;
		this.mdxMemeberProp = null;

		this.n = null;
		this.f = null;
	}
	CMdx.prototype.clone = function () {
		let res = new CMdx();

		res.mdxTuple = this.mdxTuple && this.mdxTuple.clone();
		res.mdxSet = this.mdxSet && this.mdxSet.clone();
		res.mdxKPI = this.mdxKPI && this.mdxKPI.clone();

		res.mdxMemeberProp = this.mdxMemeberProp && this.mdxMemeberProp.clone();

		res.n = this.n;
		res.f = this.f;

		return res;
	};

	function CMdxTuple() {
		this.c = null;
		this.ct = null;
		this.si = null;
		this.fi = null;
		this.bc = null;
		this.fc = null;
		this.i = null;
		this.u = null;
		this.st = null;
		this.b = null;

		this.metadataStringIndexes = null;
	}
	CMdxTuple.prototype.clone = function () {
		let res = new CMdxTuple();

		res.c = this.c;
		res.ct = this.ct;
		res.si = this.si;
		res.fi = this.fi;
		res.bc = this.bc;
		res.fc = this.fc;
		res.i = this.i;
		res.u = this.u;
		res.st = this.st;
		res.b = this.b;

		res.metadataStringIndexes = this.metadataStringIndexes && this.metadataStringIndexes.clone();

		return res;
	};

	function CMdxSet() {
		this.ns = null;
		this.c = null;
		this.o = null;

		this.metadataStringIndexes = null;
	}
	CMdxSet.prototype.clone = function () {
		let res = new CMdxSet();

		res.ns = this.ns;
		res.c = this.c;
		res.o = this.o;

		res.metadataStringIndexes = this.metadataStringIndexes && this.metadataStringIndexes.clone();

		return res;
	};

	function CMetadataStringIndex() {
		this.x = null;
		this.s = null;

	}
	CMetadataStringIndex.prototype.clone = function () {
		let res = new CMetadataStringIndex();

		res.x = this.x;
		res.s = this.s;

		return res;
	};

	function CMdxMemeberProp() {
		this.n = null;
		this.np = null;

	}
	CMdxMemeberProp.prototype.clone = function () {
		let res = new CMdxMemeberProp();

		res.n = this.n;
		res.np = this.np;

		return res;
	};

	function CMdxKPI() {
		this.n = null;
		this.np = null;
		this.p = null;
	}
	CMdxKPI.prototype.clone = function () {
		let res = new CMdxKPI();

		res.n = this.n;
		res.np = this.np;
		res.p = this.p;

		return res;
	};

	function CMetadataBlock() {
		this.elems = null;
	}
	CMetadataBlock.prototype.clone = function () {
		let res = new CMetadataBlock();

		res.elems = this.elems;

		return res;
	};

	function CMetadataRecord() {
		this.t = null;
		this.v = null;
	}
	CMetadataRecord.prototype.clone = function () {
		let res = new CMetadataRecord();

		res.t = this.t;
		res.v = this.v;

		return res;
	};

	function CFutureMetadataBlock() {
		this.extLst = null;
	}
	CFutureMetadataBlock.prototype.clone = function () {
		let res = new CFutureMetadataBlock();

		res.extLst = this.extLst && this.extLst.clone();

		return res;
	};

	function CMetadataBlockExt() {
		this.richValueBlock = null;
		this.dynamicArrayProperties = null;
	}
	CMetadataBlockExt.prototype.clone = function () {
		let res = new CMetadataBlockExt();

		res.richValueBlock = this.richValueBlock && this.richValueBlock.clone();
		res.dynamicArrayProperties = this.dynamicArrayProperties && this.dynamicArrayProperties.clone();

		return res;
	};
	//TODO move to extensions?
	function CDynamicArrayProperties() {
		this.fDynamic = null;
		this.fCollapsed = null;
	}
	CDynamicArrayProperties.prototype.clone = function () {
		let res = new CDynamicArrayProperties();

		res.fDynamic = this.fDynamic;
		res.fCollapsed = this.fCollapsed;

		return res;
	};

	function CRichValueBlock() {
		this.i = null;
	}
	CRichValueBlock.prototype.clone = function () {
		let res = new CRichValueBlock();

		res.i = this.i;

		return res;
	};

	function CCustomFunctionEngine(wb) {
		this.wb = wb;
		this.funcsMapInfo = {};

		this.localiztionMap = {};//{en: {"SUM": "SUMMA"}{"SUMMA": "SUM"}}

		this.prefixName = "";
		this.activeLocale = null;

		this.needRecalculate = null;
	}
	CCustomFunctionEngine.prototype.add = function (func, options) {
		//options ->
		/*{"params":
		[
			{
				"defaultValue": ""
				"description": "First number. *"
				"name": "first"
				"isOptional": false
				"parentName": ""
				"type": "number"
			},
			{
				"defaultValue": ""
				"description": "Second number. *"
				"name": "second"
				"isOptional": true
				"parentName": ""
				"type": "string"
			}
		]
		*/

		this._add(func, options);
		this.needRecalculate = true;
	};

	CCustomFunctionEngine.prototype._add = function (func, options) {
		let oThis = this;
		let funcName = func.name.toUpperCase();
		if (funcName.length === 0) {
			console.log("REGISTRAION_ERROR_INVALID_FUNCTION_NAME");
			return;
		}


		//prefix add for separate main function from custom function
		//!!!_xldudf
		funcName = this.prefixName + funcName;

		let oFormulaList = AscCommonExcel.cFormulaFunction;
		if (!this.funcsMapInfo[funcName] && oFormulaList[funcName]) {
			console.log("REGISTRAION_ERROR_CONFLICTED_FUNCTION_NAME");
		}

		if (this.funcsMapInfo[funcName]) {
			this.remove(funcName);
		}

		let params = options && options.params;
		let argsInfo = this._getParamsInfo(func, params);

		let argumentsType = [];
		let argumentsMin = 0;
		let argumentsMax = argsInfo ? argsInfo.length : 0;


		let argsFuncLength = func.length;
		if (argsFuncLength > argumentsMax) {
			console.log("REGISTRAION_ERROR_INVALID_FUNCTION_ARGUMENTS_COUNT");
			return;
		}

		let supportedTypes = {
			"number": 1,
			"string": 1,
			"boolean": 1,
			"any": 1,
			"number[][]": 1,
			"string[][]": 1,
			"boolean[][]": 1,
			"any[][]": 1
		};
		let supportedArrTypes = {
			"number[][]": 1,
			"string[][]": 1,
			"boolean[][]": 1,
			"any[][]": 1
		};
		if (argsInfo) {
			let optionalCount = 0;
			for (let i = 0; i < argsInfo.length; i++) {
				argumentsType.push(this.getTypeByString(argsInfo[i].type));
				if (!supportedTypes[argsInfo[i].type]) {
					let paramName = (params && params[i]) ? params[i].name : "";
					console.log("Registration custom function \"" +  funcName + "\" warning. Invalid param \"" + paramName + "\" type. The following types must be used: number, string, boolean, any, number[][], string[][], boolean[][], any[][].");
				}
				if (argsInfo[i].isOptional) {
					optionalCount++;
				} else {
					optionalCount = 0;
				}
			}
			argumentsMin = argsInfo.length - optionalCount;
		}

		let returnInfo = options.returnInfo;
		if (options.returnInfo && !supportedTypes[returnInfo.type]) {
			console.log("Registration custom function \"" +  funcName + "\" warning. Invalid return type. The following types must be used: number, string, boolean, any, number[][], string[][], boolean[][], any[][].");
		}

		let returnValueType = supportedArrTypes[returnInfo.type] ? AscCommonExcel.cReturnFormulaType.array : null;

		/**
		 * @constructor
		 * @extends {AscCommonExcel.cBaseFunction}
		 */
		function newFunc() {
		}

		//***array-formula***
		newFunc.prototype = Object.create(AscCommonExcel.cBaseFunction.prototype);
		newFunc.prototype.constructor = newFunc;
		newFunc.prototype.name = funcName;
		//newFunc.prototype.argumentsMin = argumentsMin;
		//newFunc.prototype.argumentsMax = argumentsMax;
		//argumentsType - other arguments type, need convert
		newFunc.prototype.argumentsType = argumentsType;
		newFunc.prototype.returnValueType = returnValueType;
		newFunc.prototype.Calculate = function (arg) {
			try {

				if (arg.length < argumentsMin || arg.length > argumentsMax) {
					return new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
				}

				for (let i in arg) {
					if (arg[i] && arg[i].type === AscCommonExcel.cElementType.error) {
						return arg[i];
					}
				}

				//prepare arguments
				let args = [];
				for (let i = 0; i < argsInfo.length; i++) {
					let type = argsInfo[i].type;
					let defaultValue = argsInfo[i].defaultValue;

					if (!arg[i] && !defaultValue) {
						continue;
					}

					if (arg[i] && arg[i].type === AscCommonExcel.cElementType.error && type === AscCommonExcel.cElementType.error) {
						args.push(arg[i].toString());
					} else {
						if (arg[i] && arg[i].type === AscCommonExcel.cElementType.error) {
							return arg[i];
						}

						let elem = oThis.prepareInputArg(arg[i], type, defaultValue);
						if (elem && elem.type === AscCommonExcel.cElementType.error) {
							return elem;
						}

						args.push(elem);
					}
				}

				let res = func.apply(this, args);

				//prepare result
				let returnInfo = options.returnInfo;
				return oThis.prepareResult(res, returnInfo.type);
			} catch (e) {
				console.log("ERROR CUSTOM FUNCTION CALCULATE");
				return  new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
			}
		};

		this.addToFunctionsList(newFunc, options);
	};

	CCustomFunctionEngine.prototype.remove = function (sName) {
		sName = sName.toUpperCase();

		let isFound = false;
		if (AscCommonExcel.cFormulaFunctionGroup["Custom"]) {
			let aCustomFunc = AscCommonExcel.cFormulaFunctionGroup["Custom"];
			for (let i in aCustomFunc) {
				if (aCustomFunc[i] && aCustomFunc[i].prototype && aCustomFunc[i].prototype.name === sName) {
					aCustomFunc.splice(i - 0, 1);
					isFound = true;
					break;
				}
			}
		}

		if (isFound) {
			AscCommonExcel.removeCustomFunction(sName);
			this.wb.initFormulasList && this.wb.initFormulasList();
			if (this.wb && this.wb.Api) {
				this.wb.Api.formulasList = AscCommonExcel.getFormulasInfo();
			}
			this.removeInfo(sName);

			this.wb.handlers && this.wb.handlers.trigger("asc_onRemoveCustomFunction");
			return true;
		}
		return false;
	};

	CCustomFunctionEngine.prototype.clear = function () {
		if (AscCommonExcel.cFormulaFunctionGroup["Custom"] && AscCommonExcel.cFormulaFunctionGroup["Custom"].length) {
			let aCustomFunc = AscCommonExcel.cFormulaFunctionGroup["Custom"];
			for (let i = 0; i < aCustomFunc.length; i++) {
				let sName = aCustomFunc[i].prototype.name;
				AscCommonExcel.removeCustomFunction(sName);
			}
			AscCommonExcel.cFormulaFunctionGroup["Custom"] = [];

			this.wb.initFormulasList && this.wb.initFormulasList();
			if (this.wb && this.wb.Api) {
				this.wb.Api.formulasList = AscCommonExcel.getFormulasInfo();
			}
			this.clearInfo();

			this.wb.handlers && this.wb.handlers.trigger("asc_onRemoveCustomFunction");

			return true;
		}

		return false;
	};

	CCustomFunctionEngine.prototype.setActiveLocale = function (sLocale) {
		this.activeLocale = sLocale;
	};

	CCustomFunctionEngine.prototype.getActiveLocale = function () {
		return this.activeLocale;
	};

	CCustomFunctionEngine.prototype._getParamsInfo = function (func, params) {
		let aArgs = this._getArgsByFunc(func);
		let argsInfo = [];
		if (!aArgs) {
			return argsInfo;
		}

		let paramsMap = {};
		if (params) {
			for (let i in params) {
				paramsMap[params[i].name] = params[i];
			}
		}
		for (let i = 0; i < aArgs.length; i++) {
			let type = "any";
			let curParams = paramsMap && paramsMap[aArgs[i]];
			let _isOptional = false;
			let _defaultValue = null;
			if (curParams) {
				if (curParams.type) {
					type = curParams.type;
				}
				_isOptional = curParams.isOptional;
				_defaultValue = curParams.defaultValue;
			}
			argsInfo.push({type: type, isOptional: _isOptional, defaultValue: _defaultValue});
		}
		return argsInfo;
	};

	CCustomFunctionEngine.prototype._getArgsByFunc = function (func) {
		const funcCommentsRegExp = /((\/\/.*$)|(\/\*[\s\S]*?\*\/))/mg;
		const funcArgsNamesRegExp = /([^\s,]+)/g;
		let sFunc = func.toString().replace(funcCommentsRegExp, '');
		return sFunc.slice(sFunc.indexOf('(') + 1, sFunc.indexOf(')')).match(funcArgsNamesRegExp);
	};

	CCustomFunctionEngine.prototype.addToFunctionsList = function (newFunc, params) {
		AscCommonExcel.cFormulaFunctionGroup['Custom'] = AscCommonExcel.cFormulaFunctionGroup['Custom'] || [];

		let translations = params.nameLocale;
		let description = params.description;
		let args = params.params;

		let funcName = newFunc.prototype.name;

		//already added function
		let isNewFunc = false;
		if (this.funcsMapInfo[funcName]) {
			//reload translations
			this.pushTranslations(funcName, translations);
			//reload description
			this.funcsMapInfo[funcName].description = description;

			//reload args info
			this.funcsMapInfo[funcName].args = [];
			if (args) {
				for (let i = 0; i < args.length; i++) {
					this.funcsMapInfo[funcName].args.push(new CCustomFunctionArgInfo(args[i].name, args[i].isOptional));
				}
			}

			let customFunctionList = AscCommonExcel.cFormulaFunctionGroup["Custom"];
			for (let i in customFunctionList) {
				if (customFunctionList[i] && customFunctionList[i].prototype.name === funcName) {
					customFunctionList.splice(i - 0, 1);
					break;
				}
			}
		} else {
			this.funcsMapInfo[funcName] = new CCustomFunctionInfo(funcName);
			this.pushTranslations(funcName, translations);
			this.funcsMapInfo[funcName].description = description;

			if (args) {
				for (let i = 0; i < args.length; i++) {
					if (!this.funcsMapInfo[funcName].args) {
						this.funcsMapInfo[funcName].args = [];
					}
					this.funcsMapInfo[funcName].args.push(new CCustomFunctionArgInfo(args[i].name, args[i].isOptional));
				}
			}
			isNewFunc = true;
		}

		//add or reload
		if (AscCommonExcel.cFormulaFunctionToLocale && ((!this.funcsMapInfo[funcName].addLocalization
			&& !AscCommonExcel.cFormulaFunctionToLocale[funcName]) || this.funcsMapInfo[funcName].addLocalization)) {

			//need get from interface short formula lang("en", ...)
			let localName = this.getTranslationName(funcName, this.activeLocale);
			AscCommonExcel.cFormulaFunctionLocalized[localName] = newFunc;
			AscCommonExcel.cFormulaFunctionToLocale[funcName] = localName;

			this.funcsMapInfo[funcName].addLocalization = true;
			this.funcsMapInfo[funcName].description = description;
		}

		AscCommonExcel.cFormulaFunctionGroup["Custom"].push(newFunc);
		AscCommonExcel.addNewFunction(newFunc);
		this.wb.initFormulasList && this.wb.initFormulasList();
		if (this.wb && this.wb.Api) {
			this.wb.Api.formulasList = AscCommonExcel.getFormulasInfo();
		}
		if (isNewFunc) {
			this.wb.handlers && this.wb.handlers.trigger("asc_onAddCustomFunction");
		}
	};

	CCustomFunctionEngine.prototype.pushTranslations = function (funcName, translations) {
		for (let i in translations) {
			if (!translations[i]) {
				continue;
			}
			if (!this.localiztionMap[i]) {
				this.localiztionMap[i] = {};
			}
			if (!this.localiztionMap[i].fullNameToLocalName) {
				this.localiztionMap[i].fullNameToLocalName = {};
			}
			if (!this.localiztionMap[i].localNameToFullName) {
				this.localiztionMap[i].localNameToFullName = {};
			}


			this.localiztionMap[i].fullNameToLocalName[funcName] = (translations[i] + "").toUpperCase();
			this.localiztionMap[i].localNameToFullName[translations[i]] = funcName;
		}
	};

	CCustomFunctionEngine.prototype.getTypeByString = function (_type) {
		let res =  AscCommonExcel.cElementType.number;
		switch (_type) {
			case "number":
				res = AscCommonExcel.cElementType.number;
				break;
			case "string":
				res = AscCommonExcel.cElementType.string;
				break;
			case "boolean":
				res = AscCommonExcel.cElementType.bool;
				break;
			default:
				res = AscCommonExcel.cElementType.any;
				break;
		}
		return res;
	};

	CCustomFunctionEngine.prototype.prepareInputArg = function (_elem, _type, _defaultValue) {
		if (!_elem) {
			return _defaultValue;
		}

		let res = null;
		switch (_type) {
			case "number":
				if (_elem.type === AscCommonExcel.cElementType.error) {
					return _elem;
				}
				if (_elem.type === AscCommonExcel.cElementType.array || _elem.type === AscCommonExcel.cElementType.cellsRange || _elem.type === AscCommonExcel.cElementType.cellsRange3D) {
					//TODO ms -> calc error
					res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
				} else {
					res = _elem.tocNumber();
					if (res.type !== AscCommonExcel.cElementType.error) {
						res = res.toNumber();
					} else {
						res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
					}
				}
				break;
			case "string":
				if (_elem.type === AscCommonExcel.cElementType.array || _elem.type === AscCommonExcel.cElementType.cellsRange || _elem.type === AscCommonExcel.cElementType.cellsRange3D) {
					//TODO ms -> calc error
					res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
				} else {
					res = _elem.tocString();
					if (res.type !== AscCommonExcel.cElementType.error) {
						res = res.toString();
					} else {
						res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
					}
				}
				break;
			case "boolean":
				if (_elem.type === AscCommonExcel.cElementType.array || _elem.type === AscCommonExcel.cElementType.cellsRange || _elem.type === AscCommonExcel.cElementType.cellsRange3D) {
					//TODO ms -> calc error
					res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
				} else {
					res = _elem.tocBool();
					if (res.type !== AscCommonExcel.cElementType.error && res.toBool) {
						res = res.toBool();
					} else {
						res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
					}
				}
				break;
			case "any":
				if (_elem.type === AscCommonExcel.cElementType.array || _elem.type === AscCommonExcel.cElementType.cellsRange || _elem.type === AscCommonExcel.cElementType.cellsRange3D) {
					//TODO ms -> calc error
					res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
				} else {
					if (_elem.type === AscCommonExcel.cElementType.cell || _elem.type === AscCommonExcel.cElementType.cell3D) {
						_elem = _elem.getValue();
					}
					res = _elem.getValue();
				}
				break;
			case "number[][]":
				res = _elem.toArray(true, true, function (elem) {
					return elem.tocNumber();
				});
				break;
			case "string[][]":
				res = _elem.toArray(true, true, function (elem) {
					return elem.tocString();
				});
				break;
			case "boolean[][]":
				res = _elem.toArray(true, true, function (arrayElem) {
					let _res = arrayElem.tocBool();
					if (_res.type !== AscCommonExcel.cElementType.error && _res.toBool) {
						return _res;
					} else {
						return new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
					}
				});
				break;
			case "any[][]":
				res = _elem.toArray(true, true);
				break;
			default:
				res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
				break;
		}
		return res;
	};

	CCustomFunctionEngine.prototype.getFunc = function (name, bLocale) {
		if (bLocale) {
			let activeLocale = this.activeLocale;
			if (this.localiztionMap[activeLocale]) {
				if (this.localiztionMap[activeLocale].localNameToFullName[name]) {
					name = this.localiztionMap[activeLocale].localNameToFullName[name];
				}
			}
		}

		return this.funcsMapInfo[name];
	};
	
	CCustomFunctionEngine.prototype.getDescription = function (name, ignoreLocale) {
		let res = null;

		let activeLocale = this.activeLocale;
		if (!ignoreLocale && this.localiztionMap[activeLocale]) {
			if (this.localiztionMap[activeLocale].localNameToFullName[name]) {
				name = this.localiztionMap[activeLocale].localNameToFullName[name];
			}
		}

		if (this.funcsMapInfo[name]) {
			res = this.funcsMapInfo[name].description;
		}

		return res;
	};

	CCustomFunctionEngine.prototype.getTranslationName = function (name, lang) {
		let res = name;
		if (this.localiztionMap[lang]) {
			let localName = this.localiztionMap[lang].fullNameToLocalName[name];
			if (localName) {
				res = this.localiztionMap[lang].fullNameToLocalName[name];
			}
		}
		return res;
	};

	CCustomFunctionEngine.prototype.prepareResult = function (val, _type) {
		let res = null;
		switch (_type) {
			case "number":
				if (typeof val === "object") {
					res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
				} else {
					res = this._toNumber(val);
				}
				break;
			case "string":
				if (typeof val === "object") {
					res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
				} else {
					res = this._toString(val);
				}
				break;
			case "boolean":
				if (typeof val === "object") {
					res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
				} else {
					res = this._toBool(val);
				}
				break;
			case "any":
				if (typeof val === "object") {
					res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
				} else {
					res = this._toAny(val);
				}
				break;
			case "number[][]":
				if (Asc.typeOf(val) !== "array" || !val[0]) {
					res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
				} else {
					res = this._tocArray(val, AscCommonExcel.cElementType.number, true);
				}
				break;
			case "string[][]":
				if (Asc.typeOf(val) !== "array" || !val[0]) {
					res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
				} else {
					res = this._tocArray(val, AscCommonExcel.cElementType.string, true);
				}
				break;
			case "boolean[][]":
				if (Asc.typeOf(val) !== "array" || !val[0]) {
					res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
				} else {
					res = this._tocArray(val, AscCommonExcel.cElementType.bool, true);
				}
				break;
			case "any[][]":
				if (Asc.typeOf(val) !== "array" || !val[0]) {
					res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
				} else {
					res = this._tocArray(val, null, true);
				}
				break;
			default:
				res = new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);


		}
		return res;
	};

	CCustomFunctionEngine.prototype._toNumber = function (val) {
		let res;
		if (val === true || val === false) {
			res = new AscCommonExcel.cBool(val);
		} else {
			res = typeof val === "string" ? new AscCommonExcel.cString(val) : new AscCommonExcel.cNumber(val);
		}
		return res;
	};

	CCustomFunctionEngine.prototype._toString = function (val) {
		return new AscCommonExcel.cString(val + "");
	};

	CCustomFunctionEngine.prototype._toBool = function (val) {
		if (this._checkBool(val)) {
			return new AscCommonExcel.cBool(val);
		} else {
			if (Number.isFinite(val)) {
				return new AscCommonExcel.cNumber(val);
			} else {
				return new AscCommonExcel.cString(val);
			}
		}
	};

	CCustomFunctionEngine.prototype._toAny = function (val) {
		if (this._checkBool(val)) {
			return new AscCommonExcel.cBool(val);
		} else {
			if (Number.isFinite(val)) {
				return new AscCommonExcel.cNumber(val);
			} else {
				return new AscCommonExcel.cString(val);
			}
		}
	};

	CCustomFunctionEngine.prototype._tocArray = function (array, resType, checkOnError) {
		var oArray = [], _res = new AscCommonExcel.cArray();

		function isOneDimensional(arr) {
			if (!Array.isArray(arr)) {
				return false;
			}

			for (let i = 0; i < arr.length; i++) {
				if (Array.isArray(arr[i])) {
					return false;
				}
			}
			return true;
		}

		if (isOneDimensional(array)) {
			let newArr = [];
			newArr.push(array);
			array = newArr;
		}

		for (var i = 0; i < array.length; i++) {
			for (var j = 0; j < array[i].length; j++) {
				if (typeof array[i][j] === "object") {
					return new AscCommonExcel.cError(AscCommonExcel.cErrorType.wrong_value_type);
				}

				let isError = checkOnError && this._checkOnErrorByString(array[i][j]);
				if (isError) {
					return isError;
				}
				if (!oArray[i]) {
					oArray[i] = [];
				}

				switch (resType) {
					case AscCommonExcel.cElementType.number:
						oArray[i][j] = this._toNumber((array[i][j]));
						break;
					case AscCommonExcel.cElementType.string:
						oArray[i][j] = this._toString((array[i][j]));
						break;
					case AscCommonExcel.cElementType.bool:
						oArray[i][j] = this._toBool((array[i][j]));
						break;
					case null://any
						oArray[i][j] = this._toAny((array[i][j]));
						break;
				}
			}
		}

		_res.fillFromArray(oArray);

		return _res;
	};

	CCustomFunctionEngine.prototype._checkBool = function (val) {
		if (val === true || val === false) {
			return true;
		} else {
			val = val.toLowerCase && val.toLowerCase();
			if (val === "true" || val === "false") {
				return true;
			}
		}
		return false;
	};

	CCustomFunctionEngine.prototype._checkOnErrorByString = function (_str) {
		for (let i in window["AscCommon"].cErrorOrigin) {
			if (window["AscCommon"].cErrorOrigin[i] === _str) {
				return new AscCommonExcel.cError(_str);
			}
		}
		for (let i in window["AscCommon"].cErrorLocal) {
			if (window["AscCommon"].cErrorLocal[i] === _str) {
				return new AscCommonExcel.cError(_str);
			}
		}
		return false;
	};

	CCustomFunctionEngine.prototype.removeInfo = function (sName) {
		if (this.funcsMapInfo[sName]) {
			delete this.funcsMapInfo[sName];
		}
		this.removeLocalizationInfo(sName);
	};

	CCustomFunctionEngine.prototype.removeLocalizationInfo = function (sName) {
		for (let i in this.localiztionMap) {
			if (this.localiztionMap.hasOwnProperty(i)) {
				if (this.localiztionMap[i].fullNameToLocalName && this.localiztionMap[i].fullNameToLocalName[sName]) {
					delete this.localiztionMap[i].fullNameToLocalName[sName];
				}

				if (this.localiztionMap[i].localNameToFullName) {
					for (let j in this.localiztionMap[i].localNameToFullName) {
						if (this.localiztionMap[i].localNameToFullName.hasOwnProperty(j)) {
							if (this.localiztionMap[i].localNameToFullName[j] === sName) {
								delete this.localiztionMap[i].localNameToFullName[j];
								break;
							}
						}
					}
				}
			}
		}
	};

	CCustomFunctionEngine.prototype.clearInfo = function () {
		this.funcsMapInfo = {};
		this.localiztionMap = {};
	};

	function CCustomFunctionInfo(name) {
		this.name = name;
		this.description = null;

		this.args = null;

		this.addLocalization = null;
	}
	CCustomFunctionInfo.prototype.asc_getDescription = function () {
		return this.description;
	};
	CCustomFunctionInfo.prototype.asc_getArg = function (num) {
		if (num == null) {
			return this.args;
		}
		if (this.args && this.args[num]) {
			return this.args[num];
		}
		return null;
	};

	function CCustomFunctionArgInfo(sName, bOptional) {
		this.sName = sName;
		this.bOptional = bOptional;
	}
	CCustomFunctionArgInfo.prototype.asc_getName = function () {
		return this.sName;
	};
	CCustomFunctionArgInfo.prototype.asc_getIsOptional = function () {
		return this.bOptional;
	};

	function CWorkbookInfo(name, id) {
		this.name = name;
		this.id = id;

		this.sheets = null;
	}
	CWorkbookInfo.prototype.addSheet = function (name, index) {
		if (!this.sheets) {
			this.sheets = [];
		}
		let newObj = new CWorksheetInfo(name, index);
		this.sheets.push(newObj);
	};
	CWorkbookInfo.prototype.asc_getName = function () {
		return this.name;
	};
	CWorkbookInfo.prototype.asc_getId = function () {
		return this.id;
	};
	CWorkbookInfo.prototype.asc_getSheets = function () {
		return this.sheets;
	};

	function CWorksheetInfo(name, index) {
		this.name = name;
		this.index = index;
	}
	CWorksheetInfo.prototype.asc_getName = function () {
		return this.name;
	};
	CWorksheetInfo.prototype.asc_getIndex = function () {
		return this.index;
	};

	//----------------------------------------------------------export----------------------------------------------------
	var prot;
	window['Asc'] = window['Asc'] || {};
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].g_oColorManager = g_oColorManager;
	window['AscCommonExcel'].g_oDefaultFormat = g_oDefaultFormat;
	window['AscCommonExcel'].g_nColorTextDefault = g_nColorTextDefault;
	window['AscCommonExcel'].g_nColorHyperlink = g_nColorHyperlink;
	window['AscCommonExcel'].c_maxOutlineLevel = c_maxOutlineLevel;
	window['AscCommonExcel'].g_oThemeColorsDefaultModsSpreadsheet = g_oThemeColorsDefaultModsSpreadsheet;
	window['AscCommonExcel'].g_StyleCache = g_StyleCache;
	window['AscCommonExcel'].map_themeExcel_to_themePresentation = map_themeExcel_to_themePresentation;
	window['AscCommonExcel'].g_nRowStructSize = g_nRowStructSize;
	window['AscCommonExcel'].shiftGetBBox = shiftGetBBox;
	window['AscCommonExcel'].getStringFromMultiText = getStringFromMultiText;
	window['AscCommonExcel'].getStringFromMultiTextSkipToSpace = getStringFromMultiTextSkipToSpace;
	window['AscCommonExcel'].isEqualMultiText = isEqualMultiText;
	window['AscCommonExcel'].RgbColor = RgbColor;
	window['AscCommonExcel'].createRgbColor = createRgbColor;
	window['AscCommonExcel'].ThemeColor = ThemeColor;
	window['AscCommonExcel'].CorrectAscColor = CorrectAscColor;
	window['AscCommonExcel'].Fragment = Fragment;
	window['AscCommonExcel'].Font = Font;
	window["Asc"]["c_oAscPatternType"] = c_oAscPatternType;
	prot = c_oAscPatternType;
	prot["DarkDown"] = prot.DarkDown;
	prot["DarkGray"] = prot.DarkGray;
	prot["DarkGrid"] = prot.DarkGrid;
	prot["DarkHorizontal"] = prot.DarkHorizontal;
	prot["DarkTrellis"] = prot.DarkTrellis;
	prot["DarkUp"] = prot.DarkUp;
	prot["DarkVertical"] = prot.DarkVertical;
	prot["Gray0625"] = prot.Gray0625;
	prot["Gray125"] = prot.Gray125;
	prot["LightDown"] = prot.LightDown;
	prot["LightGray"] = prot.LightGray;
	prot["LightGrid"] = prot.LightGrid;
	prot["LightHorizontal"] = prot.LightHorizontal;
	prot["LightTrellis"] = prot.LightTrellis;
	prot["LightUp"] = prot.LightUp;
	prot["LightVertical"] = prot.LightVertical;
	prot["MediumGray"] = prot.MediumGray;
	prot["None"] = prot.None;
	prot["Solid"] = prot.Solid;
	window["Asc"]["asc_CGradientFill"] = window['AscCommonExcel'].GradientFill = GradientFill;
	prot = GradientFill.prototype;
	prot["asc_getType"] = prot.asc_getType;
	prot["asc_setType"] = prot.asc_setType;
	prot["asc_getDegree"] = prot.asc_getDegree;
	prot["asc_setDegree"] = prot.asc_setDegree;
	prot["asc_getLeft"] = prot.asc_getLeft;
	prot["asc_setLeft"] = prot.asc_setLeft;
	prot["asc_getRight"] = prot.asc_getRight;
	prot["asc_setRight"] = prot.asc_setRight;
	prot["asc_getTop"] = prot.asc_getTop;
	prot["asc_setTop"] = prot.asc_setTop;
	prot["asc_getBottom"] = prot.asc_getBottom;
	prot["asc_setBottom"] = prot.asc_setBottom;
	prot["asc_getGradientStops"] = prot.asc_getGradientStops;
	prot["asc_putGradientStops"] = prot.asc_putGradientStops;
	window["Asc"]["asc_CGradientStop"] = window['AscCommonExcel'].GradientStop = GradientStop;
	prot = GradientStop.prototype;
	prot["asc_getPosition"] = prot.asc_getPosition;
	prot["asc_setPosition"] = prot.asc_setPosition;
	prot["asc_getColor"] = prot.asc_getColor;
	prot["asc_setColor"] = prot.asc_setColor;
	window["Asc"]["asc_CPatternFill"] = window['AscCommonExcel'].PatternFill = PatternFill;
	prot = PatternFill.prototype;
	prot["asc_getType"] = prot.asc_getType;
	prot["asc_setType"] = prot.asc_setType;
	prot["asc_getFgColor"] = prot.asc_getFgColor;
	prot["asc_setFgColor"] = prot.asc_setFgColor;
	prot["asc_getBgColor"] = prot.asc_getBgColor;
	prot["asc_setBgColor"] = prot.asc_setBgColor;
	window["Asc"]["asc_CFill2"] = window['AscCommonExcel'].Fill = Fill;
	prot = Fill.prototype;
	prot["asc_getPatternFill"] = prot.asc_getPatternFill;
	prot["asc_setPatternFill"] = prot.asc_setPatternFill;
	prot["asc_getGradientFill"] = prot.asc_getGradientFill;
	prot["asc_setGradientFill"] = prot.asc_setGradientFill;
	window['AscCommonExcel'].BorderProp = BorderProp;
	window['AscCommonExcel'].Border = Border;
	window['AscCommonExcel'].Num = Num;
	window["Asc"]["asc_CellXfs"] = window['AscCommonExcel'].CellXfs = CellXfs;
	prot = CellXfs.prototype;
	prot["asc_getFillColor"] = prot.asc_getFillColor;
	prot["asc_getFill"] = prot.asc_getFill;
	prot["asc_getFontName"] = prot.asc_getFontName;
    prot["asc_getFontSize"] = prot.asc_getFontSize;
    prot["asc_getFontColor"] = prot.asc_getFontColor;
    prot["asc_getFontBold"] = prot.asc_getFontBold;
    prot["asc_getFontItalic"] = prot.asc_getFontItalic;
    prot["asc_getFontUnderline"] = prot.asc_getFontUnderline;
    prot["asc_getFontStrikeout"] = prot.asc_getFontStrikeout;
    prot["asc_getFontSubscript"] = prot.asc_getFontSubscript;
    prot["asc_getFontSuperscript"] = prot.asc_getFontSuperscript;
	prot["asc_getNumFormat"] = prot.asc_getNumFormat;
	prot["asc_getNumFormatInfo"] = prot.asc_getNumFormatInfo;
	prot["asc_getHorAlign"] = prot.asc_getHorAlign;
	prot["asc_getVertAlign"] = prot.asc_getVertAlign;
	prot["asc_getAngle"] = prot.asc_getAngle;
	prot["asc_getIndent"] = prot.asc_getIndent;
	prot["asc_getWrapText"] = prot.asc_getWrapText;
	prot["asc_getShrinkToFit"] = prot.asc_getShrinkToFit;
	prot["asc_getPreview"] = prot.asc_getPreview;
	prot["asc_getLocked"] = prot.asc_getLocked;
	prot["asc_getHidden"] = prot.asc_getHidden;
	prot["asc_setFillColor"] = prot.asc_setFillColor;
	prot["asc_setFill"] = prot.asc_setFill;
	prot["asc_setFontName"] = prot.asc_setFontName;
	prot["asc_setFontSize"] = prot.asc_setFontSize;
	prot["asc_setFontColor"] = prot.asc_setFontColor;
	prot["asc_setFontBold"] = prot.asc_setFontBold;
	prot["asc_setFontItalic"] = prot.asc_setFontItalic;
	prot["asc_setFontUnderline"] = prot.asc_setFontUnderline;
	prot["asc_setFontStrikeout"] = prot.asc_setFontStrikeout;
	prot["asc_setFontSubscript"] = prot.asc_setFontSubscript;
	prot["asc_setFontSuperscript"] = prot.asc_setFontSuperscript;
	prot["asc_setBorder"] = prot.asc_setBorder;
	prot["asc_setNumFormatInfo"] = prot.asc_setNumFormatInfo;


	window['AscCommonExcel'].Align = Align;
	window['AscCommonExcel'].CCellStyles = CCellStyles;
	window['AscCommonExcel'].CCellStyle = CCellStyle;
	window['AscCommonExcel'].StyleManager = StyleManager;
	window['AscCommonExcel'].SheetMergedStyles = SheetMergedStyles;
	window['AscCommonExcel'].Hyperlink = Hyperlink;
	window['AscCommonExcel'].SheetFormatPr = SheetFormatPr;
	window['AscCommonExcel'].Col = Col;
	window['AscCommonExcel'].Row = Row;
	window['AscCommonExcel'].CMultiTextElem = CMultiTextElem;
	window['AscCommonExcel'].CCellValue = CCellValue;
	window['AscCommonExcel'].RangeDataManager = RangeDataManager;
	window['AscCommonExcel'].CSharedStrings = CSharedStrings;
	window['AscCommonExcel'].CWorkbookFormulas = CWorkbookFormulas;
	window["Asc"]["sparklineGroup"] = window['AscCommonExcel'].sparklineGroup = sparklineGroup;
	prot = sparklineGroup.prototype;
	prot["asc_getId"] = prot.asc_getId;
	prot["asc_getType"] = prot.asc_getType;
	prot["asc_getLineWeight"] = prot.asc_getLineWeight;
	prot["asc_getDisplayEmpty"] = prot.asc_getDisplayEmpty;
	prot["asc_getMarkersPoint"] = prot.asc_getMarkersPoint;
	prot["asc_getHighPoint"] = prot.asc_getHighPoint;
	prot["asc_getLowPoint"] = prot.asc_getLowPoint;
	prot["asc_getFirstPoint"] = prot.asc_getFirstPoint;
	prot["asc_getLastPoint"] = prot.asc_getLastPoint;
	prot["asc_getNegativePoint"] = prot.asc_getNegativePoint;
	prot["asc_getDisplayXAxis"] = prot.asc_getDisplayXAxis;
	prot["asc_getDisplayHidden"] = prot.asc_getDisplayHidden;
	prot["asc_getMinAxisType"] = prot.asc_getMinAxisType;
	prot["asc_getMaxAxisType"] = prot.asc_getMaxAxisType;
	prot["asc_getRightToLeft"] = prot.asc_getRightToLeft;
	prot["asc_getManualMax"] = prot.asc_getManualMax;
	prot["asc_getManualMin"] = prot.asc_getManualMin;
	prot["asc_getColorSeries"] = prot.asc_getColorSeries;
	prot["asc_getColorNegative"] = prot.asc_getColorNegative;
	prot["asc_getColorAxis"] = prot.asc_getColorAxis;
	prot["asc_getColorMarkers"] = prot.asc_getColorMarkers;
	prot["asc_getColorFirst"] = prot.asc_getColorFirst;
	prot["asc_getColorLast"] = prot.asc_getColorLast;
	prot["asc_getColorHigh"] = prot.asc_getColorHigh;
	prot["asc_getColorLow"] = prot.asc_getColorLow;
	prot["asc_getDataRanges"] = prot.asc_getDataRanges;
	prot["asc_setType"] = prot.asc_setType;
	prot["asc_setLineWeight"] = prot.asc_setLineWeight;
	prot["asc_setDisplayEmpty"] = prot.asc_setDisplayEmpty;
	prot["asc_setMarkersPoint"] = prot.asc_setMarkersPoint;
	prot["asc_setHighPoint"] = prot.asc_setHighPoint;
	prot["asc_setLowPoint"] = prot.asc_setLowPoint;
	prot["asc_setFirstPoint"] = prot.asc_setFirstPoint;
	prot["asc_setLastPoint"] = prot.asc_setLastPoint;
	prot["asc_setNegativePoint"] = prot.asc_setNegativePoint;
	prot["asc_setDisplayXAxis"] = prot.asc_setDisplayXAxis;
	prot["asc_setDisplayHidden"] = prot.asc_setDisplayHidden;
	prot["asc_setMinAxisType"] = prot.asc_setMinAxisType;
	prot["asc_setMaxAxisType"] = prot.asc_setMaxAxisType;
	prot["asc_setRightToLeft"] = prot.asc_setRightToLeft;
	prot["asc_setManualMax"] = prot.asc_setManualMax;
	prot["asc_setManualMin"] = prot.asc_setManualMin;
	prot["asc_setColorSeries"] = prot.asc_setColorSeries;
	prot["asc_setColorNegative"] = prot.asc_setColorNegative;
	prot["asc_setColorAxis"] = prot.asc_setColorAxis;
	prot["asc_setColorMarkers"] = prot.asc_setColorMarkers;
	prot["asc_setColorFirst"] = prot.asc_setColorFirst;
	prot["asc_setColorLast"] = prot.asc_setColorLast;
	prot["asc_setColorHigh"] = prot.asc_setColorHigh;
	prot["asc_setColorLow"] = prot.asc_setColorLow;
	prot["asc_getStyles"] = prot.asc_getStyles;
	prot["asc_setStyle"] = prot.asc_setStyle;
	window['AscCommonExcel'].sparkline = sparkline;
	window['AscCommonExcel'].TablePart = TablePart;
	window['AscCommonExcel'].AutoFilter = AutoFilter;
	window['AscCommonExcel'].SortState = SortState;
	window['AscCommonExcel'].TableColumn = TableColumn;
	window['AscCommonExcel'].TableStyleInfo = TableStyleInfo;
	window['AscCommonExcel'].FilterColumn = FilterColumn;
	window['AscCommonExcel'].Filters = Filters;
	window['AscCommonExcel'].Filter = Filter;
	window['AscCommonExcel'].DateGroupItem = DateGroupItem;
	window['AscCommonExcel'].SortCondition = SortCondition;
	window['AscCommonExcel'].AutoFilterDateElem = AutoFilterDateElem;
	window['AscCommonExcel'].QueryTable = QueryTable;
	window['AscCommonExcel'].QueryTableRefresh = QueryTableRefresh;
	window['AscCommonExcel'].QueryTableField = QueryTableField;
	window['AscCommonExcel'].QueryTableDeletedField = QueryTableDeletedField;
	window['AscCommonExcel'].c_oAscPatternType = c_oAscPatternType;

	window['AscCommonExcel'].OleSizeSelectionRange = OleSizeSelectionRange;


	window["Asc"]["CustomFilters"]			= window["Asc"].CustomFilters = CustomFilters;
	prot									= CustomFilters.prototype;
	prot["asc_getAnd"]						= prot.asc_getAnd;
	prot["asc_getCustomFilters"]			= prot.asc_getCustomFilters;
	prot["asc_setAnd"]						= prot.asc_setAnd;
	prot["asc_setCustomFilters"]			= prot.asc_setCustomFilters;

	window["Asc"]["CustomFilter"]			= window["Asc"].CustomFilter = CustomFilter;
	prot									= CustomFilter.prototype;
	prot["asc_getOperator"]					= prot.asc_getOperator;
	prot["asc_getVal"]						= prot.asc_getVal;
	prot["asc_setOperator"]					= prot.asc_setOperator;
	prot["asc_setVal"]						= prot.asc_setVal;

	window["Asc"]["DynamicFilter"]			= window["Asc"].DynamicFilter = DynamicFilter;
	prot									= DynamicFilter.prototype;
	prot["asc_getType"]						= prot.asc_getType;
	prot["asc_getVal"]						= prot.asc_getVal;
	prot["asc_getMaxVal"]					= prot.asc_getMaxVal;
	prot["asc_setType"]						= prot.asc_setType;
	prot["asc_setVal"]						= prot.asc_setVal;
	prot["asc_setMaxVal"]					= prot.asc_setMaxVal;

	window["Asc"]["ColorFilter"]			= window["Asc"].ColorFilter = ColorFilter;
	prot									= ColorFilter.prototype;
	prot["asc_getCellColor"]				= prot.asc_getCellColor;
	prot["asc_getCColor"]					= prot.asc_getCColor;
	prot["asc_getDxf"]						= prot.asc_getDxf;
	prot["asc_setCellColor"]				= prot.asc_setCellColor;
	prot["asc_setDxf"]						= prot.asc_setDxf;
	prot["asc_setCColor"]					= prot.asc_setCColor;

	window["Asc"]["Top10"]					= window["Asc"].Top10 = Top10;
	prot									= Top10.prototype;
	prot["asc_getFilterVal"]				= prot.asc_getFilterVal;
	prot["asc_getPercent"]					= prot.asc_getPercent;
	prot["asc_getTop"]						= prot.asc_getTop;
	prot["asc_getVal"]						= prot.asc_getVal;
	prot["asc_setFilterVal"]				= prot.asc_setFilterVal;
	prot["asc_setPercent"]					= prot.asc_setPercent;
	prot["asc_setTop"]						= prot.asc_setTop;
	prot["asc_setVal"]						= prot.asc_setVal;

	window["Asc"]["asc_CPageMargins"] = window["Asc"].asc_CPageMargins = asc_CPageMargins;
	prot = asc_CPageMargins.prototype;
	prot["asc_getLeft"] = prot.asc_getLeft;
	prot["asc_getRight"] = prot.asc_getRight;
	prot["asc_getTop"] = prot.asc_getTop;
	prot["asc_getBottom"] = prot.asc_getBottom;
	prot["asc_setLeft"] = prot.asc_setLeft;
	prot["asc_setRight"] = prot.asc_setRight;
	prot["asc_setTop"] = prot.asc_setTop;
	prot["asc_setBottom"] = prot.asc_setBottom;
	prot["asc_setHeader"] = prot.asc_setHeader;
	prot["asc_setFooter"] = prot.asc_setFooter;

	window["Asc"]["asc_CPageSetup"] = window["Asc"].asc_CPageSetup = asc_CPageSetup;
	prot = asc_CPageSetup.prototype;
	prot["asc_getOrientation"] = prot.asc_getOrientation;
	prot["asc_getWidth"] = prot.asc_getWidth;
	prot["asc_getHeight"] = prot.asc_getHeight;
	prot["asc_setOrientation"] = prot.asc_setOrientation;
	prot["asc_setWidth"] = prot.asc_setWidth;
	prot["asc_setHeight"] = prot.asc_setHeight;
	prot["asc_getFitToWidth"] = prot.asc_getFitToWidth;
	prot["asc_getFitToHeight"] = prot.asc_getFitToHeight;
	prot["asc_setFitToWidth"] = prot.asc_setFitToWidth;
	prot["asc_setFitToHeight"] = prot.asc_setFitToHeight;
	prot["asc_getScale"] = prot.asc_getScale;
	prot["asc_setScale"] = prot.asc_setScale;
	prot["asc_getFirstPageNumber"] = prot.asc_getFirstPageNumber;
	prot["asc_setFirstPageNumber"] = prot.asc_setFirstPageNumber;

	window["Asc"]["asc_CPageOptions"] = window["Asc"].asc_CPageOptions = asc_CPageOptions;
	prot = asc_CPageOptions.prototype;
	prot["asc_getPageMargins"] = prot.asc_getPageMargins;
	prot["asc_getPageSetup"] = prot.asc_getPageSetup;
	prot["asc_getGridLines"] = prot.asc_getGridLines;
	prot["asc_getHeadings"] = prot.asc_getHeadings;
	prot["asc_setPageMargins"] = prot.asc_setPageMargins;
	prot["asc_setPageSetup"] = prot.asc_setPageSetup;
	prot["asc_setGridLines"] = prot.asc_setGridLines;
	prot["asc_setHeadings"] = prot.asc_setHeadings;
	prot["asc_setPrintTitlesWidth"] = prot.asc_setPrintTitlesWidth;
	prot["asc_setPrintTitlesHeight"] = prot.asc_setPrintTitlesHeight;
	prot["asc_getPrintTitlesWidth"] = prot.asc_getPrintTitlesWidth;
	prot["asc_getPrintTitlesHeight"] = prot.asc_getPrintTitlesHeight;
	prot["asc_getHeaderFooter"] = prot.asc_getHeaderFooter;
	prot["asc_setHeaderFooter"] = prot.asc_setHeaderFooter;
	prot["asc_getHorizontalCentered"] = prot.asc_getHorizontalCentered;
	prot["asc_getVerticalCentered"] = prot.asc_getVerticalCentered;
	prot["asc_setHorizontalCentered"] = prot.asc_setHorizontalCentered;
	prot["asc_setVerticalCentered"] = prot.asc_setVerticalCentered;


	window["Asc"]["CHeaderFooter"] = window["Asc"].CHeaderFooter = CHeaderFooter;
	window["Asc"]["CHeaderFooterData"] = window["Asc"].CHeaderFooterData = CHeaderFooterData;

	window["Asc"]["CSortProperties"] = window["Asc"].CSortProperties = CSortProperties;
	prot = CSortProperties.prototype;
	prot["asc_getHasHeaders"] = prot.asc_getHasHeaders;
	prot["asc_getColumnSort"] = prot.asc_getColumnSort;
	prot["asc_getLevels"] = prot.asc_getLevels;
	prot["asc_getSortList"] = prot.asc_getSortList;
	prot["asc_updateSortList"] = prot.asc_updateSortList;
	prot["asc_setHasHeaders"] = prot.asc_setHasHeaders;
	prot["asc_setColumnSort"] = prot.asc_setColumnSort;
	prot["asc_getLevelProps"] = prot.asc_getLevelProps;
	prot["asc_setLevels"] = prot.asc_setLevels;
	prot["asc_getLockChangeHeaders"] = prot.asc_getLockChangeHeaders;
	prot["asc_getLockChangeOrientation"] = prot.asc_getLockChangeOrientation;
	prot["asc_getCaseSensitive"] = prot.asc_getCaseSensitive;
	prot["asc_setCaseSensitive"] = prot.asc_setCaseSensitive;
	prot["asc_addBySortList"] = prot.asc_addBySortList;
	prot["asc_getRangeStr"] = prot.asc_getRangeStr;
	prot["asc_getSelection"] = prot.asc_getSelection;
	prot["asc_setSelection"] = prot.asc_setSelection;

	window["Asc"]["CSortPropertiesLevel"] = window["Asc"].CSortPropertiesLevel = CSortPropertiesLevel;
	prot = CSortPropertiesLevel.prototype;
	prot["asc_getIndex"] = prot.asc_getIndex;
	prot["asc_getName"] = prot.asc_getName;
	prot["asc_getSortBy"] = prot.asc_getSortBy;
	prot["asc_getDescending"] = prot.asc_getDescending;
	prot["asc_getColor"] = prot.asc_getColor;
	prot["asc_setIndex"] = prot.asc_setIndex;
	prot["asc_setName"] = prot.asc_setName;
	prot["asc_setSortBy"] = prot.asc_setSortBy;
	prot["asc_setDescending"] = prot.asc_setDescending;
	prot["asc_setColor"] = prot.asc_setColor;

	window["Asc"]["CSortLevelInfo"] = window["Asc"].CSortLevelInfo = CSortLevelInfo;
	prot = CSortLevelInfo.prototype;
	prot["asc_getColorsFill"] = prot.asc_getColorsFill;
	prot["asc_getColorsFont"] = prot.asc_getColorsFont;
	prot["asc_getIsTextData"] = prot.asc_getIsTextData;
	prot["asc_getIsDate"]     = prot.asc_getIsDate;

	window["Asc"]["CRemoveDuplicatesProps"] = window["Asc"].CRemoveDuplicatesProps = CRemoveDuplicatesProps;
	prot = CRemoveDuplicatesProps.prototype;
	prot["asc_getHasHeaders"] = prot.asc_getHasHeaders;
	prot["asc_getColumnList"] = prot.asc_getColumnList;
	prot["asc_updateColumnList"] = prot.asc_updateColumnList;
	prot["asc_setHasHeaders"] = prot.asc_setHasHeaders;
	prot["asc_getDuplicateValues"] = prot.asc_getDuplicateValues;
	prot["asc_getUniqueValues"] = prot.asc_getUniqueValues;

	window["AscCommonExcel"].CFunctionInfo = CFunctionInfo;
	prot = CFunctionInfo.prototype;
	prot["asc_getArgumentMin"] = prot.asc_getArgumentMin;
	prot["asc_getArgumentMax"] = prot.asc_getArgumentMax;
	prot["asc_getArgumentsValue"] = prot.asc_getArgumentsValue;
	prot["asc_getArgumentsType"] = prot.asc_getArgumentsType;
	prot["asc_getArgumentsResult"] = prot.asc_getArgumentsResult;
	prot["asc_getFormulaResult"] = prot.asc_getFormulaResult;
	prot["asc_getFunctionResult"] = prot.asc_getFunctionResult;
	prot["asc_getName"] = prot.asc_getName;
	prot["asc_getArguments"] = prot.asc_getArguments;
	prot["asc_setArguments"] = prot.asc_setArguments;


	window["Asc"]["asc_CExternalReference"] = window["Asc"].asc_CExternalReference = asc_CExternalReference;
	prot = asc_CExternalReference.prototype;
	prot["asc_getType"] = prot.asc_getType;
	prot["asc_getData"] = prot.asc_getData;
	prot["asc_getSource"] = prot.asc_getSource;
	prot["asc_getId"] = prot.asc_getId;
	prot["asc_isExternalLink"] = prot.isExternalLink;
	prot["asc_getPath"] = prot.asc_getPath;
	prot["asc_getLink"] = prot.asc_getLink;



	window["AscCommonExcel"].CPrintPreviewState = CPrintPreviewState;

	window["AscCommonExcel"].FromXml_ST_PatternType         = FromXml_ST_PatternType;

	window["AscCommonExcel"].ToXml_ST_PatternType           = ToXml_ST_PatternType;
	window["AscCommonExcel"].FromXml_ST_BorderStyle         = FromXml_ST_BorderStyle;
	window["AscCommonExcel"].ToXml_ST_BorderStyle           = ToXml_ST_BorderStyle;


	window["AscCommonExcel"].CT_Connection = CT_Connection;
	window["AscCommonExcel"].CT_Filter = CT_Filter;

	window["AscCommonExcel"].ExternalReference = ExternalReference;
	window["AscCommonExcel"].ExternalSheetDataSet = ExternalSheetDataSet;
	window["AscCommonExcel"].ExternalRow = ExternalRow;
	window["AscCommonExcel"].ExternalCell = ExternalCell;
	window["AscCommonExcel"].ExternalDefinedName = ExternalDefinedName;

	window['AscCommonExcel'].ToXml_ST_DataValidationOperator = ToXml_ST_DataValidationOperator;
	window['AscCommonExcel'].FromXml_ST_DataValidationOperator = FromXml_ST_DataValidationOperator;
	window['AscCommonExcel'].ToXml_CFOperatorType = ToXml_CFOperatorType;
	window['AscCommonExcel'].FromXml_CFOperatorType = FromXml_CFOperatorType;
	window['AscCommonExcel'].ToXml_ST_TimePeriod = ToXml_ST_TimePeriod;
	window['AscCommonExcel'].FromXml_ST_TimePeriod = FromXml_ST_TimePeriod;
	window["AscCommonExcel"].FromXml_ST_FilterOperator = FromXml_ST_FilterOperator;
	window["AscCommonExcel"].ToXml_ST_FilterOperator = ToXml_ST_FilterOperator;
	window["AscCommonExcel"].ToXml_ST_DynamicFilterType = ToXml_ST_DynamicFilterType;
	window["AscCommonExcel"].FromXml_ST_DynamicFilterType = FromXml_ST_DynamicFilterType;
	window["AscCommonExcel"].ToXml_ST_DateTimeGrouping = ToXml_ST_DateTimeGrouping;
	window["AscCommonExcel"].FromXml_ST_DateTimeGrouping = FromXml_ST_DateTimeGrouping;
	window["AscCommonExcel"].ToXml_ST_HorizontalAlignment = ToXml_ST_HorizontalAlignment;
	window["AscCommonExcel"].FromXml_ST_HorizontalAlignment = FromXml_ST_HorizontalAlignment;
	window["AscCommonExcel"].ToXml_ST_VerticalAlignment = ToXml_ST_VerticalAlignment;
	window["AscCommonExcel"].FromXml_ST_VerticalAlignment = FromXml_ST_VerticalAlignment;
	window["AscCommonExcel"].ToXml_ST_CfvoType = ToXml_ST_CfvoType;
	window["AscCommonExcel"].FromXml_ST_CfvoType = FromXml_ST_CfvoType;
	window["AscCommonExcel"].ToXml_ST_IconSetType = ToXml_ST_IconSetType;
	window["AscCommonExcel"].FromXml_ST_IconSetType = FromXml_ST_IconSetType;
	window["AscCommonExcel"].ToXml_ST_CfType = ToXml_ST_CfType;
	window["AscCommonExcel"].FromXml_ST_CfType = FromXml_ST_CfType;
	window["AscCommonExcel"].ToXml_ST_DataValidationErrorStyle = ToXml_ST_DataValidationErrorStyle;
	window["AscCommonExcel"].FromXml_ST_DataValidationErrorStyle = FromXml_ST_DataValidationErrorStyle;
	window["AscCommonExcel"].ToXml_ST_DataValidationType = ToXml_ST_DataValidationType;
	window["AscCommonExcel"].FromXml_ST_DataValidationType = FromXml_ST_DataValidationType;
	window["AscCommonExcel"].ToXml_ST_DataValidationImeMode = ToXml_ST_DataValidationImeMode;
	window["AscCommonExcel"].FromXml_ST_DataValidationImeMode = FromXml_ST_DataValidationImeMode;
	window["AscCommonExcel"].ToXml_ST_GradientType = ToXml_ST_GradientType;
	window["AscCommonExcel"].FromXml_ST_GradientType = FromXml_ST_GradientType;
	window["AscCommonExcel"].FromXml_ST_FontScheme = FromXml_ST_FontScheme;
	window["AscCommonExcel"].ToXml_ST_FontScheme = ToXml_ST_FontScheme;
	window["AscCommonExcel"].FromXml_ST_PageOrder = FromXml_ST_PageOrder;
	window["AscCommonExcel"].ToXml_ST_PageOrder = ToXml_ST_PageOrder;
	window["AscCommonExcel"].FromXml_ST_PageOrientation = FromXml_ST_PageOrientation;
	window["AscCommonExcel"].ToXml_ST_PageOrientation = ToXml_ST_PageOrientation;
	window["AscCommonExcel"].FromXml_ST_CellComments = FromXml_ST_CellComments;
	window["AscCommonExcel"].ToXml_ST_CellComments = ToXml_ST_CellComments;
	window["AscCommonExcel"].FromXml_ST_PrintError = FromXml_ST_PrintError;
	window["AscCommonExcel"].ToXml_ST_PrintError = ToXml_ST_PrintError;
	window["AscCommonExcel"].FromXml_ST_SortBy = FromXml_ST_SortBy;
	window["AscCommonExcel"].ToXml_ST_SortBy = ToXml_ST_SortBy;
	window["AscCommonExcel"].FromXml_ST_TabularSlicerCacheSortOrder = FromXml_ST_TabularSlicerCacheSortOrder;
	window["AscCommonExcel"].ToXml_ST_TabularSlicerCacheSortOrder = ToXml_ST_TabularSlicerCacheSortOrder;
	window["AscCommonExcel"].FromXml_ST_SlicerCacheCrossFilter = FromXml_ST_SlicerCacheCrossFilter;
	window["AscCommonExcel"].ToXml_ST_SlicerCacheCrossFilter = ToXml_ST_SlicerCacheCrossFilter;
	window["AscCommonExcel"].FromXml_ST_SortMethod = FromXml_ST_SortMethod;
	window["AscCommonExcel"].ToXml_ST_SortMethod = ToXml_ST_SortMethod;
	window["AscCommonExcel"].FromXML_ST_DispBlanksAs = FromXML_ST_DispBlanksAs;
	window["AscCommonExcel"].ToXML_ST_DispBlanksAs = ToXML_ST_DispBlanksAs;
	window["AscCommonExcel"].FromXml_ST_SparklineType = FromXml_ST_SparklineType;
	window["AscCommonExcel"].ToXml_ST_SparklineType = ToXml_ST_SparklineType;
	window["AscCommonExcel"].FromXml_ST_UnderlineValues = FromXml_ST_UnderlineValues;
	window["AscCommonExcel"].ToXml_ST_UnderlineValues = ToXml_ST_UnderlineValues;
	window["AscCommonExcel"].FromXml_ST_SparklineAxisMinMax = FromXml_ST_SparklineAxisMinMax;
	window["AscCommonExcel"].ToXml_ST_SparklineAxisMinMax = ToXml_ST_SparklineAxisMinMax;
	window["AscCommonExcel"].FromXml_ST_CellFormulaType = FromXml_ST_CellFormulaType;
	window["AscCommonExcel"].ToXml_ST_CellFormulaType = ToXml_ST_CellFormulaType;
	window["AscCommonExcel"].FromXml_ST_TableStyleType = FromXml_ST_TableStyleType;
	window["AscCommonExcel"].ToXml_ST_TableStyleType = ToXml_ST_TableStyleType;
	window["AscCommonExcel"].FromXml_ST_DataBarDirection = FromXml_ST_DataBarDirection;
	window["AscCommonExcel"].ToXml_ST_DataBarDirection = ToXml_ST_DataBarDirection;
	window["AscCommonExcel"].FromXml_ST_DataBarAxisPosition = FromXml_ST_DataBarAxisPosition;
	window["AscCommonExcel"].ToXml_ST_DataBarAxisPosition = ToXml_ST_DataBarAxisPosition;
	window["AscCommonExcel"].FromXml_ST_CellValueType = FromXml_ST_CellValueType;
	window["AscCommonExcel"].ToXml_ST_CellValueType = ToXml_ST_CellValueType;
	window["AscCommonExcel"].FromXml_ST_SheetViewType = FromXml_ST_SheetViewType;
	window["AscCommonExcel"].ToXml_ST_SheetViewType = ToXml_ST_SheetViewType;

	window["AscCommonExcel"].CCellWatch = CCellWatch;
	prot = CCellWatch.prototype;
	prot["asc_getWorkbook"] = prot.asc_getWorkbook;
	prot["asc_getSheet"] = prot.asc_getSheet;
	prot["asc_getName"] = prot.asc_getName;
	prot["asc_getCell"] = prot.asc_getCell;
	prot["asc_getValue"] = prot.asc_getValue;
	prot["asc_getFormula"] = prot.asc_getFormula;

	window["AscCommonExcel"].CRowColBreaks = CRowColBreaks;
	window["AscCommonExcel"].CBreak = CBreak;

	window["Asc"]["asc_CSeriesSettings"] = window["Asc"].asc_CSeriesSettings = asc_CSeriesSettings;
	prot = asc_CSeriesSettings.prototype;
	prot["asc_getSeriesIn"] = prot.asc_getSeriesIn;
	prot["asc_getType"] = prot.asc_getType;
	prot["asc_getDateUnit"] = prot.asc_getDateUnit;
	prot["asc_getTrend"] = prot.asc_getTrend;
	prot["asc_getStepValue"] = prot.asc_getStepValue;
	prot["asc_getStopValue"] = prot.asc_getStopValue;
	prot["asc_getContextMenuAllowedProps"] = prot.asc_getContextMenuAllowedProps;
	prot["asc_getContextMenuChosenProperty"] = prot.asc_getContextMenuChosenProperty;
	prot["asc_getToolbarMenuAllowedProps"] = prot.asc_getToolbarMenuAllowedProps;

	prot["asc_setSeriesIn"] = prot.asc_setSeriesIn;
	prot["asc_setType"] = prot.asc_setType;
	prot["asc_setDateUnit"] = prot.asc_setDateUnit;
	prot["asc_setTrend"] = prot.asc_setTrend;
	prot["asc_setStepValue"] = prot.asc_setStepValue;
	prot["asc_setStopValue"] = prot.asc_setStopValue;
	prot["asc_isValidStepValue"] = prot.asc_isValidStepValue;
	prot["asc_isValidStopValue"] = prot.asc_isValidStopValue;
	prot["asc_setContextMenuAllowedProps"] = prot.asc_setContextMenuAllowedProps;
	prot["asc_setContextMenuChosenProperty"] = prot.asc_setContextMenuChosenProperty;
	prot["asc_setToolbarMenuAllowedProps"] = prot.asc_setToolbarMenuAllowedProps;


	window["AscCommonExcel"].CTimeline = CTimeline;
	window["AscCommonExcel"].CTimelineCacheDefinition = CTimelineCacheDefinition;
	window["AscCommonExcel"].CTimelineCachePivotTable = CTimelineCachePivotTable;
	window["AscCommonExcel"].CTimelineState = CTimelineState;
	window["AscCommonExcel"].CTimelineRange = CTimelineRange;
	window["AscCommonExcel"].CTimelineStyles = CTimelineStyles;
	window["AscCommonExcel"].CTimelineStyle = CTimelineStyle;
	window["AscCommonExcel"].CTimelineStyleElement = CTimelineStyleElement;
	window["AscCommonExcel"].CTimelinePivotFilter = CTimelinePivotFilter;

	window["AscCommonExcel"].CCalcPr = CCalcPr;

	window["Asc"]["asc_CCalcSettings"] = window["Asc"].asc_CCalcSettings = asc_CCalcSettings;
	prot = asc_CCalcSettings.prototype;
	prot["asc_getIterativeCalc"] = prot.asc_getIterativeCalc;
	prot["asc_getMaxIterations"] = prot.asc_getMaxIterations;
	prot["asc_getMaxChange"] = prot.asc_getMaxChange;

	prot["asc_setIterativeCalc"] = prot.asc_setIterativeCalc;
	prot["asc_setMaxIterations"] = prot.asc_setMaxIterations;
	prot["asc_setMaxChange"] = prot.asc_setMaxChange;
	prot["asc_initSettings"] = prot.asc_initSettings;

	window["AscCommonExcel"].CWorkbookPr = CWorkbookPr;
	
	window["AscCommonExcel"].CMetadata = CMetadata;
	window["AscCommonExcel"].CMetadataType = CMetadataType;
	window["AscCommonExcel"].CMetadataString = CMetadataString;
	window["AscCommonExcel"].CFutureMetadata = CFutureMetadata;
	window["AscCommonExcel"].CMdx = CMdx;
	window["AscCommonExcel"].CMdxTuple = CMdxTuple;
	window["AscCommonExcel"].CMdxSet = CMdxSet;
	window["AscCommonExcel"].CMetadataStringIndex = CMetadataStringIndex;
	window["AscCommonExcel"].CMdxMemeberProp = CMdxMemeberProp;
	window["AscCommonExcel"].CMdxKPI = CMdxKPI;
	window["AscCommonExcel"].CMetadataBlock = CMetadataBlock;
	window["AscCommonExcel"].CMetadataRecord = CMetadataRecord;
	window["AscCommonExcel"].CFutureMetadataBlock = CFutureMetadataBlock;
	window["AscCommonExcel"].CMetadataBlockExt = CMetadataBlockExt;
	window["AscCommonExcel"].CDynamicArrayProperties = CDynamicArrayProperties;
	window["AscCommonExcel"].CRichValueBlock = CRichValueBlock;
	window["AscCommonExcel"].CCustomFunctionEngine = CCustomFunctionEngine;

	window["AscCommonExcel"].CCustomFunctionInfo = CCustomFunctionInfo;
	prot = CCustomFunctionInfo.prototype;
	prot["asc_getDescription"] = prot.asc_getDescription;
	prot["asc_getArg"] = prot.asc_getArg;

	window["AscCommonExcel"].CCustomFunctionArgInfo = CCustomFunctionArgInfo;
	prot = CCustomFunctionArgInfo.prototype;
	prot["asc_getName"] = prot.asc_getName;
	prot["asc_getIsOptional"] = prot.asc_getIsOptional;


	window["AscCommonExcel"].CWorkbookInfo = CWorkbookInfo;
	prot = CWorkbookInfo.prototype;
	prot["asc_getName"] = prot.asc_getName;
	prot["asc_getId"] = prot.asc_getId;
	prot["asc_getSheets"] = prot.asc_getSheets;

	window["AscCommonExcel"].CWorksheetInfo = CWorksheetInfo;
	prot = CWorksheetInfo.prototype;
	prot["asc_getName"] = prot.asc_getName;
	prot["asc_getIndex"] = prot.asc_getIndex;




})(window);
