/*
 * (c) Copyright Ascensio System SIA 2010-2026
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

(function(window, undefined)
{
	let memory = null;
	function getMemory()
	{
		if (!memory) {
			let memoryInitSize = 1024 * 10; // 10Kb
			memory = new AscCommon.CMemory(true);
			memory.Init(memoryInitSize);
		}
		memory.Seek(0);
		return memory;
	}
	
	/**
	 * @constructor
	 */
	function CPdfTextMetafile() {
		AscCommon.CGraphicsBase.apply(this, arguments);
		
		this._memory = getMemory();
		
		this._fontInfo = {
			FontFamily : {Name : "", Index : -1},
			FontSize : 0,
			Italic : false,
			Bold : false
		};
		
		this._lineStartPos = 0;
		this._newLine = true;
		this._charCountInLine = 0;
		this._startX = 0;
		this._x = 0;
		this._prevX = 0;
		this._y = 0;
		
		this._ascent = 0;
		this._descent = 0;
		this._curAscent = 0;
		this._curDescent = 0;
		
		this._transform = new AscCommon.CMatrix();
		this._states = []; // we need only transform here
	}

	CPdfTextMetafile.prototype = Object.create(AscCommon.CGraphicsBase.prototype);
	CPdfTextMetafile.prototype.constructor = CPdfTextMetafile;
	
	CPdfTextMetafile.prototype.transform = function(sx, shy, shx, sy, tx, ty) {
		if (this._transform.sx != sx || this._transform.shx != shx || this._transform.shy != shy ||
			this._transform.sy != sy || this._transform.tx != tx || this._transform.ty != ty) {

			this._transform.sx  = sx;
			this._transform.shx = shx;
			this._transform.shy = shy;
			this._transform.sy  = sy;
			this._transform.tx  = tx;
			this._transform.ty  = ty;
		}
	};
	CPdfTextMetafile.prototype.transform3 = function(m) {
		this.transform(m.sx, m.shy, m.shx, m.sy, m.tx, m.ty);
	};
	CPdfTextMetafile.prototype.reset = function() {
		this.transform(1, 0, 0, 1, 0, 0);
	};
	CPdfTextMetafile.prototype.GetTransform = function() {
		return this._transform;
	};
	CPdfTextMetafile.prototype.SaveGrState = function() {
		this._states.push(this._transform.CreateDublicate());
		
	};
	CPdfTextMetafile.prototype.RestoreGrState = function() {
		if (this._states.length) {
			this._transform = this._states[this._states.length - 1];
			this._states.length--;
		}
	};
	CPdfTextMetafile.prototype.SetFontInternal = function(name, size, style) {
		this.SetFont({
			FontFamily : {Name : name, Index : -1},
			FontSize : size,
			Italic : !!(style & AscFonts.FontStyle.FontStyleItalic),
			Bold : !!(style & AscFonts.FontStyle.FontStyleBold)
		});
	};

	CPdfTextMetafile.prototype.SetFont = function(font) {
		if (!font) return;
		
		if (this._fontInfo.FontFamily.Name === font.FontFamily.Name
			&& this._fontInfo.FontSize === font.FontSize
			&& this._fontInfo.Bold === font.Bold
			&& this._fontInfo.Italic === font.Italic) {
			return;
		}
		
		this._fontInfo.FontFamily.Name = font.FontFamily.Name;
		this._fontInfo.FontSize = font.FontSize;
		this._fontInfo.Bold = font.Bold;
		this._fontInfo.Italic = font.Italic;
		
		let style = 0;
		if (font.Italic)
			style |= AscFonts.FontStyle.FontStyleItalic;
		
		if (font.Bold)
			style |= AscFonts.FontStyle.FontStyleBold;
		
		g_oTextMeasurer.SetFontInternal(this._fontInfo.FontFamily.Name, this._fontInfo.FontSize, style);
		this._curAscent  = g_oTextMeasurer.GetAscender();
		this._curDescent = Math.abs(g_oTextMeasurer.GetDescender());
	};

	CPdfTextMetafile.prototype.FillTextCode = function(x, y, codePoint) {
		let advance = g_oTextMeasurer.MeasureCode(codePoint);
		this.tg(0, x, y, [codePoint], advance.Width, advance.Height);
	};
	CPdfTextMetafile.prototype.tg = function(gid, x, y, codepoints, advX, advY) {
		let unicode = (Array.isArray(codepoints)) ? codepoints[0] : codepoints;
		
		if (this._newLine || (this._x - x) > 0.001 || Math.abs(this._y - y) > 0.001)
			this.startLine(x, y);
		
		this.updateFontMetrics();
		
		if (this._charCountInLine)
		{
			if ((x - this._x) > 0.1)
			{
				this._memory.WriteLong((this._x - this._prevX) * 10000);
				this._memory.WriteLong(0xFFFF);
				this._memory.WriteLong((x - this._x) * 10000);
				this._charCountInLine++;
				this._prevX = this._x;
				this._x = x;
			}
			
			this._memory.WriteLong((x - this._prevX) * 10000);
		}
		
		this._memory.WriteLong(unicode);
		this._memory.WriteLong(advX * 10000);
		
		this._prevX = x;
		this._x     = x + advX;
		this._y     = y;
		
		this._charCountInLine++;
	};

	CPdfTextMetafile.prototype.endLine = function() {
		if (0 === this._charCountInLine)
			return;
		
		let curPos = this._memory.GetCurPosition();
		this._memory.Seek(this._lineStartPos);
		this._memory.WriteLong(this._ascent * 10000);
		this._memory.WriteLong(this._descent * 10000);
		this._memory.WriteLong((this._x - this._startX) * 10000);
		this._memory.WriteLong(this._charCountInLine);
		this._memory.Seek(curPos);
		
		this._newLine = true;
		this._charCountInLine = 0;
	};
	
	CPdfTextMetafile.prototype.startLine = function(x, y) {
		
		this.endLine();
		
		this._startX = x;
		
		let _x = this._transform.TransformPointX(x, y);
		let _y = this._transform.TransformPointY(x, y);
		
		this._memory.WriteLong(_x * 10000);
		this._memory.WriteLong(_y * 10000);
		
		if (this._transform.IsIdentity2()) {
			this._memory.WriteByte(0);
		}
		else {
			this._memory.WriteByte(1);
			
			let angle = this._transform.GetRotation();
			let rad = (angle * Math.PI) / 180;
			
			let ex = Math.cos(rad);
			let ey = Math.sin(rad);
			
			this._memory.WriteLong(ex * 10000);
			this._memory.WriteLong(ey * 10000);
		}
		
		this._lineStartPos = this._memory.GetCurPosition();
		this._memory.Skip(4); // ascent
		this._memory.Skip(4); // descent
		this._memory.Skip(4); // total width
		this._memory.Skip(4); // char count
		
		this._ascent = 0;
		this._descent = 0;
		
		this._charCountInLine = 0;
		this._newLine = false;
	};
	CPdfTextMetafile.prototype.updateFontMetrics = function() {
		this._ascent = Math.max(this._ascent, this._curAscent);
		this._descent = Math.max(this._descent, this._curDescent);
	};
	CPdfTextMetafile.prototype.getData = function() {
		return this._memory.data.subarray(0, this._memory.pos);
	};

	window['AscPDF'].CPdfTextMetafile = CPdfTextMetafile;

})(window);
