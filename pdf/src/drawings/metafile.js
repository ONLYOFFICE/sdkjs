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

(function(window, undefined)
{
	function CPdfTextMetafile() {
		AscCommon.CGraphicsBase.apply(this, arguments);
		
		this.m_oTransform = new AscCommon.CMatrix();
		this.Memory       = null;
		this.m_oGrFonts   = new AscCommon.CGrRFonts();
		
		this.fontFile = null;
		this.fontInfo = {
			FontFamily : {Name : "", Index : -1},
			FontSize : 0,
			Italic : false,
			Bold : false
		};
		
		this._lineStartPos = 0;
		this._newLine = true;
		this._charCountInLine = 0;
		this._startX = 0;
		this._endX   = 0;
		this._prevX  = 0;
		
		this._ascent = 0;
		this._descent = 0;
		this._curAscent = 0;
		this._curDescent = 0;
	}

	CPdfTextMetafile.prototype = Object.create(AscCommon.CGraphicsBase.prototype);
	CPdfTextMetafile.prototype.constructor = CPdfTextMetafile;

	CPdfTextMetafile.prototype.transform = function(sx, shy, shx, sy, tx, ty) {
		if (this.m_oTransform.sx != sx || this.m_oTransform.shx != shx || this.m_oTransform.shy != shy ||
			this.m_oTransform.sy != sy || this.m_oTransform.tx != tx || this.m_oTransform.ty != ty) {

			this.m_oTransform.sx  = sx;
			this.m_oTransform.shx = shx;
			this.m_oTransform.shy = shy;
			this.m_oTransform.sy  = sy;
			this.m_oTransform.tx  = tx;
			this.m_oTransform.ty  = ty;
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
		
		if (this.fontInfo.FontFamily.Name === font.FontFamily.Name
			&& this.fontInfo.FontSize === font.FontSize
			&& this.fontInfo.Bold === font.Bold
			&& this.fontInfo.Italic === font.Italic) {
			return;
		}
		
		this.fontInfo.FontFamily.Name = font.FontFamily.Name;
		this.fontInfo.FontSize = font.FontSize;
		this.fontInfo.Bold = font.Bold;
		this.fontInfo.Italic = font.Italic;
		
		let style = 0;
		if (font.Italic)
			style |= AscFonts.FontStyle.FontStyleItalic;
		
		if (font.Bold)
			style |= AscFonts.FontStyle.FontStyleBold;
		
		this.fontFile = g_oTextMeasurer.SetFontInternal(this.fontInfo.FontFamily.Name, this.fontInfo.FontSize, style);
		
		this._curAscent  = g_oTextMeasurer.GetAscender();
		this._curDescent = Math.abs(g_oTextMeasurer.GetDescender());
	};

	CPdfTextMetafile.prototype.FillTextCode = function(x, y, codePoint) {
		let advance = g_oTextMeasurer.MeasureCode(codePoint);
		this.tg(0, x, y, [codePoint], advance.Width, advance.Height);
	};
	CPdfTextMetafile.prototype.tg = function(gid, x, y, codepoints, advX, advY) {
		let unicode = (Array.isArray(codepoints)) ? codepoints[0] : codepoints;
		
		if (this._newLine || (this._x - x) > 0.001)
			this.startLine(x, y);
		
		this.updateFontMetrics();
		
		if (this._charCountInLine)
			this.Memory.WriteLong((x - this._prevX) * 10000);

		this.Memory.WriteLong(unicode);
		this.Memory.WriteLong(advX * 10000);
		
		this._curX  = x;
		this._prevX = x;
		this._endX  = x + advX;
		
		this._charCountInLine++;
	};

	CPdfTextMetafile.prototype.endLine = function() {
		if (0 === this._charCountInLine)
			return;
		
		let curPos = this.Memory.GetCurPosition();
		this.Memory.Seek(this._lineStartPos);
		this.Memory.WriteLong(this._ascent * 10000);
		this.Memory.WriteLong(this._descent * 10000);
		this.Memory.WriteLong((this._endX - this._startX) * 10000);
		this.Memory.WriteLong(this._charCountInLine);
		this.Memory.Seek(curPos);
		
		this._newLine = true;
		this._charCountInLine = 0;
	};
	
	CPdfTextMetafile.prototype.startLine = function(x, y) {
		
		this.endLine();
		
		this._startX = x;
		
		let _x = this.m_oTransform.TransformPointX(x, y);
		let _y = this.m_oTransform.TransformPointY(x, y);
		
		this.Memory.WriteLong(_x * 10000);
		this.Memory.WriteLong(_y * 10000);
		
		if (this.m_oTransform.IsIdentity2()) {
			this.Memory.WriteByte(0);
		}
		else {
			this.Memory.WriteByte(1);
			
			let angle = this.m_oTransform.GetRotation();
			let rad = (angle * Math.PI) / 180;
			
			let ex = Math.cos(rad);
			let ey = Math.sin(rad);
			
			this.Memory.WriteLong(ex * 10000);
			this.Memory.WriteLong(ey * 10000);
		}
		
		this._lineStartPos = this.Memory.GetCurPosition();
		this.Memory.Skip(4); // ascent
		this.Memory.Skip(4); // descent
		this.Memory.Skip(4); // total width
		this.Memory.Skip(4); // char count
		
		this._ascent = 0;
		this._descent = 0;
		
		this._charCountInLine = 0;
		this._newLine = false;
	};
	CPdfTextMetafile.prototype.Start_Command = function(commandId) {
		if (AscFormat.DRAW_COMMAND_LINE_RANGE === commandId) {
			this._newLine = true;
			this._charCountInLine = 0;
		}
	};
	CPdfTextMetafile.prototype.End_Command = function(commandId) {
		if (AscFormat.DRAW_COMMAND_LINE_RANGE === commandId) {
			this.endLine();
		}
	};
	CPdfTextMetafile.prototype.updateFontMetrics = function(commandId) {
		this._ascent = Math.max(this._ascent, this._curAscent);
		this._descent = Math.max(this._descent, this._curDescent);
	};

	window['AscPDF'].CPdfTextMetafile = CPdfTextMetafile;

})(window);
