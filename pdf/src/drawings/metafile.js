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
		AscCommon.CMetafile.apply(this, arguments);
		
		// Internal state for line tracking
		this.m_currentLineStartPos = -1; 
		this.m_currentLineWidth = 0;
		this.m_charCountInLine = 0;
		this.m_curCharWidth = 0;
	}

	CPdfTextMetafile.prototype = Object.create(AscCommon.CMetafile.prototype);
	CPdfTextMetafile.prototype.constructor = CPdfTextMetafile;

	// Override all methods of the parent class
	for (let prop in AscCommon.CMetafile.prototype) {
		if (typeof AscCommon.CMetafile.prototype[prop] === 'function') {
			CPdfTextMetafile.prototype[prop] = function() { return; };
		}
	}

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
		let _lastFont = this.m_oFontSlotFont;
		_lastFont.Name = name;
		_lastFont.Size = size;
		_lastFont.Bold = (style & AscFonts.FontStyle.FontStyleBold) ? true : false;
		_lastFont.Italic = (style & AscFonts.FontStyle.FontStyleItalic) ? true : false;

		this.m_oFontTmp.FontFamily.Name = _lastFont.Name;
		this.m_oFontTmp.Bold = _lastFont.Bold;
		this.m_oFontTmp.Italic = _lastFont.Italic;
		this.m_oFontTmp.FontSize = _lastFont.Size;
		this.SetFont(this.m_oFontTmp);
	};

	CPdfTextMetafile.prototype.SetFont = function(font, isFromPicker) {
		if (!font) return;

		// Calculate the style for the metric
		let style = 0;
		if (font.Italic) style += 2;
		if (font.Bold) style += 1;

		// Check whether the font has changed relative to the current state of m_oFont
        // CMetafile usually already contains m_oFont; we use it for comparison
		if (this.m_oFont.FontSize !== font.FontSize) {
			this.m_oFont.Name = font.FontFamily.Name;
			this.m_oFont.FontSize = font.FontSize;
			this.m_oFont.Style = style;
			
			// 1. If there was an active line, we close it (and record the final width)
			this.finishCurrentLine();

			// 2. Update font metrics using the global meter
			g_oTextMeasurer.SetFontInternal(font.FontFamily.Name, font.FontSize, style);
			this.m_currentAscent = g_oTextMeasurer.GetAscender();
			this.m_currentDescent = Math.abs(g_oTextMeasurer.GetDescender());
			
			// 3. Reset the line start flag so that tg starts a new entry
			this.m_currentLineStartPos = -1;
		}
	};

	CPdfTextMetafile.prototype.FillTextCode = function(x, y, code) {
		let _code = code;
		if (null != this.LastFontOriginInfo.Replace)
			_code = g_fontApplication.GetReplaceGlyph(_code, this.LastFontOriginInfo.Replace);

		if (this.FontPicker)
			this.FontPicker.FillTextCode(_code);

		this.tg(undefined, x, y, [code]);
	};
	CPdfTextMetafile.prototype.tg = function(gid, x, y, codepoints) {
		let unicode = (Array.isArray(codepoints)) ? codepoints[0] : codepoints;
		let glyphWidth = this.m_curCharWidth;

		let _x = this.m_oTransform.TransformPointX(x,y);
		let _y = this.m_oTransform.TransformPointY(x,y);

		if (x <= this.m_LastX || y !== this.m_LastY) {
			this.finishCurrentLine();
		}

		// If the line hasn't started yet, write the line header
		if (this.m_currentLineStartPos === -1) {
			this.Memory.WriteLong(_x * 10000);
			this.Memory.WriteLong(_y * 10000);
			
			let isComplex = !this.m_oTransform.IsIdentity2();
			if (!isComplex) {
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

			this.Memory.WriteLong(this.m_currentAscent * 10000);
			this.Memory.WriteLong(this.m_currentDescent * 10000);

			// 1. Save the position for the final width (dWidthLine)
			this.m_currentLineStartPos = this.Memory.GetCurPosition();
			this.Memory.Skip(4);

			// 2. Save the position for nCount (number of characters)
			this.m_countPosition = this.Memory.GetCurPosition();
			this.Memory.Skip(4);

			this.m_currentLineWidth = 0;
			this.m_charCountInLine = 0;
		}
		else if (this.m_charCountInLine > 0) {
			if (this.m_LastX !== undefined) {
				let spaceW = x - this.m_LastX - this.m_LastCharW;

				// If the jump is longer than the width of the previous character, we account for the width of that jump
				if (this.m_LastX + this.m_LastCharW !== x) {
					let curCharWidth = this.m_curCharWidth;
					this.m_curCharWidth = spaceW;

					this.tg(undefined, this.m_LastX + this.m_LastCharW, y, [65535]);
					this.m_curCharWidth = curCharWidth;
				}
			}
			
			this.Memory.WriteLong((x - this.m_LastX) * 10000);
		}

		this.Memory.WriteLong(unicode);
		this.Memory.WriteLong(glyphWidth * 10000);
		
		this.m_currentLineWidth += glyphWidth;
		this.m_charCountInLine++;

		this.m_LastCharW = glyphWidth;
		this.m_LastX = x;
		this.m_LastY = y;
	};

	// Method for finalizing a line
	CPdfTextMetafile.prototype.finishCurrentLine = function() {
		if (this.m_currentLineStartPos !== -1) {
			let totalWidth = this.m_currentLineWidth;
			let endPosition = this.Memory.GetCurPosition();

			// Record the width
			this.Memory.Seek(this.m_currentLineStartPos);
			this.Memory.WriteLong(totalWidth * 10000);

			// Record the number of characters
			this.Memory.Seek(this.m_countPosition);
			this.Memory.WriteLong(this.m_charCountInLine);

			this.Memory.Seek(endPosition);
			
			// Reset the flag so that the next iteration of tg starts a new structure
			this.m_currentLineStartPos = -1;
		}
	};

	window['AscPDF'].CPdfTextMetafile = CPdfTextMetafile;

})(window);
