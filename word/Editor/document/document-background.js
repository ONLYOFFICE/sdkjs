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

(function(window)
{
	/**
	 *
	 * @constructor
	 */
	function DocumentBackground(color, unifill, shape)
	{
		this.Color   = color ? color : null;
		this.Unifill = unifill ? unifill : null;
		this.shape   = shape ? shape : null;
	}
	DocumentBackground.prototype.copy = function()
	{
		return new DocumentBackground(this.Color, this.Unifill, this.shape);
	};
	DocumentBackground.prototype.draw = function(graphics, sectPr, theme, colorMap)
	{
		let brush = this._getBrush();
		if (!brush || !brush.isVisible())
			return;
		
		let h = sectPr.GetPageHeight();
		let w = sectPr.GetPageWidth();

		graphics.StartDrawShape();
		
		let shapeDrawer = new AscCommon.CShapeDrawer();
		brush.check(theme, colorMap);
		shapeDrawer.fromShape2(new AscFormat.ObjectToDraw(brush, null, w, h, null, null), graphics, null);
		
		if (brush.isSolidFill())
		{
			let RGBA = brush.getRGBAColor();
			graphics.setEndGlobalAlphaColor(RGBA.R, RGBA.G, RGBA.B);
		}
		shapeDrawer.draw(null);

		graphics.EndDrawShape();
	};
	DocumentBackground.prototype._getBrush = function()
	{
		let brush = null;
		if (this.shape)
			brush = this.shape.brush;
		else if (this.Unifill)
			brush = this.Unifill;
		else if (this.Color)
			brush = AscFormat.CreateSolidFillRGB(this.Color.r, this.Color.g, this.Color.b);
		
		return brush;
	};
	DocumentBackground.prototype.writeToBinary = function(writer)
	{
		let startPos = writer.GetCurPosition();
		writer.Skip(4);
		
		let flags = 0;
		if (this.Color)
		{
			flags |= 1;
			this.Color.Write_ToBinary(writer);
		}
		
		if (this.Unifill)
		{
			flags |= 2;
			this.Unifill.Write_ToBinary(writer);
		}
		
		if (this.shape)
		{
			flags |= 4;
			writer.WriteString2(this.shape.GetId());
		}
		
		let endPos = writer.GetCurPosition();
		writer.Seek(startPos);
		writer.WriteLong(flags);
		writer.Seek(endPos);
	};
	DocumentBackground.prototype.readFromBinary = function(reader)
	{
		let flags = reader.GetLong();
		
		if (flags & 1)
		{
			this.Color = new CDocumentColor(0, 0, 0);
			this.Color.Read_FromBinary(reader);
		}
		
		if (flags & 2)
		{
			this.Unifill = new AscFormat.CUniFill();
			this.Unifill.Read_FromBinary(reader);
		}
		
		if (flags & 4)
		{
			let shapeId = reader.GetString2();
			this.shape = AscCommon.g_oTableId.GetById(shapeId);
		}
	};
	DocumentBackground.prototype.getAscColor = function()
	{
		let Unifill = this.Unifill;
		if (Unifill && Unifill.fill && Unifill.fill.color)
			return AscCommon.CreateAscColor(Unifill.fill.color);
		else if (this.Color)
			return AscCommon.CreateAscColorCustom(this.Color.r, this.Color.g, this.Color.b, false);
		
		return null;
	};
	DocumentBackground.prototype.isDefault = function()
	{
		if (this.Unifill || this.shape)
			return false;
		
		return (!this.Color || this.Color.IsEqualRGB({r : 255, g : 255, b : 255}));
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].DocumentBackground = DocumentBackground;
	
})(window);
