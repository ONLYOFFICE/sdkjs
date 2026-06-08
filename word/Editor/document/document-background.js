/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
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
