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


(function()
{
	/**
	 * @constructor
	 */
	function DocumentPageColumn()
	{
		this.Bounds = new CDocumentBounds(0, 0, 0, 0);
		this.Pos    = 0;
		this.EndPos = -1;
		this.Empty  = true;
		
		this.X      = 0;
		this.Y      = 0;
		this.XLimit = 0;
		this.YLimit = 0;
		
		this.SpaceBefore = 0;
		this.SpaceAfter  = 0;
	}
	DocumentPageColumn.prototype.Copy = function()
	{
		var NewColumn = new DocumentPageColumn();
		
		NewColumn.Bounds.CopyFrom(this.Bounds);
		NewColumn.Pos    = this.Pos;
		NewColumn.EndPos = this.EndPos;
		NewColumn.X      = this.X;
		NewColumn.Y      = this.Y;
		NewColumn.XLimit = this.XLimit;
		NewColumn.YLimit = this.YLimit;
		
		return NewColumn;
	};
	DocumentPageColumn.prototype.Shift = function(Dx, Dy)
	{
		this.X      += Dx;
		this.XLimit += Dx;
		this.Y      += Dy;
		this.YLimit += Dy;
		
		this.Bounds.Shift(Dx, Dy);
	};
	DocumentPageColumn.prototype.Reset = function()
	{
		this.Bounds.Reset();
		this.Pos    = 0;
		this.EndPos = -1;
		this.Empty  = true;
		
		this.X      = 0;
		this.Y      = 0;
		this.XLimit = 0;
		this.YLimit = 0;
	};
	DocumentPageColumn.prototype.IsEmpty = function()
	{
		return this.Empty;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.DocumentPageColumn = DocumentPageColumn;
	
})();
