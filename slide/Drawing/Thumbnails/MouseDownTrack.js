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

function CMouseDownTrack(Thumbnails)
{
	this.Thumbnails = Thumbnails;
	this.SetDefault();
}
CMouseDownTrack.prototype.Check = function()
{
	if(this.Started)
	{
		let aThPages = this.Thumbnails.m_arrPages;
		if(!aThPages[this.Page])
		{
			this.Reset();
		}
	}
};
CMouseDownTrack.prototype.SetDefault = function()
{
	this.Started = false;
	this.StartedSimple = true;
	this.Page = -1;
	this.X = -1;
	this.Y = -1;
	this.Position = -1;
};
CMouseDownTrack.prototype.Reset = function()
{
	this.SetDefault();
};
CMouseDownTrack.prototype.Start = function(Page, X, Y)
{
	this.Started = true;
	this.StartedSimple = true;
	this.Page = Page;
	this.X = X;
	this.Y = Y;
};
CMouseDownTrack.prototype.IsStarted = function()
{
	this.Check();
	return this.Started;
};
CMouseDownTrack.prototype.IsSimple = function()
{
	this.Check();
	return this.StartedSimple;
};
CMouseDownTrack.prototype.IsDragged = function()
{
	return this.IsStarted() && !this.IsSimple() && this.GetPosition() !== -1;
};
CMouseDownTrack.prototype.ResetSimple = function(Position)
{
	this.StartedSimple = false;
	this.Position = Position;
};
CMouseDownTrack.prototype.SetPosition = function(Position)
{
	this.Position = Position;
};
CMouseDownTrack.prototype.GetPosition = function()
{
	this.Check();
	return this.Position;
};
CMouseDownTrack.prototype.GetPage = function()
{
	this.Check();
	return this.Page;
};
CMouseDownTrack.prototype.GetX = function()
{
	this.Check();
	return this.X;
};
CMouseDownTrack.prototype.GetY = function()
{
	this.Check();
	return this.Y;
};

CMouseDownTrack.prototype.IsMoved = function(X, Y)
{
	this.Check();
	if (Math.abs(this.X - X) > 10 || Math.abs(this.Y - Y) > 10)
		return true;
	return false;
};
CMouseDownTrack.prototype.IsSamePos = function()
{
	this.Check();
	return this.Position === this.Page || this.Position === (this.Page + 1);
};

function COutlineMouseDownTrack(Thumbnails)
{
	this.Thumbnails = Thumbnails;
	this.SetDefault();
}
COutlineMouseDownTrack.prototype.SetDefault = function()
{
	this.Started = false;
	this.StartedSimple = true;
	this.X = -1;
	this.Y = -1;
	this.FocusPage = -1;
	this.IsHitInTextRect = false;
};
COutlineMouseDownTrack.prototype.Reset = function()
{
	this.SetDefault();
};
COutlineMouseDownTrack.prototype.Start = function(Page, X, Y)
{
	this.Started = true;
	this.StartedSimple = true;
	this.X = X;
	this.Y = Y;
};
COutlineMouseDownTrack.prototype.IsStarted = function()
{
	return this.Started;
};
COutlineMouseDownTrack.prototype.IsSimple = function()
{
	return this.StartedSimple;
};
COutlineMouseDownTrack.prototype.IsDragged = function()
{
	return this.IsStarted() && !this.IsSimple();
};
COutlineMouseDownTrack.prototype.ResetSimple = function()
{
	this.StartedSimple = false;
};
COutlineMouseDownTrack.prototype.GetX = function()
{
	return this.X;
};
COutlineMouseDownTrack.prototype.GetY = function()
{
	return this.Y;
};

COutlineMouseDownTrack.prototype.IsMoved = function(X, Y)
{
	if (Math.abs(this.X - X) > 10 || Math.abs(this.Y - Y) > 10)
		return true;
	return false;
};
COutlineMouseDownTrack.prototype.SetFocusPage = function (page) {
	this.FocusPage = page;
};
COutlineMouseDownTrack.prototype.SetIsHitInTextRect = function (pr) {
	this.IsHitInTextRect = pr;
};

window["AscCommon"] = window["AscCommon"] || {};
window["AscCommon"].CMouseDownTrack = CMouseDownTrack;
window["AscCommon"].COutlineMouseDownTrack = COutlineMouseDownTrack;
