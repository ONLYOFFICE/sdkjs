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
(function(){

	window.AscEmbed = window.AscEmbed || {};

	function GlobalScrollLocker()
	{
		this.frames = [];
		this.x = window.scrollX;
		this.y = window.scrollY;
		this.lockCounter = 0;

		window.addEventListener("scroll", this.onScroll.bind(this), true);
	}

	GlobalScrollLocker.prototype.onScroll = function()
	{
		let isPrevent = false;

		if (0 !== this.lockCounter)
			isPrevent = true;

		if (!isPrevent)
		{
			for (let i = 0, len = this.frames.length; i < len; i++)
			{
				if (document.activeElement === this.frames[i])
				{
					isPrevent = true;
					this.lockWithTimeout(500);
					break;
				}
			}
		}

		if (isPrevent)
		{
			window.scrollTo(this.x, this.y);
			return;
		}

		this.x = window.scrollX;
		this.y = window.scrollY;
	};

	GlobalScrollLocker.prototype.push = function(frame)
	{
		this.frames.push(frame);
	};

	GlobalScrollLocker.prototype.lockWithTimeout = function(interval)
	{
		this.lockCounter++;
		var _t = this;
		setTimeout(function(){
			_t.lockCounter--;
		}, interval);
	};

	function ScrollLocker(frame)
	{
		this.frame  = frame;

		if (!window.AscEmbed.AscEmbedGlobalScroller)
			window.AscEmbed.AscEmbedGlobalScroller = new GlobalScrollLocker();

		window.AscEmbed.AscEmbedGlobalScroller.push(this.frame);

		window.addEventListener("blur", this.onBlur.bind(this), false);

		window.addEventListener("pointermove", this.onMove.bind(this), false);
		window.addEventListener("wheel", this.onMove.bind(this), false);

		this.frame.addEventListener("pointerover", this.onOver.bind(this), false);
		this.frame.addEventListener("pointerleave", this.onLeave.bind(this), false);
	}

	ScrollLocker.prototype.onBlur = function()
	{
		if (document.activeElement === this.frame)
		{
			window.AscEmbed.AscEmbedGlobalScroller.lockWithTimeout(500);
		}
	};

	ScrollLocker.prototype.onOver = function()
	{
	};

	ScrollLocker.prototype.onLeave = function()
	{
		window.AscEmbed.AscEmbedGlobalScroller.lockWithTimeout(100);
		this.frame.blur();
	};

	ScrollLocker.prototype.onMove = function()
	{
		if (document.activeElement === this.frame)
		{
			window.AscEmbed.AscEmbedGlobalScroller.lockWithTimeout(100);
			this.frame.blur();
		}
	};

	window.AscEmbed.initWorker = function(frame)
	{
		window.AscEmbed.workers = window.AscEmbed.workers || [];
		let worker = new ScrollLocker(frame);
		window.AscEmbed.workers.push(worker);
		return worker;
	};
})();
