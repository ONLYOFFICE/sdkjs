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
(function(){

	window.AscEmbed = window.AscEmbed || {};
	function ScrollLocker(frame)
	{
		this.frame  = frame;
		this.x = window.scrollX;
		this.y = window.scrollY;
		this.lockCounter = 0;

		document.addEventListener("scroll", this.onScroll.bind(this), false);
		window.addEventListener("blur", this.onBlur.bind(this), false);

		window.addEventListener("pointermove", this.onMove.bind(this), false);
		window.addEventListener("wheel", this.onMove.bind(this), false);

		this.frame.addEventListener("pointerover", this.onOver.bind(this), false);
		this.frame.addEventListener("pointerleave", this.onLeave.bind(this), false);
	}

	ScrollLocker.prototype.onScroll = function()
	{
		if (document.activeElement === this.frame || (0 !== this.lockCounter))
		{
			window.scrollTo(this.x, this.y);
			return;
		}
		this.x = window.scrollX;
		this.y = window.scrollY;
	};

	ScrollLocker.prototype.onBlur = function()
	{
		if (document.activeElement === this.frame)
		{
			this.lockWithTimeout(500);
		}
	};

	ScrollLocker.prototype.onOver = function()
	{
	};

	ScrollLocker.prototype.onLeave = function()
	{
		this.lockWithTimeout(100);
		this.frame.blur();
	};

	ScrollLocker.prototype.onMove = function()
	{
		if (document.activeElement === this.frame)
		{
			this.lockWithTimeout(100);
			this.frame.blur();
		}
	};

	ScrollLocker.prototype.lockWithTimeout = function(interval)
	{
		this.lockCounter++;
		var _t = this;
		setTimeout(function(){
			_t.lockCounter--;
		}, interval);
	};

	window.AscEmbed.initWorker = function(frame)
	{
		window.AscEmbed.workers = window.AscEmbed.workers || [];
		let worker = new ScrollLocker(frame);
		window.AscEmbed.workers.push(worker);
		return worker;
	};
})();
