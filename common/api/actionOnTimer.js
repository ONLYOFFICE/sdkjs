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

(function (window)
{
	const MAX_ACTION_TIME = 20;
	const TIMEOUT_TIME    = 20;

	/**
	 * @constructor
	 */
	function CActionOnTimerBase()
	{
		this.TimerId            = null;
		this.Start              = false;
		this.FirstActionOnTimer = false;
	}
	CActionOnTimerBase.prototype.Begin = function()
	{
		if (this.Start)
			this.End();

		this.Start = true;

		this.OnBegin.apply(this, arguments);

		let oThis = this;
		if (this.FirstActionOnTimer)
			this.TimerId = setTimeout(function(){oThis.Continue();}, TIMEOUT_TIME);
		else
			this.Continue();
	};
	CActionOnTimerBase.prototype.Continue = function()
	{
		if (this.IsContinue())
		{
			let nTime = performance.now();
			this.OnStartTimer();

			while (this.IsContinue())
			{
				if (performance.now() - nTime > MAX_ACTION_TIME)
					break;

				this.DoAction();
			}

			this.OnEndTimer();
		}

		let oThis = this;
		if (this.IsContinue())
			this.TimerId = setTimeout(function(){oThis.Continue();}, TIMEOUT_TIME);
		else
			this.End();
	};
	CActionOnTimerBase.prototype.End = function()
	{
		this.Reset();

		if (this.Start)
		{
			this.Start = false;
			this.OnEnd();
		}
	};
	CActionOnTimerBase.prototype.Stop = function()
	{
		this.Reset();
		this.Start = false;
	};
	CActionOnTimerBase.prototype.IsInProgress = function()
	{
		return this.Start;
	};
	CActionOnTimerBase.prototype.Reset = function()
	{
		if (this.TimerId)
			clearTimeout(this.TimerId);

		this.TimerId = null;
		this.Index   = -1;
	};
	CActionOnTimerBase.prototype.SetDoFirstActionOnTimer = function(isOnTimer)
	{
		this.FirstActionOnTimer = isOnTimer;
	};
	//------------------------------------------------------------------------------------------------------------------
	// The following functions need to be overridden in the derived class
	//------------------------------------------------------------------------------------------------------------------
	CActionOnTimerBase.prototype.OnBegin = function()
	{
	};
	CActionOnTimerBase.prototype.OnEnd = function()
	{
	};
	CActionOnTimerBase.prototype.IsContinue = function()
	{
		return false;
	};
	CActionOnTimerBase.prototype.OnStartTimer = function()
	{
	};
	CActionOnTimerBase.prototype.DoAction = function()
	{
	};
	CActionOnTimerBase.prototype.OnEndTimer = function()
	{
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].CActionOnTimerBase = CActionOnTimerBase;

})(window);
