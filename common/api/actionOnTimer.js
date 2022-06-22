/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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
	// Следующие функции нужно переопределять в классе-наследнике
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
