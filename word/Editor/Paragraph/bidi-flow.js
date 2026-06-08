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
	 * Class for handling bidirectional flow of text or other content
	 * @param handler - handler for elements in the flow
	 * @constructor
	 */
	function BidiFlow(handler)
	{
		this.handler   = handler;
		this.buffer    = [];
		this.direction = null;
		
		
		this.neutralBuffer = [];
		this.bufferLTR     = [];
	}
	/**
	 * @param isRtlDirection - main flow direction
	 */
	BidiFlow.prototype.begin = function(isRtlDirection)
	{
		this.direction            = isRtlDirection ? AscBidi.DIRECTION.R : AscBidi.DIRECTION.L;
		this.buffer.length        = 0;
		this.neutralBuffer.length = 0;
		this.bufferLTR.length     = 0;
	};
	BidiFlow.prototype.add = function(element, bidiType)
	{
		if (bidiType === AscBidi.TYPE.PM)
		{
			if (AscBidi.DIRECTION.L === this.direction && this.buffer.length)
				this.end();
			
			bidiType = AscBidi.DIRECTION.R === this.direction ? AscBidi.TYPE.R : AscBidi.TYPE.L;
		}
		
		if (bidiType === AscBidi.TYPE.B)
			this.end();
		
		if (AscBidi.DIRECTION.R === this.direction)
		{
			if (bidiType & AscBidi.FLAG.STRONG && bidiType & AscBidi.FLAG.RTL)
			{
				this.flushLTR();
				this.flushNeutralRTL()
				this.buffer.push([element, AscBidi.DIRECTION.R]);
			}
			else if (bidiType & AscBidi.FLAG.STRONG)
			{
				this.flushNeutralLTR();
				this.bufferLTR.push(element);
			}
			else
			{
				this.neutralBuffer.push([bidiType, element]);
			}
		}
		else
		{
			if (bidiType & AscBidi.FLAG.STRONG && bidiType & AscBidi.FLAG.RTL)
			{
				this.flushNeutralRTL();
				this.buffer.push([element, AscBidi.DIRECTION.R]);
			}
			else if (bidiType & AscBidi.FLAG.STRONG || 0 === this.buffer.length)
			{
				this.flush();
				this.handler.handleBidiFlow(element, AscBidi.DIRECTION.L);
			}
			else
			{
				this.neutralBuffer.push([bidiType, element]);
			}
		}
	};
	BidiFlow.prototype.end = function()
	{
		if (AscBidi.DIRECTION.R === this.direction || this.buffer.length)
		{
			this.flushLTR();
			this.flushNeutralRTL();
		}
		
		this.flush();
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	BidiFlow.prototype.flush = function()
	{
		for (let i = this.buffer.length - 1; i >= 0; --i)
		{
			this.handler.handleBidiFlow(this.buffer[i][0], this.buffer[i][1]);
		}
		this.buffer.length = 0;
		
		for (let i = 0; i < this.neutralBuffer.length; ++i)
		{
			this.handler.handleBidiFlow(this.neutralBuffer[i][1], AscBidi.DIRECTION.L);
		}
		this.neutralBuffer.length = 0;
	};
	BidiFlow.prototype.flushNeutralRTL = function()
	{
		let weakBuffer   = [];
		let numberBuffer = [];
		let buffer       = this.buffer;
		
		function flushNumber()
		{
			for (let i = numberBuffer.length - 1; i >= 0; --i)
			{
				buffer.push([numberBuffer[i], AscBidi.DIRECTION.L]);
			}
			numberBuffer.length = 0;
		}
		
		function flushWeak()
		{
			for (let i = 0; i < weakBuffer.length; ++i)
			{
				let type = weakBuffer[i][0];
				if (type & AscBidi.FLAG.NUMBER
					|| (type & AscBidi.FLAG.NUMBER_SEP_TER
						&& 0 !== i
						&& i !== weakBuffer.length - 1
						&& weakBuffer[i - 1][0] & AscBidi.FLAG.NUMBER
						&& weakBuffer[i + 1][0] & AscBidi.FLAG.NUMBER))
				{
					numberBuffer.push(weakBuffer[i][1]);
				}
				else
				{
					flushNumber();
					buffer.push([weakBuffer[i][1], AscBidi.DIRECTION.L]);
				}
			}
			flushNumber();
			weakBuffer.length = 0;
		}
		
		for (let i = 0; i < this.neutralBuffer.length; ++i)
		{
			let type = this.neutralBuffer[i][0];
			if (AscBidi.FLAG.NEUTRAL & type)
			{
				flushWeak(this.buffer);
				this.buffer.push([this.neutralBuffer[i][1], AscBidi.DIRECTION.R]);
			}
			else
			{
				weakBuffer.push([type, this.neutralBuffer[i][1]]);
			}
		}
		flushWeak(this.buffer);
		this.neutralBuffer.length = 0;
	};
	BidiFlow.prototype.flushNeutralLTR = function()
	{
		if (!this.bufferLTR.length)
			return this.flushNeutralRTL();
		
		for (let i = 0; i < this.neutralBuffer.length; ++i)
		{
			this.bufferLTR.push(this.neutralBuffer[i][1]);
		}
		this.neutralBuffer.length = 0;
	};
	BidiFlow.prototype.flushLTR = function()
	{
		for (let i = this.bufferLTR.length - 1; i >= 0; --i)
		{
			this.buffer.push([this.bufferLTR[i], AscBidi.DIRECTION.L]);
		}
		this.bufferLTR.length = 0;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.BidiFlow = BidiFlow;
	
})(window);


