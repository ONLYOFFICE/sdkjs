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


