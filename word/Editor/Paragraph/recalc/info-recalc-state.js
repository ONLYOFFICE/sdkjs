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
	function InfoRecalcState()
	{
		AscWord.ParagraphRecalculateStateBase.call(this);
		this.fast          = false;
		this.Comments      = [];
		this.ComplexFields = [];
		this.PermRanges    = [];
	}
	InfoRecalcState.prototype = Object.create(AscWord.ParagraphRecalculateStateBase.prototype);
	InfoRecalcState.prototype.constructor = InfoRecalcState;
	InfoRecalcState.prototype.setFast = function(isFast)
	{
		this.fast = isFast;
	};
	InfoRecalcState.prototype.isFastRecalculation = function()
	{
		return this.fast;
	};
	InfoRecalcState.prototype.Reset = function(prevInfo)
	{
		this.Comments      = [];
		this.ComplexFields = [];
		this.PermRanges    = [];

		if (!prevInfo)
			return;

		if (prevInfo.Comments)
			this.Comments = prevInfo.Comments.slice();

		if (prevInfo.ComplexFields)
		{
			for (let index = 0, count = prevInfo.ComplexFields.length; index < count; ++index)
			{
				this.ComplexFields[index] = prevInfo.ComplexFields[index].Copy();
			}
		}

		if (prevInfo.PermRanges)
			this.PermRanges = prevInfo.PermRanges.slice();
	};
	InfoRecalcState.prototype.AddComment = function(Id)
	{
		this.Comments.push(Id);
	};
	InfoRecalcState.prototype.RemoveComment = function(Id)
	{
		var CommentsLen = this.Comments.length;
		for (var CurPos = 0; CurPos < CommentsLen; CurPos++)
		{
			if (this.Comments[CurPos] === Id)
			{
				this.Comments.splice(CurPos, 1);
				break;
			}
		}
	};
	InfoRecalcState.prototype.addPermRange = function(rangeId)
	{
		this.PermRanges.push(rangeId);
	};
	InfoRecalcState.prototype.removePermRange = function(rangeId)
	{
		let pos = this.PermRanges.indexOf(rangeId);
		if (-1 === pos)
			return;

		if (this.PermRanges.length - 1 === pos)
			--this.PermRanges.length;
		else
			this.PermRanges.splice(pos, 1);
	};
	InfoRecalcState.prototype.processFieldChar = function(oFieldChar)
	{
		if (!oFieldChar || !oFieldChar.IsUse())
			return;

		var oComplexField = oFieldChar.GetComplexField();

		if (oFieldChar.IsBegin())
		{
			this.ComplexFields.push(new CComplexFieldStatePos(oComplexField, true));
		}
		else if (oFieldChar.IsSeparate())
		{
			for (var nIndex = 0, nCount = this.ComplexFields.length; nIndex < nCount; ++nIndex)
			{
				if (oComplexField === this.ComplexFields[nIndex].ComplexField)
				{
					this.ComplexFields[nIndex].SetFieldCode(false);
					break;
				}
			}
		}
		else if (oFieldChar.IsEnd())
		{
			for (var nIndex = 0, nCount = this.ComplexFields.length; nIndex < nCount; ++nIndex)
			{
				if (oComplexField === this.ComplexFields[nIndex].ComplexField)
				{
					this.ComplexFields.splice(nIndex, 1);
					break;
				}
			}
		}
	};
	InfoRecalcState.prototype.isComplexField = function()
	{
		return (this.ComplexFields.length > 0 ? true : false);
	};
	InfoRecalcState.prototype.isComplexFieldCode = function()
	{
		if (!this.isComplexField())
			return false;

		for (var nIndex = 0, nCount = this.ComplexFields.length; nIndex < nCount; ++nIndex)
		{
			if (this.ComplexFields[nIndex].IsFieldCode())
				return true;
		}

		return false;
	};
	InfoRecalcState.prototype.isHiddenComplexFieldPart = function()
	{
		for (let fieldIndex = 0, fieldCount = this.ComplexFields.length; fieldIndex < fieldCount; ++fieldIndex)
		{
			let isFieldCode = this.ComplexFields[fieldIndex].IsFieldCode();
			let isShowCode  = this.ComplexFields[fieldIndex].IsShowFieldCode();
			if (isFieldCode !== isShowCode)
				return true;
		}

		return false;
	};
	InfoRecalcState.prototype.processFieldCharAndCollectComplexField = function(oChar)
	{
		if (oChar.IsBegin())
		{
			var oComplexField = oChar.GetComplexField();
			if (!oComplexField)
			{
				oChar.SetUse(false);
			}
			else
			{
				oChar.SetUse(true);
				oComplexField.SetBeginChar(oChar);
				this.ComplexFields.push(new CComplexFieldStatePos(oComplexField, true));
			}
		}
		else if (oChar.IsEnd())
		{
			if (this.ComplexFields.length > 0)
			{
				oChar.SetUse(true);
				var oComplexField = this.ComplexFields[this.ComplexFields.length - 1].ComplexField;
				oComplexField.SetEndChar(oChar);
				this.ComplexFields.splice(this.ComplexFields.length - 1, 1);

				if (this.ComplexFields.length > 0 && this.ComplexFields[this.ComplexFields.length - 1].IsFieldCode())
					this.ComplexFields[this.ComplexFields.length - 1].ComplexField.SetInstructionCF(oComplexField);
			}
			else
			{
				oChar.SetUse(false);
			}
		}
		else if (oChar.IsSeparate())
		{
			if (this.ComplexFields.length > 0)
			{
				oChar.SetUse(true);
				var oComplexField = this.ComplexFields[this.ComplexFields.length - 1].ComplexField;
				oComplexField.SetSeparateChar(oChar);
				this.ComplexFields[this.ComplexFields.length - 1].SetFieldCode(false);
			}
			else
			{
				oChar.SetUse(false);
			}
		}
	};
	InfoRecalcState.prototype.processInstruction = function(oInstruction)
	{
		if (this.ComplexFields.length <= 0)
			return;

		var oComplexField = this.ComplexFields[this.ComplexFields.length - 1].ComplexField;
		if (oComplexField && null === oComplexField.GetSeparateChar())
			oComplexField.SetInstruction(oInstruction);
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.Paragraph.InfoRecalcState = InfoRecalcState;

})(window);
