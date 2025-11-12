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

(function()
{
	/**
	 * Class for processing the current stack of complex fields when we passing through a paragraph
	 * @constructor
	 */
	function ParagraphComplexFieldStack()
	{
		this.CF     = [];
		this.hidden = null;
	}
	ParagraphComplexFieldStack.prototype.resetPage = function(paragraph, page)
	{
		this.hidden = null;
		
		let endInfo = paragraph.GetEndInfoByPage(page - 1);
		if (endInfo)
			this.CF = endInfo.GetComplexFields();
		else
			this.CF = [];
	};
	ParagraphComplexFieldStack.prototype.resetRange = function(paragraph, line, range)
	{
		if (!paragraph.IsRecalculated() || !paragraph.Lines[line] || !paragraph.Lines[line].Ranges[range])
			return;
		
		this.setState(paragraph.Lines[line].CF);
		if (range > 0)
		{
			let startPos = paragraph.Get_StartRangePos2(line, 0);
			let endPos   = paragraph.Get_EndRangePos2(line, range - 1);
			
			let _t = this;
			paragraph.CheckRunContent(function(run, startPos, endPos)
			{
				let isRemovedInReview = (reviewtype_Remove === run.GetReviewType());
				for (let pos = startPos; pos < endPos; ++pos)
				{
					let item     = run.private_CheckInstrText(run.Content[pos]);
					var itemType = item.Type;
					
					if (_t.isHiddenFieldContent() && para_End !== itemType && para_FieldChar !== itemType)
						continue;
					
					if (para_FieldChar === itemType)
						_t.processFieldCharAndCollectComplexField(item);
					else if (para_InstrText === itemType && !isRemovedInReview)
						_t.processInstruction(item);
				}
			}, startPos, endPos);
		}
	};
	/**
	 * @param element {AscWord.CRunElementBase}
	 * @returns {boolean}
	 */
	ParagraphComplexFieldStack.prototype.checkRunElement = function(element)
	{
		if ((this.isHiddenFieldContent() || this.isHiddenComplexFieldPart())
			&& para_End !== element.Type
			&& para_FieldChar !== element.Type)
			return false;
		
		if (para_FieldChar === element.Type)
			this.processFieldChar(element);
		
		return true;
	};
	/**
	 * Находимся ли мы внутри содержимого скрытой части поля или внутри скрытого поля
	 * @returns {boolean}
	 */
	ParagraphComplexFieldStack.prototype.isHiddenFieldContent = function()
	{
		this.updateHidden();
		return this.hidden;
	};
	/**
	 * Данная функция используется при пересчете, когда мы собираем сложное поле.
	 * @param oChar
	 */
	ParagraphComplexFieldStack.prototype.processFieldCharAndCollectComplexField = function(oChar)
	{
		this.hidden = null;
		
		if (oChar.IsBegin())
		{
			let oComplexField = oChar.GetComplexField();
			if (!oComplexField)
			{
				oChar.SetUse(false);
			}
			else
			{
				oChar.SetUse(true);
				oComplexField.SetBeginChar(oChar);
				this.CF.push(new CComplexFieldStatePos(oComplexField, true));
			}
		}
		else if (oChar.IsEnd())
		{
			if (this.CF.length > 0)
			{
				oChar.SetUse(true);
				let oComplexField = this.CF[this.CF.length - 1].ComplexField;
				oComplexField.SetEndChar(oChar);
				this.CF.splice(this.CF.length - 1, 1);
				
				if (this.CF.length > 0 && this.CF[this.CF.length - 1].IsFieldCode())
					this.CF[this.CF.length - 1].ComplexField.SetInstructionCF(oComplexField);
			}
			else
			{
				oChar.SetUse(false);
			}
		}
		else if (oChar.IsSeparate())
		{
			if (this.CF.length > 0)
			{
				oChar.SetUse(true);
				let oComplexField = this.CF[this.CF.length - 1].ComplexField;
				oComplexField.SetSeparateChar(oChar);
				this.CF[this.CF.length - 1].SetFieldCode(false);
			}
			else
			{
				oChar.SetUse(false);
			}
		}
	};
	/**
	 * Данная функция используется, когда мы просто хотим отследить где мы находимся, относительно сложных полей
	 * @param oChar
	 */
	ParagraphComplexFieldStack.prototype.processFieldChar = function(oChar)
	{
		this.hidden = null;
		
		if (!oChar || !oChar.IsUse())
			return;
		
		let oComplexField = oChar.GetComplexField();
		
		if (oChar.IsBegin())
		{
			this.CF.push(new CComplexFieldStatePos(oComplexField, true));
		}
		else if (oChar.IsSeparate())
		{
			if (this.CF.length > 0)
			{
				this.CF[this.CF.length - 1].SetFieldCode(false);
			}
		}
		else if (oChar.IsEnd())
		{
			if (this.CF.length > 0)
			{
				this.CF.splice(this.CF.length - 1, 1);
			}
		}
	};
	ParagraphComplexFieldStack.prototype.processInstruction = function(oInstruction)
	{
		if (this.CF.length <= 0)
			return;
		
		let oComplexField = this.CF[this.CF.length - 1].ComplexField;
		if (oComplexField && null === oComplexField.GetSeparateChar())
			oComplexField.SetInstruction(oInstruction);
	};
	ParagraphComplexFieldStack.prototype.isComplexField = function()
	{
		return (this.CF.length > 0);
	};
	ParagraphComplexFieldStack.prototype.isComplexFieldCode = function()
	{
		for (let index = 0, count = this.CF.length; index < count; ++index)
		{
			if (this.CF[index].IsFieldCode())
				return true;
		}
		
		return false;
	};
	ParagraphComplexFieldStack.prototype.isHiddenComplexFieldPart = function()
	{
		for (let fieldIndex = 0, fieldCount = this.CF.length; fieldIndex < fieldCount; ++ fieldIndex)
		{
			let isFieldCode = this.CF[fieldIndex].IsFieldCode();
			let isShowCode  = this.CF[fieldIndex].IsShowFieldCode();
			if (isFieldCode !== isShowCode)
				return true;
		}
		
		return false;
	};
	ParagraphComplexFieldStack.prototype.isCurrentComplexField = function()
	{
		for (let index = 0, count = this.CF.length; index < count; ++index)
		{
			if (this.CF[index].ComplexField.IsCurrent())
				return true;
		}
		
		return false;
	};
	ParagraphComplexFieldStack.prototype.isHyperlinkField = function()
	{
		let haveHyperlink  = false;
		let haveOtherField = false;
		
		for (let index = 0, count = this.CF.length; index < count; ++index)
		{
			let instruction = this.CF[index].ComplexField.GetInstruction();
			if (instruction && AscWord.fieldtype_HYPERLINK === instruction.GetType())
				haveHyperlink = true;
			else
				haveOtherField = true;
			
		}
		
		return (haveHyperlink && !haveOtherField);
	};
	ParagraphComplexFieldStack.prototype.getState = function()
	{
		let state = {
			hidden : this.hidden,
			CF     : []
		};
		
		for (let index = 0, count = this.CF.length; index < count; ++index)
		{
			state.CF[index] = this.CF[index].Copy();
		}
		
		return state;
	};
	ParagraphComplexFieldStack.prototype.setState = function(state)
	{
		if (!state)
			return;
		
		this.hidden = state.hidden;
		this.CF     = state.CF;
	};
	ParagraphComplexFieldStack.prototype.getReferenceToHyperlink = function()
	{
		for (let index = this.CF.length - 1; index >= 0; --index)
		{
			let instruction = this.CF[index].ComplexField.GetInstruction();
			if (instruction &&
				(AscWord.fieldtype_HYPERLINK === instruction.GetType()
					|| AscWord.fieldtype_REF === instruction.GetType() && instruction.GetHyperlink()
					|| AscWord.fieldtype_NOTEREF === instruction.GetType() && instruction.GetHyperlink()))
				return this.CF[index].ComplexField;
		}
		
		return null;
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	ParagraphComplexFieldStack.prototype.updateHidden = function()
	{
		if (null !== this.hidden)
			return;
		
		this.hidden = false;
		for (let index = 0, count = this.CF.length; index < count; ++index)
		{
			if (this.CF[index].ComplexField.IsHidden())
			{
				this.hidden = true;
				break;
			}
		}
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.ParagraphComplexFieldStack = ParagraphComplexFieldStack;
	
})(window);
