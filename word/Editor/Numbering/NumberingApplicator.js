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
	const NumberingType = Asc.c_oAscJSONNumberingType;

	/**
	 * Класс для применения нумерации к документу
	 * @param {AscWord.CDocument} document
	 * @constructor
	 */
	function CNumberingApplicator(document)
	{
		this.Document  = document;
		this.Numbering = document.GetNumbering();

		this.NumPr      = null;
		this.Paragraphs = [];
		this.NumInfo    = new AscWord.CNumInfo();

		this.LastBulleted = null;
		this.LastNumbered = null;
	}

	/**
	 * Применяем нумерацию по заданному объекту
	 * @param numInfo {object}
	 */
	CNumberingApplicator.prototype.Apply = function(numInfo)
	{
		if (!this.Document || !numInfo)
			return false;

		this.NumInfo    = numInfo;
		this.NumPr      = this.GetCurrentNumPr();
		this.Paragraphs = this.GetParagraphs();

		if (!this.Paragraphs.length)
			return false;

		let result = false;
		if (this.IsRemoveNumbering())
			result = this.RemoveNumbering();
		else if (this.IsBulleted())
			result = this.ApplyBulleted();
		else if (this.IsNumbered())
			result = this.ApplyNumbered();
		else if (this.IsSingleLevel())
			result = this.ApplySingleLevel();
		else if (this.IsMultilevel())
			result = this.ApplyMultilevel();

		if (result)
			this.UpdateDocumentOutline();

		return result;
	};
	CNumberingApplicator.prototype.ResetLast = function()
	{
		this.SetLastNumbered(null);
		this.SetLastBulleted(null);
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CNumberingApplicator.prototype.GetLastBulleted = function()
	{
		return this.LastBulleted;
	};
	CNumberingApplicator.prototype.SetLastBulleted = function(numId, ilvl)
	{
		if (!numId)
			this.LastBulleted = null;
		else
			this.LastBulleted = new AscWord.NumPr(numId, ilvl);
	};
	CNumberingApplicator.prototype.GetLastNumbered = function()
	{
		return this.LastNumbered;
	};
	CNumberingApplicator.prototype.SetLastNumbered = function(numId, ilvl)
	{
		if (!numId)
			this.LastNumbered = null;
		else
			this.LastNumbered = new AscWord.NumPr(numId, ilvl);
	};
	CNumberingApplicator.prototype.GetCurrentNumPr = function()
	{
		let document = this.Document;

		let numPr = document.GetSelectedNum();
		if (!numPr && !document.IsTextSelectionUse())
			numPr = document.GetSelectedNum(true);

		return numPr;
	};
	CNumberingApplicator.prototype.GetParagraphs = function()
	{
		let paragraphs = [];
		if (this.NumPr)
			paragraphs = this.Document.GetAllParagraphsByNumbering(this.NumPr);
		else
			paragraphs = this.Document.GetSelectedParagraphs();

		return paragraphs;
	};
	CNumberingApplicator.prototype.IsRemoveNumbering = function()
	{
		return (NumberingType.Remove === this.NumInfo.Type);
	};
	CNumberingApplicator.prototype.IsBulleted = function()
	{
		return (NumberingType.Bullet === this.NumInfo.Type && 0 === this.NumInfo.Lvl.length);
	};
	CNumberingApplicator.prototype.IsNumbered = function()
	{
		return (NumberingType.Number === this.NumInfo.Type && 0 === this.NumInfo.Lvl.length);
	};
	CNumberingApplicator.prototype.IsSingleLevel = function()
	{
		return (NumberingType.Remove !== this.NumInfo.Type && 1 === this.NumInfo.Lvl.length);
	};
	CNumberingApplicator.prototype.IsMultilevel = function()
	{
		return (this.NumInfo.Lvl.length >= 8);
	};
	CNumberingApplicator.prototype.RemoveNumbering = function()
	{
		let document = this.Document;
		if (document.IsNumberingSelection())
			document.RemoveSelection();

		for (let index = 0, count = this.Paragraphs.length; index < count; ++index)
		{
			this.Paragraphs[index].RemoveNumPr();
		}
		return true;
	};
	CNumberingApplicator.prototype.ApplyBulleted = function()
	{
		if (this.ApplyBulletedToCurrent())
			return true;

		// 1. Пытаемся присоединить список к списку предыдущего параграфа (если только он маркированный)
		// 2. Пытаемся присоединить список к списку следующего параграфа (если он маркированный)
		// 3. Пытаемся добавить список, который добавлялся предыдущий раз
		// 4. Создаем новый маркированный список

		let numberingManager = this.Document.GetNumbering();

		let numId = null;
		let ilvl  = 0;

		let prevNumPr = this.GetPrevNumPr();
		if (prevNumPr && numberingManager.CheckFormat(prevNumPr.NumId, prevNumPr.Lvl, Asc.c_oAscNumberingFormat.Bullet))
		{
			numId = prevNumPr.NumId;
			ilvl  = prevNumPr.Lvl;
		}

		if (!numId)
		{
			let nextNumPr = this.GetNextNumPr();
			if (nextNumPr && numberingManager.CheckFormat(nextNumPr.NumId, nextNumPr.Lvl, Asc.c_oAscNumberingFormat.Bullet))
			{
				numId = nextNumPr.NumId;
				ilvl  = nextNumPr.Lvl;
			}
		}

		let isCheckPrev = false;
		if (!numId)
		{
			let lastNumPr = this.GetLastBulleted();
			let lastNum   = lastNumPr ? this.Numbering.GetNum(lastNumPr.NumId) : null;
			if (lastNum && lastNum.GetLvl(0).IsBulleted())
			{
				let newNum = this.Numbering.CreateNum();
				newNum.CreateDefault(c_oAscMultiLevelNumbering.Bullet);
				newNum.SetLvl(lastNum.GetLvl(lastNumPr.Lvl).Copy(), 0);

				numId = newNum.GetId();
				ilvl  = 0;

				isCheckPrev = true;
			}
		}

		if (!numId)
		{
			let newNum = this.Numbering.CreateNum();
			newNum.CreateDefault(c_oAscMultiLevelNumbering.Bullet);

			numId = newNum.GetId();
			ilvl  = 0;

			isCheckPrev = true;
		}

		if (isCheckPrev)
		{
			let result = this.CheckPrevNumPr(numId, ilvl);
			if (result)
			{
				numId = result.NumId;
				ilvl  = result.Lvl;
			}
		}
		
		this.MergeTextPrFromCommonNum(numId, ilvl);
		this.SetLastBulleted(numId, ilvl);
		this.ApplyNumPr(numId, ilvl);
		return true;
	};
	CNumberingApplicator.prototype.ApplyBulletedToCurrent = function()
	{
		let numPr = this.NumPr;
		if (!numPr)
			return false;

		let num = this.Numbering.GetNum(numPr.NumId);
		if (!num)
			return false;

		let lvl;

		let lastNumPr = this.GetLastBulleted();
		let lastNum   = lastNumPr ? this.Numbering.GetNum(lastNumPr.NumId) : null;
		if (lastNum && lastNum.GetLvl(numPr.Lvl).IsBulleted())
		{
			lvl = lastNum.GetLvl(lastNumPr.Lvl).Copy();
		}
		else
		{
			lvl = num.GetLvl(numPr.Lvl).Copy();

			let textPr = lvl.GetTextPr().Copy();
			textPr.RFonts.SetAll("Symbol");
			lvl.SetByType(c_oAscNumberingLevel.Bullet, numPr.Lvl, String.fromCharCode(0x00B7), textPr);
		}

		lvl.ParaPr = num.GetLvl(numPr.Lvl).ParaPr.Copy();

		num.SetLvl(lvl, numPr.Lvl);
		this.SetLastBulleted(numPr.NumId, numPr.Lvl);
		return true;
	};
	CNumberingApplicator.prototype.ApplyNumbered = function()
	{
		if (this.ApplyNumberedToCurrent())
			return true;

		// 1. Пытаемся присоединить список к списку предыдущего параграфа (если он нумерованный)
		// 2. Пытаемся присоединить список к списку следующего параграфа (если он нумерованный)
		// 3. Пытаемся добавить список, который добавлялся предыдущий раз (добавляем его копию, и опционально продолжаем)
		// 4. Создаем новый нумерованный список

		let numberingManager = this.Document.GetNumbering();

		let numId = null;
		let ilvl  = 0;

		let prevNumPr = this.GetPrevNumPr();
		if (prevNumPr && numberingManager.CheckFormat(prevNumPr.NumId, prevNumPr.Lvl, Asc.c_oAscNumberingFormat.Decimal))
		{
			numId = prevNumPr.NumId;
			ilvl  = undefined;
		}

		if (!numId)
		{
			let nextNumPr = this.GetNextNumPr();
			if (nextNumPr && numberingManager.CheckFormat(nextNumPr.NumId, nextNumPr.Lvl, Asc.c_oAscNumberingFormat.Decimal))
			{
				numId = nextNumPr.NumId;
				ilvl  = undefined;
			}
		}

		let isCheckPrev = false;
		if (!numId)
		{
			let lastNumPr = this.GetLastNumbered();
			let lastNum   = lastNumPr ? this.Numbering.GetNum(lastNumPr.NumId) : null;
			if (lastNum && lastNum.GetLvl(0).IsNumbered())
			{
				let newNum = this.Numbering.CreateNum();

				if (lastNum.IsHaveRelatedLvlText())
				{
					for (let ilvlTemp = 0; ilvlTemp < 9; ++ilvlTemp)
					{
						newNum.SetLvl(lastNum.GetLvl(ilvlTemp).Copy(), ilvlTemp);
					}
				}
				else
				{
					newNum.CreateDefault(c_oAscMultiLevelNumbering.Numbered);
					newNum.SetLvl(lastNum.GetLvl(lastNumPr.Lvl).Copy(), 0);
				}

				numId = newNum.GetId();
				ilvl  = 0;

				isCheckPrev = true;
			}
		}

		if (!numId)
		{
			let newNum = this.Numbering.CreateNum();
			newNum.CreateDefault(c_oAscMultiLevelNumbering.Numbered);

			numId = newNum.GetId();
			ilvl  = 0;

			isCheckPrev = true;
		}

		if (isCheckPrev)
		{
			let result = this.CheckPrevNumPr(numId, ilvl);
			if (result)
			{
				numId = result.NumId;
				ilvl  = result.Lvl;
			}
		}
		
		let _ilvl = undefined !== ilvl ? ilvl : 0;
		
		this.MergeTextPrFromCommonNum(numId, _ilvl);
		this.SetLastNumbered(numId, _ilvl);
		this.ApplyNumPr(numId, ilvl);
		return true;
	};
	CNumberingApplicator.prototype.ApplyNumberedToCurrent = function()
	{
		let numPr = this.NumPr;
		if (!numPr)
			return false;

		let num = this.Numbering.GetNum(numPr.NumId);
		if (!num)
			return false;

		let lvl;

		let lastNumPr = this.GetLastNumbered();
		let lastNum   = lastNumPr ? this.Numbering.GetNum(lastNumPr.NumId) : null;
		if (lastNum && lastNum.GetLvl(numPr.Lvl).IsNumbered())
		{
			if (lastNum.IsHaveRelatedLvlText())
			{
				// В этом случае мы не можем подменить просто текущий уровень, меняем целиком весь список
				for (let iLvl = 0; iLvl < 9; ++iLvl)
				{
					num.SetLvl(lastNum.GetLvl(iLvl).Copy(), iLvl);
				}
			}
			else
			{
				lvl        = lastNum.GetLvl(lastNum.Lvl).Copy();
				lvl.ParaPr = num.GetLvl(numPr.Lvl).ParaPr.Copy();
				lvl.ResetNumberedText(numPr.Lvl);
			}
		}
		else
		{
			lvl = num.GetLvl(numPr.Lvl).Copy();
			lvl.SetByType(c_oAscNumberingLevel.DecimalDot_Right, numPr.Lvl);
			lvl.ParaPr = num.GetLvl(numPr.Lvl).ParaPr.Copy();
		}

		if (lvl)
			num.SetLvl(lvl, numPr.Lvl);

		this.SetLastNumbered(numPr.NumId, numPr.Lvl);
		return true;
	};
	CNumberingApplicator.prototype.ApplySingleLevel = function()
	{
		let numLvl = this.CreateSingleNumberingLvl();
		if (!numLvl)
			return false;

		let commonNumPr = this.NumPr ? this.NumPr : null;
		if (commonNumPr && commonNumPr.NumId)
		{
			let num = this.Numbering.GetNum(commonNumPr.NumId);
			if (num)
			{
				this.MergePrToLvl(num.GetLvl(commonNumPr.Lvl), numLvl);
				numLvl.ResetNumberedText(commonNumPr.Lvl);
				num.SetLvl(numLvl, commonNumPr.Lvl);
				this.SetLastSingleLevel(commonNumPr.NumId, commonNumPr.Lvl);
			}
		}
		else
		{
			let num   = this.CreateBaseNum();
			let numId = num.GetId();
			let ilvl  = !commonNumPr || !commonNumPr.Lvl ? 0 : commonNumPr.Lvl;
			
			this.MergePrToLvl(num.GetLvl(ilvl), numLvl);

			numLvl.ResetNumberedText(ilvl);
			num.SetLvl(numLvl, ilvl);

			let prevNumPr = this.CheckPrevNumPr(numId, ilvl);
			if (prevNumPr)
			{
				numId = prevNumPr.NumId;
				ilvl  = prevNumPr.Lvl;
			}

			this.ApplyNumPr(numId, ilvl);
			this.SetLastSingleLevel(numId, ilvl);
		}

		return true;
	};
	CNumberingApplicator.prototype.SetLastSingleLevel = function(numId, ilvl)
	{
		if (NumberingType.Bullet === this.NumInfo.Type)
			this.SetLastBulleted(numId, ilvl);
		else if (NumberingType.Number === this.NumInfo.Type)
			this.SetLastNumbered(numId, ilvl);
	};
	CNumberingApplicator.prototype.ApplyMultilevel = function()
	{
		let commonNumId = this.NumPr ? this.NumPr.NumId : null;

		let num, numId;
		if (commonNumId)
		{
			num = this.Numbering.GetNum(commonNumId);
		}
		else
		{
			num   = this.CreateBaseNum();
			numId = num.GetId();
		}

		if (!num)
			return false;

		for (let iLvl = 0; iLvl < 9; ++iLvl)
		{
			let numLvl = this.CreateSingleNumberingLvl(iLvl);
			if (numLvl)
			{
				let pStyle = num.GetLvl(iLvl).GetPStyle();
				if (pStyle)
					numLvl.SetPStyle(pStyle);
				
				num.SetLvl(numLvl, iLvl);
			}
		}

		if (numId)
		{
			this.ApplyNumPr(numId, 0);
			this.ApplyToHeadings(numId);
		}
		else
		{
			this.LinkNumWithHeadings(num);
		}

		return true;
	};
	CNumberingApplicator.prototype.GetPrevNumPr = function()
	{
		if (!this.Paragraphs || !this.Paragraphs.length)
			return null;

		let prevParagraph = this.Paragraphs[0].GetPrevParagraph();
		while (prevParagraph)
		{
			if (prevParagraph.GetNumPr() || !prevParagraph.IsEmpty())
				break;
			
			prevParagraph = prevParagraph.GetPrevParagraph();
		}
		
		return prevParagraph ? prevParagraph.GetNumPr() : null;
	};
	CNumberingApplicator.prototype.GetNextNumPr = function()
	{
		if (!this.Paragraphs || !this.Paragraphs.length)
			return null;
		
		let nextParagraph = this.Paragraphs[this.Paragraphs.length - 1].GetNextParagraph();
		while (nextParagraph)
		{
			if (nextParagraph.GetNumPr() || !nextParagraph.IsEmpty())
				break;
			
			nextParagraph = nextParagraph.GetNextParagraph();
		}
		
		return nextParagraph ? nextParagraph.GetNumPr() : null;
	};
	CNumberingApplicator.prototype.CheckPrevNumPr = function(numId, ilvl)
	{
		if (this.Paragraphs.length !== 1 || this.Document.IsSelectionUse())
			return new AscWord.NumPr(numId, ilvl);

		var prevParagraph = this.Paragraphs[0].GetPrevParagraph();
		while (prevParagraph)
		{
			if (prevParagraph.GetNumPr() || !prevParagraph.IsEmpty())
				break;

			prevParagraph = prevParagraph.GetPrevParagraph();
		}

		let prevNumPr = prevParagraph ? prevParagraph.GetNumPr() : null;
		if (prevNumPr)
		{
			let prevLvl = this.Numbering.GetNum(prevNumPr.NumId).GetLvl(prevNumPr.Lvl);
			let currLvl = this.Numbering.GetNum(numId).GetLvl(ilvl);

			if (prevLvl.IsSimilar(currLvl))
				return new AscWord.NumPr(prevNumPr.NumId, prevNumPr.Lvl);
		}

		return new AscWord.NumPr(numId, ilvl);
	};
	CNumberingApplicator.prototype.MergeTextPrFromCommonNum = function(numId, iLvl)
	{
		let commonNumPr = this.NumPr ? this.NumPr : this.GetCommonNumPr(true);
		if (!commonNumPr || !commonNumPr.NumId)
			return;
		
		let num       = this.Numbering.GetNum(numId);
		let commonNum = this.Numbering.GetNum(commonNumPr.NumId);
		if (!num || !commonNum)
			return;
		
		let numLvl = num.GetLvl(iLvl);
		
		numLvl = numLvl.Copy();
		this.MergePrToLvl(commonNum.GetLvl(commonNumPr.Lvl), numLvl);
		num.SetLvl(numLvl, iLvl);
	};
	CNumberingApplicator.prototype.GetCommonNumPr = function(skipBulletedCheck)
	{
		let isDiffLvl = false;
		let isDiffId  = false;
		let numId     = null;
		let ilvl      = null;

		for (let index = 0, count = this.Paragraphs.length; index < count; ++index)
		{
			if (isDiffLvl)
				break;

			let numPr = this.Paragraphs[index].GetNumPr();
			if (numPr)
			{
				if (null === ilvl)
					ilvl = numPr.Lvl;
				else if (ilvl !== numPr.Lvl)
					isDiffLvl = true;

				if (null === numId)
					numId = numPr.NumId;
				else if (numId !== numPr.NumId)
					isDiffId = true;
			}
			else
			{
				isDiffLvl = true;
			}
		}

		if (isDiffLvl)
		{
			numId = null;
			ilvl  = 0;
		}
		else if (isDiffId)
		{
			numId = null;
		}
		else if (numId)
		{
			// TODO: Проверить нужна ли эта проверка, если да, то написать тесты (если нет - тоже)
			if (!skipBulletedCheck &&
				this.NumInfo.IsBulleted()
				&& !this.Numbering.CheckFormat(numId, ilvl, Asc.c_oAscNumberingFormat.Bullet))
			{
				numId = null;
			}
		}

		return new AscWord.NumPr(numId, ilvl);
	};
	CNumberingApplicator.prototype.GetCommonNumId = function()
	{
		let numId = null;
		for (let index = 0, count = this.Paragraphs.length; index < count; ++index)
		{
			let numPr = this.Paragraphs[index].GetNumPr();
			if (numPr)
			{
				if (null === numId)
					numId = numPr.NumId;
				else if (numId !== numPr.NumId)
					return null;
			}
			else
			{
				return null;
			}
		}
		return numId;
	};
	CNumberingApplicator.prototype.CreateSingleNumberingLvl = function(ilvl)
	{
		let _ilvl = ilvl ? ilvl : 0;

		if (!this.NumInfo
			|| !this.NumInfo.Lvl
			|| this.NumInfo.Lvl.length <= _ilvl)
			return null;

		return AscWord.CNumberingLvl.FromJson(this.NumInfo.Lvl[_ilvl]);
	};
	CNumberingApplicator.prototype.CreateBaseNum = function()
	{
		let num = this.Numbering.CreateNum();

		if (NumberingType.Bullet === this.NumInfo.Type)
			num.CreateDefault(c_oAscMultiLevelNumbering.Bullet);
		else if (NumberingType.Number === this.NumInfo.Type)
			num.CreateDefault(c_oAscMultiLevelNumbering.Numbered);

		return num;
	};
	CNumberingApplicator.prototype.ApplyNumPr = function(numId, ilvl)
	{
		// TODO: Надо перенести сюда всю логику подбора уровня по отступам, а в классе параграфа делать простое выставление
		let checkIndents = this.IsCheckIndents();
		for (let index = 0, count = this.Paragraphs.length; index < count; ++index)
		{
			let paragraph = this.Paragraphs[index];
			let oldNumPr  = paragraph.GetNumPr();

			if (oldNumPr)
				paragraph.ApplyNumPr(numId, oldNumPr.Lvl, checkIndents);
			else
				paragraph.ApplyNumPr(numId, ilvl, checkIndents);
		}
	};
	CNumberingApplicator.prototype.ApplyToHeadings = function(numId)
	{
		if (!numId || !this.NumInfo.Headings)
			return;

		let num          = this.Numbering.GetNum(numId);
		let styleManager = this.Document.GetStyleManager();
		if (num)
		{
			num.LinkWithHeadings(styleManager);

			for (let index = 0, count = this.Paragraphs.length; index < count; ++index)
			{
				let paragraph = this.Paragraphs[index];
				let numPr     = paragraph.GetNumPr();
				let iLvl      = numPr ? numPr.Lvl : 0;
				paragraph.SetNumPr(null);

				let pStyle = paragraph.GetParagraphStyle();
				if (!pStyle || -1 === styleManager.GetHeadingLevelById(pStyle))
					paragraph.SetParagraphStyleById(styleManager.GetDefaultHeading(iLvl));
			}

			this.Document.UpdateStylePanel();
		}
	};
	CNumberingApplicator.prototype.LinkNumWithHeadings = function(num)
	{
		if (!num || !this.NumInfo.Headings)
			return;
		
		let styleManager = this.Document.GetStyleManager();
		num.LinkWithHeadings(styleManager);
		this.Document.UpdateStylePanel();
	};
	CNumberingApplicator.prototype.UpdateDocumentOutline = function()
	{
		for (let index = 0, count = this.Paragraphs.length; index < count; ++index)
		{
			this.Paragraphs[index].UpdateDocumentOutline();
		}
	};
	CNumberingApplicator.prototype.MergePrToLvl = function(oldLvl, newLvl)
	{
		if (!oldLvl || !newLvl)
			return;
		
		let textPr = oldLvl.GetTextPr().Copy();
		textPr.Merge(newLvl.GetTextPr());
		textPr.RFonts = newLvl.GetTextPr().RFonts.Copy();
		newLvl.SetTextPr(textPr);
		
		let paraPr = oldLvl.GetParaPr().Copy();
		paraPr.Merge(newLvl.GetParaPr());
		newLvl.SetParaPr(paraPr);
		
		newLvl.SetPStyle(oldLvl.GetPStyle());
	};
	CNumberingApplicator.prototype.IsCheckIndents = function()
	{
		// TODO: Возможно, наоборот, надо тут false возвращать
		if (this.Paragraphs.length <= 1)
			return true;
		
		let paraPr = this.Paragraphs[0].Get_CompiledPr2(false).ParaPr;
		let left   = paraPr.Ind.Left;
		let first  = paraPr.Ind.FirstLine;
		
		for (let index = 1, count = this.Paragraphs.length; index < count; ++index)
		{
			let paraPr = this.Paragraphs[index].Get_CompiledPr2(false).ParaPr;
			if (Math.abs(left - paraPr.Ind.Left) > 0.001 || Math.abs(first - paraPr.Ind.FirstLine) > 0.001)
				return true;
		}
		
		return false;
	};
	//---------------------------------------------------------export---------------------------------------------------
	window["AscWord"].CNumberingApplicator = CNumberingApplicator;

})(window);
