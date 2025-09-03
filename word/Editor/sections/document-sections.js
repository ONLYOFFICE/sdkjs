/*
 * (c) Copyright Ascensio System SIA 2010-2025
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
	 * Class for managing document sections and their header/footer content.
	 *
	 * @param {AscWord.Document} logicDocument - The parent document object
	 * @constructor
	 */
	function DocumentSections(logicDocument)
	{
		this.logicDocument = logicDocument;
		this.Elements = [];
		
		this.needUpdate = true;
		
		this.paragraphToCheck = {};
		
		this.paraToSection = {};
	}
	DocumentSections.prototype.GetFirstSectPr = function()
	{
		this.Update();
		
		if (!this.Elements.length)
			return this.logicDocument.GetFinalSectPr();
		
		return this.Elements[0].SectPr;
	};
	/**
	 * @returns {AscWord.SectPr}
	 */
	DocumentSections.prototype.GetNextSectPr = function(sectPr)
	{
		this.Update();
		
		let index = this.Find(sectPr);
		if (-1 === index || index >= this.Elements.length - 1)
			return sectPr;
		
		return this.Elements[index + 1].SectPr;
	};
	DocumentSections.prototype.GetSectionsCount = function()
	{
		this.Update();
		return this.Elements.length;
	};
	DocumentSections.prototype.Find_ByHdrFtr = function(HdrFtr)
	{
		if (!HdrFtr)
			return -1;
		
		this.Update();
		
		var Count = this.Elements.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var SectPr = this.Elements[Index].SectPr;
			
			if (HdrFtr === SectPr.Get_Header_First()
				|| HdrFtr === SectPr.Get_Header_Default()
				|| HdrFtr === SectPr.Get_Header_Even()
				|| HdrFtr === SectPr.Get_Footer_First()
				|| HdrFtr === SectPr.Get_Footer_Default()
				|| HdrFtr === SectPr.Get_Footer_Even())
				return Index;
		}
		
		return -1;
	};
	DocumentSections.prototype.Reset_HdrFtrRecalculateCache = function()
	{
		this.forEachHdrFtr(function(hdrFtr){
			hdrFtr.Reset_RecalculateCache();
		});
	};
	DocumentSections.prototype.GetAllParagraphs = function(props, paragraphs)
	{
		if (!paragraphs)
			paragraphs = [];
		
		this.forEachHdrFtr(function(hdrFtr){
			hdrFtr.GetAllParagraphs(props, paragraphs);
		});
		
		return paragraphs;
	};
	DocumentSections.prototype.GetAllTables = function(props, tables)
	{
		if (!tables)
			tables = [];
		
		this.forEachHdrFtr(function(hdrFtr){
			hdrFtr.GetAllTables(props, tables);
		});
		
		return tables;
	};
	DocumentSections.prototype.GetAllDrawingObjects = function(drawings)
	{
		if (!drawings)
			drawings = [];
		
		this.forEachHdrFtr(function(hdrFtr){
			hdrFtr.GetAllDrawingObjects(drawings);
		});
		
		return drawings;
	};
	DocumentSections.prototype.UpdateBookmarks = function(bookmarkManager)
	{
		this.forEachHdrFtr(function(hdrFtr){
			hdrFtr.UpdateBookmarks(bookmarkManager);
		});
	};
	DocumentSections.prototype.Document_CreateFontMap = function(fontMap)
	{
		this.forEachHdrFtr(function(hdrFtr){
			hdrFtr.Document_CreateFontMap(fontMap);
		});
	};
	DocumentSections.prototype.Document_CreateFontCharMap = function(fontCharMap)
	{
		this.forEachHdrFtr(function(hdrFtr){
			hdrFtr.Document_CreateFontCharMap(fontCharMap);
		});
	};
	DocumentSections.prototype.Document_Get_AllFontNames = function(allFonts)
	{
		this.forEachHdrFtr(function(hdrFtr){
			hdrFtr.Document_Get_AllFontNames(allFonts);
		});
	};
	DocumentSections.prototype.GetSectPrByElement = function(element)
	{
		this.Update();
		
		let finalSectPr = this.logicDocument.GetFinalSectPr();
		if (!element)
			return finalSectPr;
		
		let docPos = element.GetDocumentPositionFromObject();
		if (!docPos || !docPos.length)
			return finalSectPr;
		
		let topClass = docPos[0].Class;
		if (topClass !== this.logicDocument)
		{
			let sectPr = topClass.Get_SectPr();
			return sectPr ? sectPr : finalSectPr;
		}
		
		let b = 0;
		let e = this.Elements.length - 1;
		
		while (b < e)
		{
			let m = (b + e) >> 1;
			let _docPos = this.Elements[m].Paragraph.GetDocumentPositionFromObject();
			if (AscWord.CompareDocumentPositions(docPos, _docPos) <= 0)
				e = m;
			else
				b = m + 1;
		}
		
		return this.Elements[b].SectPr;
	};
	DocumentSections.prototype.GetFirstParagraph = function(sectionIndex)
	{
		if (0 === sectionIndex)
			return this.logicDocument.GetFirstParagraph();
		
		let prevParagraph = this.Elements[sectionIndex - 1].Paragraph;
		return prevParagraph.GetNextParagraph();
	};
	DocumentSections.prototype.GetSectPrByIndex = function(index)
	{
		return this.Elements[index].SectPr;
	};
	DocumentSections.prototype.Find = function(sectPr)
	{
		if (!sectPr)
			return -1;
		
		this.Update();
		
		for (let i = 0, count = this.Elements.length; i < count; ++i)
		{
			if (this.Elements[i].SectPr === sectPr)
				return i;
		}
		
		return -1;
	};
	DocumentSections.prototype.Update = function()
	{
		if (!this.CheckNeedUpdate())
			return;
		
		this._clear();
		
		let paragraphs = this.logicDocument.GetAllSectPrParagraphs();
		for (let i = 0; i < paragraphs.length; ++i)
		{
			this._addSection(paragraphs[i].Get_SectionPr(), paragraphs[i]);
		}
		this._addSection(this.logicDocument.GetFinalSectPr(), this.logicDocument);
		
		// Когда полностью обновляются секции надо пересчитывать с самого начала
		this.logicDocument.RecalcInfo.Set_NeedRecalculateFromStart(true);
		
		this.needUpdate = false;
	};
	DocumentSections.prototype.CheckNeedUpdate = function()
	{
		for (let paraId in this.paragraphToCheck)
		{
			let para = this.paragraphToCheck[paraId];
			if (!para.Get_SectionPr() || para.CanAddSectionPr())
			{
				this.needUpdate = true;
				break;
			}
		}
		
		return this.needUpdate;
	};
	DocumentSections.prototype.CheckParagraph = function(paragraph)
	{
		if (paragraph && paragraph.Get_SectionPr())
			this.paragraphToCheck[paragraph.GetId()] = paragraph;
	};
	DocumentSections.prototype.UpdateOnAdd = function(items)
	{
		this.Update();
		
		for (let i = 0, count = items.length; i < count; ++i)
		{
			let paragraphs = items[i].GetAllSectPrParagraphs();
			for (let paraIndex = 0, paraCount = paragraphs.length; paraIndex < paraCount; ++paraIndex)
			{
				let paragraph = paragraphs[paraIndex];
				
				let paraId = paragraph.GetId();
				let sectPr = paragraph.Get_SectionPr();
				if (!sectPr || this.paraToSection[paraId])
					continue;
				
				let sectionPos = this.GetIndexByElement(paragraph);
				if (-1 === sectionPos)
					continue;
				
				this._insertSection(sectionPos, sectPr, paragraph);
			}
		}
	};
	DocumentSections.prototype.UpdateOnRemove = function(items, checkHdrFtr)
	{
		this.Update();
		
		let paragraphs = [];
		for (let i = 0, count = items.length; i < count; ++i)
		{
			items[i].GetAllSectPrParagraphs(paragraphs);
		}
		
		// По логике удаляемые секции должны идти последовательно, но будем проверять каждую отдельно
		for (let i = 0, count = paragraphs.length; i < count; ++i)
		{
			let sectionPos = this.GetIndexByElement(paragraphs[i]);
			if (-1 === sectionPos)
				continue;
			
			if (checkHdrFtr && sectionPos < this.Elements.length - 1)
			{
				let currSectPr = this.Elements[sectionPos].SectPr;
				let nextSectPr = this.Elements[sectionPos + 1].SectPr;
				if (nextSectPr.IsAllHdrFtrNull() && !currSectPr.IsAllHdrFtrNull())
				{
					nextSectPr.Set_Header_First(currSectPr.Get_Header_First());
					nextSectPr.Set_Header_Even(currSectPr.Get_Header_Even());
					nextSectPr.Set_Header_Default(currSectPr.Get_Header_Default());
					nextSectPr.Set_Footer_First(currSectPr.Get_Footer_First());
					nextSectPr.Set_Footer_Even(currSectPr.Get_Footer_Even());
					nextSectPr.Set_Footer_Default(currSectPr.Get_Footer_Default());
				}
			}
			
			this._removeSection(sectionPos);
		}
	};
	DocumentSections.prototype.GetCount = function()
	{
		this.Update();
		return this.Elements.length;
	};
	/**
	 * Получаем секцию по заданному номеру
	 * @param {number} nIndex
	 * @returns {DocumentSection}
	 */
	DocumentSections.prototype.Get = function(nIndex)
	{
		this.Update();
		return this.Elements[nIndex];
	};
	/**
	 * Получаем номер секции, в которой лежит заданный элемент
	 * @param element
	 * @returns {number}
	 */
	DocumentSections.prototype.GetIndexByElement = function(element)
	{
		this.Update();
		
		let b = 0;
		let e = this.Elements.length - 1;
		
		if (!element)
			return -1;
		
		let docPos = element.GetDocumentPositionFromObject();
		if (!docPos || !docPos.length)
			return -1;
		
		let topClass = docPos[0].Class;
		if (topClass !== this.logicDocument)
		{
			let sectPr = topClass.Get_SectPr();
			return this.Find(sectPr);
		}
		
		while (b < e)
		{
			let m = (b + e) >> 1;
			let _docPos = this.Elements[m].Paragraph.GetDocumentPositionFromObject();
			if (AscWord.CompareDocumentPositions(docPos, _docPos) <= 0)
				e = m;
			else
				b = m + 1;
		}
		
		return b;
	};
	/**
	 * Получаем массив всех колонтитулов, используемых в данном документе
	 * @returns {Array.CHeaderFooter}
	 */
	DocumentSections.prototype.GetAllHdrFtrs = function()
	{
		let result = [];
		this.forEachHdrFtr(function(hdrFtr){
			result.push(hdrFtr);
		});
		return result;
	};
	DocumentSections.prototype.GetAllContentControls = function(contentControls)
	{
		this.forEachHdrFtr(function(hdrFtr){
			hdrFtr.GetAllContentControls(contentControls);
		});
	};
	/**
	 * Обновляем заданную секцию
	 * @param paragraph {AscWord.Paragraph} - Параграф, в котором обновляется секция
	 * @param prevSectPr {AscWord.SectPr} - Предыдущая секция в заданном параграфе (если она была)
	 * @param checkHdrFtr {boolean} - Нужно ли проверять колонтитулы при удалении секции
	 * @returns {boolean} Если не смогли обновить, возвращаем false
	 */
	DocumentSections.prototype.UpdateSection = function(paragraph, prevSectPr, checkHdrFtr)
	{
		this.Update();
		
		let newSectPr = paragraph.Get_SectionPr();
		if (prevSectPr)
		{
			let sectionPos = this.Find(prevSectPr);
			if (-1 === sectionPos || this.paraToSection[paragraph.GetId()] !== this.Elements[sectionPos])
				return;
			
			if (newSectPr)
			{
				this.Elements[sectionPos].SectPr = newSectPr;
			}
			else
			{
				if (checkHdrFtr && sectionPos < this.Elements.length - 1)
				{
					let currSectPr = this.Elements[sectionPos].SectPr;
					let nextSectPr = this.Elements[sectionPos + 1].SectPr;
					if (nextSectPr.IsAllHdrFtrNull() && !currSectPr.IsAllHdrFtrNull())
					{
						nextSectPr.Set_Header_First(currSectPr.Get_Header_First());
						nextSectPr.Set_Header_Even(currSectPr.Get_Header_Even());
						nextSectPr.Set_Header_Default(currSectPr.Get_Header_Default());
						nextSectPr.Set_Footer_First(currSectPr.Get_Footer_First());
						nextSectPr.Set_Footer_Even(currSectPr.Get_Footer_Even());
						nextSectPr.Set_Footer_Default(currSectPr.Get_Footer_Default());
					}
				}
				this._removeSection(sectionPos);
			}
		}
		else if (newSectPr)
		{
			if (this.paraToSection[paragraph.GetId()])
			{
				this.paraToSection[paragraph.GetId()].SectPr = newSectPr;
			}
			else
			{
				let sectionPos = this.GetIndexByElement(paragraph);
				if (-1 === sectionPos)
					return;
				
				this._insertSection(sectionPos, newSectPr, paragraph)
			}
		}
	};
	DocumentSections.prototype.private_GetHdrFtrsArray = function(oCurHdrFtr)
	{
		var isEvenOdd = EvenAndOddHeaders;
		
		var nCurPos    = -1;
		var arrHdrFtrs = [];
		for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			var oSectPr = this.Elements[nIndex].SectPr;
			var isFirst = oSectPr.Get_TitlePage();
			
			var oHeaderFirst   = oSectPr.Get_Header_First();
			var oHeaderEven    = oSectPr.Get_Header_Even();
			var oHeaderDefault = oSectPr.Get_Header_Default();
			var oFooterFirst   = oSectPr.Get_Footer_First();
			var oFooterEven    = oSectPr.Get_Footer_Even();
			var oFooterDefault = oSectPr.Get_Footer_Default();
			
			if (oHeaderFirst && isFirst)
				arrHdrFtrs.push(oHeaderFirst);
			
			if (oHeaderEven && isEvenOdd)
				arrHdrFtrs.push(oHeaderEven);
			
			if (oHeaderDefault)
				arrHdrFtrs.push(oHeaderDefault);
			
			if (oFooterFirst && isFirst)
				arrHdrFtrs.push(oFooterFirst);
			
			if (oFooterEven && isEvenOdd)
				arrHdrFtrs.push(oFooterEven);
			
			if (oFooterDefault)
				arrHdrFtrs.push(oFooterDefault);
		}
		
		if (oCurHdrFtr)
		{
			for (var nIndex = 0, nCount = arrHdrFtrs.length; nIndex < nCount; ++nIndex)
			{
				if (oCurHdrFtr === arrHdrFtrs[nIndex])
				{
					nCurPos = nIndex;
					break;
				}
			}
		}
		
		return {
			HdrFtrs : arrHdrFtrs,
			CurPos  : nCurPos
		};
	};
	DocumentSections.prototype.FindNextFillingForm = function(isNext, oCurHdrFtr)
	{
		var oInfo = this.private_GetHdrFtrsArray(oCurHdrFtr);
		
		var arrHdrFtrs = oInfo.HdrFtrs;
		var nCurPos    = oInfo.CurPos;
		
		var nCount = arrHdrFtrs.length;
		
		var isCurrent = true;
		if (-1 === nCurPos)
		{
			isCurrent = false;
			nCurPos   = isNext ? 0 : arrHdrFtrs.length - 1;
			if (arrHdrFtrs[nCurPos])
				oCurHdrFtr = arrHdrFtrs[nCurPos];
		}
		
		if (nCurPos >= 0 && nCurPos <= nCount - 1)
		{
			var oRes = oCurHdrFtr.GetContent().FindNextFillingForm(isNext, isCurrent, isCurrent);
			if (oRes)
				return oRes;
			
			if (isNext)
			{
				for (var nIndex = nCurPos + 1; nIndex < nCount; ++nIndex)
				{
					oRes = arrHdrFtrs[nIndex].GetContent().FindNextFillingForm(isNext, false);
					
					if (oRes)
						return oRes;
				}
			}
			else
			{
				for (var nIndex = nCurPos - 1; nIndex >= 0; --nIndex)
				{
					oRes = arrHdrFtrs[nIndex].GetContent().FindNextFillingForm(isNext, false);
					
					if (oRes)
						return oRes;
				}
			}
		}
		
		return null;
	};
	DocumentSections.prototype.RestartSpellCheck = function()
	{
		var bEvenOdd = EvenAndOddHeaders;
		for (let nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			var SectPr = this.Elements[nIndex].SectPr;
			var bFirst = SectPr.Get_TitlePage();
			
			if (null != SectPr.HeaderFirst && true === bFirst)
				SectPr.HeaderFirst.RestartSpellCheck();
			
			if (null != SectPr.HeaderEven && true === bEvenOdd)
				SectPr.HeaderEven.RestartSpellCheck();
			
			if (null != SectPr.HeaderDefault)
				SectPr.HeaderDefault.RestartSpellCheck();
			
			if (null != SectPr.FooterFirst && true === bFirst)
				SectPr.FooterFirst.RestartSpellCheck();
			
			if (null != SectPr.FooterEven && true === bEvenOdd)
				SectPr.FooterEven.RestartSpellCheck();
			
			if (null != SectPr.FooterDefault)
				SectPr.FooterDefault.RestartSpellCheck();
		}
	};
	DocumentSections.prototype.RemoveEmptyHdrFtrs = function()
	{
		for (let nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			let oSectPr = this.Elements[nIndex].SectPr;
			oSectPr.RemoveEmptyHdrFtrs();
		}
	};
	DocumentSections.prototype.CheckRunContent = function(fCheck)
	{
		let headers = this.GetAllHdrFtrs();
		for (let index = 0, count = headers.length; index < count; ++index)
		{
			headers[index].GetContent().CheckRunContent(fCheck);
		}
	};
	//------------------------------------------------------------------------------------------------------------------
	// Search
	//------------------------------------------------------------------------------------------------------------------
	DocumentSections.prototype.Search = function(oSearchEngine)
	{
		var bEvenOdd = EvenAndOddHeaders;
		for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
		{
			var oSectPr = this.Elements[nIndex].SectPr;
			var bFirst  = oSectPr.Get_TitlePage();
			
			if (oSectPr.HeaderFirst && true === bFirst)
				oSectPr.HeaderFirst.Search(oSearchEngine, search_Header);
			
			if (oSectPr.HeaderEven && true === bEvenOdd)
				oSectPr.HeaderEven.Search(oSearchEngine, search_Header);
			
			if (oSectPr.HeaderDefault)
				oSectPr.HeaderDefault.Search(oSearchEngine, search_Header);
			
			if (oSectPr.FooterFirst && true === bFirst)
				oSectPr.FooterFirst.Search(oSearchEngine, search_Footer);
			
			if (oSectPr.FooterEven && true === bEvenOdd)
				oSectPr.FooterEven.Search(oSearchEngine, search_Footer);
			
			if (oSectPr.FooterDefault)
				oSectPr.FooterDefault.Search(oSearchEngine, search_Footer);
		}
	};
	DocumentSections.prototype.GetSearchElementId = function(bNext, CurHdrFtr)
	{
		var HdrFtrs = [];
		var CurPos  = -1;
		
		var bEvenOdd = EvenAndOddHeaders;
		var Count    = this.Elements.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var SectPr = this.Elements[Index].SectPr;
			var bFirst = SectPr.Get_TitlePage();
			
			if (null != SectPr.HeaderFirst && true === bFirst)
			{
				HdrFtrs.push(SectPr.HeaderFirst);
				
				if (CurHdrFtr === SectPr.HeaderFirst)
					CurPos = HdrFtrs.length - 1;
			}
			
			if (null != SectPr.HeaderEven && true === bEvenOdd)
			{
				HdrFtrs.push(SectPr.HeaderEven);
				
				if (CurHdrFtr === SectPr.HeaderEven)
					CurPos = HdrFtrs.length - 1;
			}
			
			if (null != SectPr.HeaderDefault)
			{
				HdrFtrs.push(SectPr.HeaderDefault);
				
				if (CurHdrFtr === SectPr.HeaderDefault)
					CurPos = HdrFtrs.length - 1;
			}
			
			if (null != SectPr.FooterFirst && true === bFirst)
			{
				HdrFtrs.push(SectPr.FooterFirst);
				
				if (CurHdrFtr === SectPr.FooterFirst)
					CurPos = HdrFtrs.length - 1;
			}
			
			if (null != SectPr.FooterEven && true === bEvenOdd)
			{
				HdrFtrs.push(SectPr.FooterEven);
				
				if (CurHdrFtr === SectPr.FooterEven)
					CurPos = HdrFtrs.length - 1;
			}
			
			if (null != SectPr.FooterDefault)
			{
				HdrFtrs.push(SectPr.FooterDefault);
				
				if (CurHdrFtr === SectPr.FooterDefault)
					CurPos = HdrFtrs.length - 1;
			}
		}
		
		var Count = HdrFtrs.length;
		
		var isCurrent = true;
		if (-1 === CurPos)
		{
			isCurrent = false;
			CurPos    = bNext ? 0 : HdrFtrs.length - 1;
			if (HdrFtrs[CurPos])
				CurHdrFtr = HdrFtrs[CurPos];
		}
		
		if (CurPos >= 0 && CurPos <= HdrFtrs.length - 1)
		{
			var Id = CurHdrFtr.GetSearchElementId(bNext, isCurrent);
			if (null != Id)
				return Id;
			
			if (true === bNext)
			{
				for (var Index = CurPos + 1; Index < Count; Index++)
				{
					Id = HdrFtrs[Index].GetSearchElementId(bNext, false);
					
					if (null != Id)
						return Id;
				}
			}
			else
			{
				for (var Index = CurPos - 1; Index >= 0; Index--)
				{
					Id = HdrFtrs[Index].GetSearchElementId(bNext, false);
					
					if (null != Id)
						return Id;
				}
			}
		}
		
		return null;
	};
	//------------------------------------------------------------------------------------------------------------------
	// Private area
	//------------------------------------------------------------------------------------------------------------------
	/**
	 * Iterates through all sections and executes a callback function for each header and footer.
	 * @param {Function} callback
	 */
	DocumentSections.prototype.forEachHdrFtr = function(callback)
	{
		this.Update();
		
		for (let i = 0, count = this.Elements.length; i < count; ++i)
		{
			let sectPr = this.Elements[i].SectPr;
			
			if (null != sectPr.HeaderFirst)
				callback.call(this, sectPr.HeaderFirst);
			
			if (null != sectPr.HeaderDefault)
				callback.call(this, sectPr.HeaderDefault);
			
			if (null != sectPr.HeaderEven)
				callback.call(this, sectPr.HeaderEven);
			
			if (null != sectPr.FooterFirst)
				callback.call(this, sectPr.FooterFirst);
			
			if (null != sectPr.FooterDefault)
				callback.call(this, sectPr.FooterDefault);
			
			if (null != sectPr.FooterEven)
				callback.call(this, sectPr.FooterEven);
		}
	};
	DocumentSections.prototype._clear = function()
	{
		//console.log(`Clear all sections`);
		this.Elements.length = 0;
		this.paraToSection = {};
	};
	DocumentSections.prototype._insertSection = function(pos, sectPr, paragraph)
	{
		//console.log(`Insert sections at pos ${pos} paraId=${paragraph.GetId()}`);
		let documentSection = new DocumentSection(sectPr, paragraph);
		this.paraToSection[paragraph.GetId()] = documentSection;
		this.Elements.splice(pos, 0, documentSection);
	};
	DocumentSections.prototype._addSection = function(sectPr, paragraph)
	{
		let documentSection = new DocumentSection(sectPr, paragraph);
		this.paraToSection[paragraph.GetId()] = documentSection;
		this.Elements.push(documentSection);
		//console.log(`Add section at the end. Current count=${this.Elements.length} paraId=${paragraph.GetId()}`);
	};
	DocumentSections.prototype._removeSection = function(pos)
	{
		let docSection = this.Elements[pos];
		//console.log(`Remove section at pos ${pos} paraId=${docSection.Paragraph.GetId()}`);
		delete this.paraToSection[docSection.Paragraph.GetId()];
		this.Elements.splice(pos, 1);
	};
	//------------------------------------------------------------------------------------------------------------------
	
	/**
	 * Represents a document section associated with a specific paragraph.
	 *
	 * @param {AscWord.SectPr} sectPr - Section properties object containing formatting settings
	 * @param {AscWord.Paragraph} paragraph - The paragraph object that ends this section
	 * @constructor
	 */
	function DocumentSection(sectPr, paragraph)
	{
		this.SectPr    = sectPr;
		this.Paragraph = paragraph;
	}
	
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.DocumentSections = DocumentSections;
	AscWord.DocumentSection  = DocumentSection;
})();
