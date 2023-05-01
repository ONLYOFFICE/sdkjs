/*
 * (c) Copyright Ascensio System SIA 2010-2023
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

(function(window)
{
	/**
	 * Document numbering collection
	 * @param logicDocument {AscWord.CDocument}
	 * @constructor
	 */
	function CNumberingCollection(logicDocument)
	{
		this.LogicDocument     = logicDocument;
		this.Numbering         = logicDocument.GetNumbering();
		this.singleBullet      = {};
		this.singleNumbering   = {};
		this.multiLevel        = {};
		this.checkNumMap       = {};
		this.checkSingleLvlMap = {};
		
		this.CheckParagraphs = {};
		this.NumToParagraph = {};
		this.ParagraphToNum = {};
		this.NeedRecollect  = true;
	}
	CNumberingCollection.prototype.Init = function()
	{
		let allParagraphs = this.LogicDocument.GetAllParagraphs();
		for (let paraIndex = 0, paraCount = allParagraphs.length; paraIndex < paraCount; ++paraIndex)
		{
			let numPr = allParagraphs[paraIndex].GetNumPr();
			if (numPr && numPr.IsValid())
				this.AddNum(numPr);
		}
	};
	CNumberingCollection.prototype.GetCollections = function()
	{
		return {
			"singleBullet"    : Object.keys(this.singleBullet),
			"singleNumbering" : Object.keys(this.singleNumbering),
			"multiLevel"      : Object.keys(this.multiLevel)
		};
	};
	CNumberingCollection.prototype.CheckParagraph = function(paragraph)
	{
		if (!paragraph)
			return;
		
		this.CheckParagraphs[paragraph.GetId()] = paragraph;
		this.NeedRecollect = true;
	};
	CNumberingCollection.prototype.GetAllParagraphsByNumbering = function(value)
	{
		this.Recollect();
		
		let result = [];
		if (!value)
		{
			return [];
		}
		else if (undefined !== value.NumId)
		{
			let numPr = value;
			this.GetAllParagraphsByNumPr(result, numPr.NumId, numPr.Lvl);
		}
		else if (Array.isArray(value))
		{
			for (let index = 0, count = value.length; index < count; ++index)
			{
				let numPr = value[index];
				this.GetAllParagraphsByNumPr(result, numPr.NumId, numPr.Lvl);
			}
		}
		
		return result;
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CNumberingCollection.prototype.GetAllParagraphsByNumPr = function(result, numId, iLvl)
	{
		if (!this.NumToParagraph[numId])
			return;
		
		if (undefined === iLvl || null === iLvl)
			iLvl = -1;
		
		if (-1 === iLvl)
		{
			for (iLvl = 0; iLvl < 9; ++iLvl)
			{
				if (!this.NumToParagraph[numId][iLvl])
					continue;
				
				for (let paraId in this.NumToParagraph[numId][iLvl])
					result.push(this.NumToParagraph[numId][iLvl][paraId]);
			}
		}
		else
		{
			if (this.NumToParagraph[numId][iLvl])
			{
				for (let paraId in this.NumToParagraph[numId][iLvl])
					result.push(this.NumToParagraph[numId][iLvl][paraId]);
			}
		}
	};
	CNumberingCollection.prototype.Recollect = function()
	{
		if (!this.NeedRecollect)
			return;
		
		this.NeedRecollect = false;
		
		console.time("Numbering.Recollect");
		
		let numToCheck = {};
		for (let paraId in this.CheckParagraphs)
		{
			let paragraph = this.CheckParagraphs[paraId];
			delete this.CheckParagraphs[paraId];
			
			if (this.ParagraphToNum[paraId])
			{
				let oldNumPr = this.ParagraphToNum[paraId];
				delete this.ParagraphToNum[paraId];
				delete this.NumToParagraph[oldNumPr.NumId][oldNumPr.Lvl][paraId];
				
				numToCheck[oldNumPr.NumId] = true;
			}
			
			if (!paragraph.IsUseInDocument())
				continue;

			let numPr = paragraph.GetNumPr();
			if (numPr && numPr.IsValid())
			{
				this.ParagraphToNum[paraId] = numPr.Copy();
				if (!this.NumToParagraph[numPr.NumId])
				{
					this.NumToParagraph[numPr.NumId] = new Array(9);
					for (let iLvl = 0; iLvl < 9; ++iLvl)
						this.NumToParagraph[numPr.NumId][iLvl] = {};
				}
				
				this.NumToParagraph[numPr.NumId][numPr.Lvl][paraId] = paragraph;
			}
		}
		
		this.ClearEmptyNumToParagraph(numToCheck);
		
		console.timeEnd("Numbering.Recollect");
	};
	CNumberingCollection.prototype.ClearEmptyNumToParagraph = function(numToCheck)
	{
		for (let numId in numToCheck)
		{
			if (!this.NumToParagraph[numId])
				continue;
			
			let empty = true;
			for (let iLvl = 0; iLvl < 9; ++iLvl)
			{
				if (!this.NumToParagraph[numId][iLvl])
					continue;
				
				for (let paraId in this.NumToParagraph[numId][iLvl])
				{
					empty = false;
					break;
				}
				
				if (!empty)
					break;
			}
			
			if (empty)
				delete this.NumToParagraph[numId];
		}
	};
	CNumberingCollection.prototype.AddNum = function(oNumPr)
	{
		const sNumId = oNumPr.NumId;
		const nLvl   = oNumPr.Lvl;
		
		if (!this.checkNumMap[sNumId])
		{
			const oNum = this.Numbering.GetNum(sNumId);
			if (oNum)
			{
				this.CheckMultiLvl(oNum);
			}
			this.checkNumMap[sNumId]       = true;
			this.checkSingleLvlMap[sNumId] = {};
		}
		
		if (!this.checkSingleLvlMap[sNumId][nLvl])
		{
			const oNum = this.Numbering.GetNum(sNumId);
			if (oNum)
			{
				this.CheckSingleLvl(oNum, nLvl);
				this.checkSingleLvlMap[sNumId][nLvl] = true;
			}
			else
			{
				for (let i = 0; i < 9; i += 1)
				{
					this.checkSingleLvlMap[sNumId][i] = true;
				}
			}
		}
	};
	CNumberingCollection.prototype.CheckSingleLvl = function(num, iLvl)
	{
		let numInfo = AscWord.CNumInfo.FromNum(num, iLvl, this.LogicDocument.GetStyles());
		if (numInfo.IsNumbered())
			this.AddToSingleNumbered(JSON.stringify(numInfo.ToJson()));
		else if (numInfo.IsBulleted())
			this.AddToSingleBullet(JSON.stringify(numInfo.ToJson()));
	};
	CNumberingCollection.prototype.CheckMultiLvl = function(num)
	{
		let numInfo = AscWord.CNumInfo.FromNum(num, null, this.LogicDocument.GetStyles());
		this.AddToMultiLvl(JSON.stringify(numInfo.ToJson()));
	};
	CNumberingCollection.prototype.AddToMultiLvl = function (sJSON)
	{
		this.multiLevel[sJSON] = true;
	};
	CNumberingCollection.prototype.AddToSingleBullet = function (sJSON)
	{
		this.singleBullet[sJSON] = true;
	};
	CNumberingCollection.prototype.AddToSingleNumbered = function (sJSON)
	{
		this.singleNumbering[sJSON] = true;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CNumberingCollection = CNumberingCollection;
	
})(window);

