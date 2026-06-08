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
	 * Document numbering collection
	 * @param logicDocument {AscWord.CDocument}
	 * @constructor
	 */
	function CNumberingCollection(logicDocument)
	{
		this.LogicDocument = logicDocument;
		this.Numbering     = logicDocument.GetNumbering();
		
		this.CheckParagraphs = {};
		this.NumToParagraph  = {};
		this.ParagraphToNum  = {};
		
		this.NeedRecollect = true;
		this.NeedUpdateUI  = true;
		
		this.BulletedCollection = {};
		this.NumberedCollection = {};
		this.MultiLvlCollection = {};
	}
	CNumberingCollection.prototype.CheckParagraph = function(paragraph)
	{
		if (!paragraph)
			return;
		
		this.CheckParagraphs[paragraph.GetId()] = paragraph;
		this.NeedRecollect = true;
		this.NeedUpdateUI  = true;
	};
	CNumberingCollection.prototype.CheckNum = function(numId, iLvl)
	{
		this.NeedUpdateUI = true;
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
	CNumberingCollection.prototype.UpdatePanel = function()
	{
		if (!this.NeedUpdateUI)
			return;
		
		this.NeedUpdateUI = false;
		
		this.InitCollections();
		this.LogicDocument.GetApi().sendEvent('asc_updateListPatterns', this.GetCollections());
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
		
		// С этими включенными флагами мы не можем до конца рассчитать нумерацию у параграфа
		if (AscCommon.g_oIdCounter.m_bLoad || AscCommon.g_oIdCounter.m_bRead)
			return;
		
		this.NeedRecollect = false;
		
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
				
				empty = AscCommon.isEmptyObject(this.NumToParagraph[numId][iLvl]);
				
				if (!empty)
					break;
			}
			
			if (empty)
				delete this.NumToParagraph[numId];
		}
	};
	CNumberingCollection.prototype.InitCollections = function()
	{
		this.Recollect();
		
		this.BulletedCollection = {};
		this.NumberedCollection = {};
		this.MultiLvlCollection = {};
		
		for (let numId in this.NumToParagraph)
		{
			for (let iLvl = 0; iLvl < 9; ++iLvl)
			{
				if (!this.NumToParagraph[numId][iLvl])
					continue;
				
				if (!AscCommon.isEmptyObject(this.NumToParagraph[numId][iLvl]))
					this.AddNumToCollections(numId, iLvl);
			}
		}
	};
	CNumberingCollection.prototype.GetCollections = function()
	{
		return {
			"singleBullet"    : Object.keys(this.BulletedCollection),
			"singleNumbering" : Object.keys(this.NumberedCollection),
			"multiLevel"      : Object.keys(this.MultiLvlCollection)
		};
	};
	CNumberingCollection.prototype.AddNumToCollections = function(numId, iLvl)
	{
		let num = this.Numbering.GetNum(numId);
		if (!num)
			return;

		this.AddToMultiLvlCollection(num);
		this.AddToSingleLvlCollection(num, iLvl);
	};
	CNumberingCollection.prototype.AddToSingleLvlCollection = function(num, iLvl)
	{
		let numLvl = num.GetLvl(iLvl);
		if (Asc.c_oAscNumberingFormat.None === numLvl.GetFormat())
			return;
		
		let numInfo = AscWord.CNumInfo.FromNum(num, iLvl, this.LogicDocument.GetStyles());
		let json    = JSON.stringify(numInfo.ToJson());
		
		if (numInfo.IsNumbered())
			this.NumberedCollection[json] = true;
		
		if (numInfo.IsBulleted())
			this.BulletedCollection[json] = true;
	};
	CNumberingCollection.prototype.AddToMultiLvlCollection = function(num)
	{
		let numInfo = AscWord.CNumInfo.FromNum(num, null, this.LogicDocument.GetStyles());
		let json    = JSON.stringify(numInfo.ToJson());
		
		this.MultiLvlCollection[json] = true;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CNumberingCollection = CNumberingCollection;
	
})(window);

