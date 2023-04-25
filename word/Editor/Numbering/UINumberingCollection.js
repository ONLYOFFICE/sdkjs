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
	 * Document numbering collection for UI
	 * @constructor
	 */
	function UINumberingCollection()
	{
		this.singleBullet      = {};
		this.singleNumbering   = {};
		this.multiLevel        = {};
		this.checkNumMap       = {};
		this.checkSingleLvlMap = {};
	}
	/**
	 * @param logicDocument {AscWord.CDocument}
	 */
	UINumberingCollection.prototype.Init = function(logicDocument)
	{
		let allParagraphs = logicDocument.GetAllParagraphs();
		let numbering     = logicDocument.GetNumbering();
		for (let paraIndex = 0, paraCount = allParagraphs.length; paraIndex < paraCount; ++paraIndex)
		{
			let numPr = allParagraphs[paraIndex].GetNumPr();
			if (numPr && numPr.IsValid())
				this.AddNum(numPr, numbering);
		}
	};
	UINumberingCollection.prototype.GetCollections = function()
	{
		return {
			"singleBullet"    : Object.keys(this.singleBullet),
			"singleNumbering" : Object.keys(this.singleNumbering),
			"multiLevel"      : Object.keys(this.multiLevel)
		};
	};
	UINumberingCollection.prototype.AddNum = function(oNumPr, numbering)
	{
		const sNumId = oNumPr.NumId;
		const nLvl   = oNumPr.Lvl;
		
		if (!this.checkNumMap[sNumId])
		{
			const oNum = numbering.GetNum(sNumId);
			if (oNum)
			{
				this.CheckMultiLvl(oNum);
			}
			this.checkNumMap[sNumId]       = true;
			this.checkSingleLvlMap[sNumId] = {};
		}
		
		if (!this.checkSingleLvlMap[sNumId][nLvl])
		{
			const oNum = numbering.GetNum(sNumId);
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
	UINumberingCollection.prototype.CheckSingleLvl = function (oNum, nLvl)
	{
		const oJSON = oNum.GetJSONNumbering(true, nLvl);
		if (oJSON["Type"] === Asc.c_oAscJSONNumberingType.Number)
		{
			this.AddToSingleNumbered(JSON.stringify(oJSON));
		}
		else if (oJSON["Type"] === Asc.c_oAscJSONNumberingType.Bullet)
		{
			this.AddToSingleBullet(JSON.stringify(oJSON));
		}
	};
	UINumberingCollection.prototype.CheckMultiLvl = function (oNum)
	{
		this.AddToMultiLvl(JSON.stringify(oNum.GetJSONNumbering()));
	};
	UINumberingCollection.prototype.AddToMultiLvl = function (sJSON)
	{
		this.multiLevel[sJSON] = true;
	};
	UINumberingCollection.prototype.AddToSingleBullet = function (sJSON)
	{
		this.singleBullet[sJSON] = true;
	};
	UINumberingCollection.prototype.AddToSingleNumbered = function (sJSON)
	{
		this.singleNumbering[sJSON] = true;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].UINumberingCollection = UINumberingCollection;
	
})(window);

