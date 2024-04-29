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
	 * Класс для информации о нумерации
	 * @param numInfo
	 * @constructor
	 */
	function CNumInfo(numInfo)
	{
		this.Type     = numInfo && numInfo["Type"] ? numInfo["Type"] : "";
		this.Lvl      = numInfo && numInfo["Lvl"] && numInfo["Lvl"].length ? numInfo["Lvl"] : [];
		this.Headings = numInfo && numInfo["Headings"] ? numInfo["Headings"] : false;
	}
	CNumInfo.Parse = function(value)
	{
		if (value instanceof CNumInfo)
			return value;
		
		let numInfo = null;
		if (typeof value === "string" || value instanceof String)
		{
			try
			{
				numInfo = CNumInfo.FromJson(JSON.parse(value));
			}
			catch (e)
			{
				return null;
			}
		}
		else if (value instanceof Object)
		{
			numInfo = CNumInfo.FromJson(value);
		}
		
		return numInfo;
	};
	CNumInfo.FromJson = function(json)
	{
		return new CNumInfo(json);
	};
	CNumInfo.FromLvl = function(lvl, iLvl, styles)
	{
		let numInfo = new CNumInfo();
		if (undefined === iLvl || null === iLvl)
		{
			let isBulleted = false;
			let isNumbered = false;
			let isHeading  = true;
			
			for (let iLvl = 0; iLvl < 9; ++iLvl)
			{
				let numLvl = lvl[iLvl];
				isBulleted = isBulleted || numLvl.IsBulleted();
				isNumbered = isNumbered || numLvl.IsNumbered();
				
				numInfo.Lvl[iLvl] = numLvl.ToJson(null);
				if (isHeading && styles)
				{
					if (numLvl.GetPStyle() !== styles.GetDefaultHeading(iLvl))
						isHeading = false;
				}
			}
			
			if (isHeading)
				numInfo.Headings = true;
			
			if (isBulleted && isNumbered)
				numInfo.Type = Asc.c_oAscJSONNumberingType.Hybrid;
			else if (isNumbered)
				numInfo.Type = Asc.c_oAscJSONNumberingType.Number;
			else if (isBulleted)
				numInfo.Type = Asc.c_oAscJSONNumberingType.Bullet;
		}
		else
		{
			let numLvl = lvl[iLvl];
			if (numLvl.GetRelatedLvlList().length <= 1)
			{
				numLvl = numLvl.Copy();
				numLvl.ResetNumberedText(0);
				numInfo.Type   = numLvl.IsBulleted() ? Asc.c_oAscJSONNumberingType.Bullet : Asc.c_oAscJSONNumberingType.Number;
				numInfo.Lvl[0] = numLvl.ToJson(null, {isSingleLvlPresetJSON: true});
			}
		}
		
		return numInfo;
	};
	CNumInfo.FromNum = function(num, iLvl, styles)
	{
		let lvl = null;
		if (num instanceof AscWord.CNum)
		{
			lvl = [];
			for (let index = 0; index < 9; ++index)
				lvl.push(num.GetLvl(index));
		}
		else if (num instanceof Asc.CAscNumbering)
		{
			lvl = [];
			for (let index = 0; index < 9; ++index)
			{
				let numLvl = new CNumberingLvl();
				numLvl.FillFromAscNumberingLvl(num.get_Lvl(index));
				lvl.push(numLvl);
			}
		}
		
		if (!lvl)
			return null;
		
		return CNumInfo.FromLvl(lvl, iLvl, styles);
	};
	CNumInfo.prototype.IsEqual = function(numInfo)
	{
		if (!numInfo
			|| this.Type !== numInfo.Type
			|| this.Headings !== numInfo.Headings
			|| this.Lvl.length !== numInfo.Lvl.length)
			return false;
		
		if (1 === this.Lvl.length)
		{
			let numLvl1 = AscWord.CNumberingLvl.FromJson(this.Lvl[0]);
			let numLvl2 = AscWord.CNumberingLvl.FromJson(numInfo.Lvl[0]);
			numLvl1.ResetNumberedText(0);
			numLvl2.ResetNumberedText(0);
			return numLvl1.IsSimilar(numLvl2);
		}
		else
		{
			for (let iLvl = 0; iLvl < 9; ++iLvl)
			{
				if (!this.Lvl[iLvl] || !numInfo.Lvl[iLvl])
					return false;
				
				let numLvl1 = AscWord.CNumberingLvl.FromJson(this.Lvl[iLvl]);
				let numLvl2 = AscWord.CNumberingLvl.FromJson(numInfo.Lvl[iLvl]);
				
				if (!numLvl1.IsEqual(numLvl2))
					return false;
			}
			return true;
		}
	};
	CNumInfo.prototype.CompareWithNum = function(num, iLvl)
	{
		if (!num)
			return false;
		
		if (null !== iLvl && undefined !== iLvl && -1 !== iLvl)
		{
			if (!this.Lvl.length || !this.Lvl[0])
				return false;
			
			let numLvl = AscWord.CNumberingLvl.FromJson(this.Lvl[0]);
			numLvl.ResetNumberedText(iLvl);
			return numLvl.IsSimilar(num.GetLvl(iLvl));
		}
		else
		{
			if (this.Lvl.length < 9)
				return false;
			
			for (let iLvl = 0; iLvl < 9; ++iLvl)
			{
				if (!this.Lvl[iLvl])
					return false;
				
				let numLvl = AscWord.CNumberingLvl.FromJson(this.Lvl[iLvl]);
				if (!numLvl.IsEqual(num.GetLvl(iLvl)))
					return false;
			}
			
			return true;
		}
	};
	CNumInfo.prototype.IsNumbered = function()
	{
		return this.Type === Asc.c_oAscJSONNumberingType.Number;
	};
	CNumInfo.prototype.IsBulleted = function()
	{
		return this.Type === Asc.c_oAscJSONNumberingType.Bullet;
	};
	CNumInfo.prototype.IsHeadings = function()
	{
		return this.Headings;
	};
	CNumInfo.prototype.IsRemove = function()
	{
		return this.Type === Asc.c_oAscJSONNumberingType.Remove;
	};
	CNumInfo.prototype.IsNone = function()
	{
		return this.IsRemove();
	};
	CNumInfo.prototype.HaveLvl = function()
	{
		return (!!this.Lvl.length);
	};
	CNumInfo.prototype.GetType = function()
	{
		return this.Type;
	};
	CNumInfo.prototype.ToJson = function()
	{
		let json = {
			"Type" : this.Type,
			"Lvl"  : this.Lvl.slice()
		};
		
		if (this.IsHeadings())
			json["Headings"] = true;
		
		return json;
	};
	CNumInfo.prototype.IsSingleLevel = function()
	{
		return (1 === this.Lvl.length);
	};
	CNumInfo.prototype.FillNum = function(num)
	{
		if (!num)
			return;
		
		if (num instanceof Asc.CAscNumbering)
		{
			for (let iLvl = 0; iLvl < 9; ++iLvl)
			{
				let numInfoLvl = this.Lvl[iLvl];
				let numLvl;
				if (numInfoLvl)
					numLvl = AscWord.CNumberingLvl.FromJson(numInfoLvl);
				else
					numLvl = AscWord.CNumberingLvl.CreateDefault(iLvl, this.IsBulleted() ? Asc.c_oAscMultiLevelNumbering.Bullet : Asc.c_oAscMultiLevelNumbering.Numbered);
				
				numLvl.FillToAscNumberingLvl(num.get_Lvl(iLvl));
			}
		}
		else if (num instanceof AscWord.CNum)
		{
			for (let iLvl = 0; iLvl < 9; ++iLvl)
			{
				let numInfoLvl = this.Lvl[iLvl];
				let numLvl;
				if (numInfoLvl)
					numLvl = AscWord.CNumberingLvl.FromJson(numInfoLvl);
				else
					numLvl = AscWord.CNumberingLvl.CreateDefault(iLvl, this.IsBulleted() ? Asc.c_oAscMultiLevelNumbering.Bullet : Asc.c_oAscMultiLevelNumbering.Numbered);
				
				num.SetLvl(numLvl, iLvl);
			}
		}
	};
	
	//---------------------------------------------------------export---------------------------------------------------
	window["AscWord"].CNumInfo = CNumInfo;
	
})(window);
