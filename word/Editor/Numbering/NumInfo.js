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
				
				numInfo.Lvl[iLvl] = numLvl;
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
				numInfo.Lvl[0] = numLvl;
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
			"Lvl"  : []
		};
		
		if (this.Lvl.length)
		{
			for (let iLvl = 0; iLvl < this.Lvl.length; ++iLvl)
			{
				json["Lvl"][iLvl] = this.Lvl[iLvl].ToJson(null, {isSingleLvlPresetJSON: true});
			}
		}
		
		if (this.IsHeadings())
			json["Headings"] = true;
		
		return json;
	};
	//---------------------------------------------------------export---------------------------------------------------
	window["AscWord"].CNumInfo = CNumInfo;
	
})(window);
