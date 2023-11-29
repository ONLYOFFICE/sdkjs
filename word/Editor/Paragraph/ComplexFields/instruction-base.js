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
	const FLAGS_MERGEFORMAT = 0x00000001;
	
	/**
	 * Base class for complex field instruction
	 * @constructor
	 */
	function FieldInstructionBase()
	{
		this.ComplexField    = null;
		this.InstructionLine = "";
		this.generalSwitches = 0x0000;
	}
	FieldInstructionBase.prototype.Type = AscWord.fieldtype_UNKNOWN;
	FieldInstructionBase.prototype.GetType = function()
	{
		return this.Type;
	};
	FieldInstructionBase.prototype.SetComplexField = function(oComplexField)
	{
		this.ComplexField = oComplexField;
	};
	FieldInstructionBase.prototype.GetComplexField = function()
	{
		return this.ComplexField;
	};
	FieldInstructionBase.prototype.ToString = function()
	{
		let result = this.writeField();
		result += this.writeGeneralSwitches();
		return result;
	};
	FieldInstructionBase.prototype.SetPr = function()
	{
	};
	FieldInstructionBase.prototype.SetInstructionLine = function(sLine)
	{
		this.InstructionLine = sLine;
	};
	FieldInstructionBase.prototype.CheckInstructionLine = function(sLine)
	{
		return (this.InstructionLine === sLine);
	};
	FieldInstructionBase.prototype.addGeneralSwitches = function(switches)
	{
		if (!switches || !switches.length)
			return;
		
		for (let i = 0; i < switches.length; ++i)
		{
			let curSwitch = switches[i].toUpperCase();
			if ("MERGEFORMAT" === curSwitch)
				this.generalSwitches |= FLAGS_MERGEFORMAT;
		}
	};
	FieldInstructionBase.prototype.isMergeFormat = function()
	{
		return !!(this.generalSwitches & FLAGS_MERGEFORMAT);
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	FieldInstructionBase.prototype.writeField = function()
	{
		// This method must be overridden
	};
	FieldInstructionBase.prototype.writeGeneralSwitches = function()
	{
		let result = "";
		
		if (this.GeneralSwitches & FLAGS_MERGEFORMAT)
			result += " \\* MERGEFORMAT";
		
		return result;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].FieldInstructionBase = FieldInstructionBase;
	
})(window);
