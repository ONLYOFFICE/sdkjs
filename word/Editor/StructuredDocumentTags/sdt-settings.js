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
	 * Класс с глобальными настройками для всех контейнеров
	 * @constructor
	 */
	function SdtGlobalSettings()
	{
		this.Color         = new AscCommonWord.CDocumentColor(220, 220, 220);
		this.ShowHighlight = false;
	}
	/**
	 * Проверяем все ли параметры выставлены по умолчанию
	 * @returns {boolean}
	 */
	SdtGlobalSettings.prototype.IsDefault = function()
	{
		if (!this.Color
			|| 220 !== this.Color.r
			|| 220 !== this.Color.g
			|| 220 !== this.Color.b
			|| false !== this.ShowHighlight)
			return false;
		
		return true;
	};
	SdtGlobalSettings.prototype.Copy = function()
	{
		var oSettings = new SdtGlobalSettings();
		
		oSettings.Color         = this.Color.Copy();
		oSettings.ShowHighlight = this.ShowHighlight;
		
		return oSettings;
	};
	SdtGlobalSettings.prototype.Write_ToBinary = function(oWriter)
	{
		// CDocumentColor : Color
		// Bool           : ShowHighlight
		
		this.Color.WriteToBinary(oWriter);
		oWriter.WriteBool(this.ShowHighlight);
	};
	SdtGlobalSettings.prototype.Read_FromBinary = function(oReader)
	{
		this.Color.ReadFromBinary(oReader);
		this.ShowHighlight = oReader.GetBool();
	};
	
	/**
	 * Класс с глобальными настройками для всех специальных форм
	 * @constructor
	 */
	function SpecialFormsGlobalSettings()
	{
		this.Highlight = new AscCommonWord.CDocumentColor(201, 200, 255);
	}
	SpecialFormsGlobalSettings.prototype.Copy = function()
	{
		var oSettings = new SpecialFormsGlobalSettings();
		
		if (this.Highlight)
			oSettings.Highlight = this.Highlight.Copy();
		
		return oSettings;
	};
	/**
	 * Проверяем все ли параметры выставлены по умолчанию
	 * @returns {boolean}
	 */
	SpecialFormsGlobalSettings.prototype.IsDefault = function()
	{
		return (undefined === this.Highlight || (!this.Highlight.IsAuto() && this.Highlight.IsEqualRGB({r : 201, g: 200, b : 255})));
	};
	SpecialFormsGlobalSettings.prototype.Write_ToBinary = function(oWriter)
	{
		var nStartPos = oWriter.GetCurPosition();
		oWriter.Skip(4);
		var nFlags = 0;
		
		if (undefined !== this.Highlight)
		{
			this.Highlight.WriteToBinary(oWriter);
			nFlags |= 1;
		}
		
		var nEndPos = oWriter.GetCurPosition();
		oWriter.Seek(nStartPos);
		oWriter.WriteLong(nFlags);
		oWriter.Seek(nEndPos);
	};
	SpecialFormsGlobalSettings.prototype.Read_FromBinary = function(oReader)
	{
		var nFlags = oReader.GetLong();
		
		if (nFlags & 1)
		{
			this.Highlight = new AscCommonWord.CDocumentColor();
			this.Highlight.ReadFromBinary(oReader);
		}
		else
		{
			this.Highlight = undefined;
		}
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].SdtGlobalSettings          = SdtGlobalSettings;
	window['AscWord'].SpecialFormsGlobalSettings = SpecialFormsGlobalSettings;
	
})(window);
