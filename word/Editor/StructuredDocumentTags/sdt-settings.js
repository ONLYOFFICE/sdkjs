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
