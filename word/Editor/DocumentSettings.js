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
	const DEFAULT_HYPHENATION_ZONE = 360;
	
	/**
	 * Класс с глобальными настройками для документа
	 * @param {AscWord.CDocument} logicDocument
	 * @constructor
	 */
	function DocumentSettings(logicDocument)
	{
		this.LogicDocument = logicDocument;
		
		this.View                 = Asc.DocumentView.None;
		this.MathSettings         = new AscWord.MathSettings();
		this.CompatibilityMode    = AscCommon.document_compatibility_mode_Current;
		this.SdtSettings          = new AscWord.SdtGlobalSettings();
		this.SpecialFormsSettings = new AscWord.SpecialFormsGlobalSettings();
		this.WriteProtection      = undefined;
		this.DocumentProtection   = undefined !== AscCommonWord.CDocProtect && logicDocument ? new AscCommonWord.CDocProtect() : null;
		// TODO: Переделать AscCommonWord.CDocProtect. Класс с Id внутри класса без Id - очень плохо
		
		// Параметры, связанные с автоматической расстановкой переносов
		this.autoHyphenation        = undefined;
		this.hyphenationZone        = undefined;
		this.consecutiveHyphenLimit = undefined;
		this.doNotHyphenateCaps     = undefined;
		
		this.ListSeparator   = undefined;
		this.DecimalSymbol   = undefined;
		this.GutterAtTop     = false;
		this.MirrorMargins   = false;
		this.TrackRevisions  = false; // Флаг рецензирования, который записан в самом файле

		// Compatibility
		this.SplitPageBreakAndParaMark        = false;
		this.DoNotExpandShiftReturn           = false;
		this.BalanceSingleByteDoubleByteWidth = false;
		this.UlTrailSpace                     = false;
		this.UseFELayout                      = false;
	}
	DocumentSettings.prototype.getCompatibilityMode = function()
	{
		return this.CompatibilityMode;
	};
	DocumentSettings.prototype.isAutoHyphenation = function()
	{
		return !!this.autoHyphenation;
	};
	DocumentSettings.prototype.setAutoHyphenation = function(isAuto)
	{
		if (this.autoHyphenation === isAuto
			|| (!isAuto && undefined === this.autoHyphenation))
			return;
		
		AscCommon.AddAndExecuteChange(new CChangesDocumentSettingsAutoHyphenation(this.LogicDocument, this.autoHyphenation, isAuto));
	};
	DocumentSettings.prototype.isHyphenateCaps = function()
	{
		return (true !== this.doNotHyphenateCaps);
	};
	DocumentSettings.prototype.setHyphenateCaps = function(isHyphenate)
	{
		if (this.doNotHyphenateCaps === !isHyphenate
			|| (isHyphenate && undefined === this.doNotHyphenateCaps))
			return;
		
		AscCommon.AddAndExecuteChange(new CChangesDocumentSettingsDoNotHyphenateCaps(this.LogicDocument, this.doNotHyphenateCaps, !isHyphenate));
	};
	DocumentSettings.prototype.getConsecutiveHyphenLimit = function()
	{
		return !this.consecutiveHyphenLimit ? 0 : this.consecutiveHyphenLimit;
	};
	DocumentSettings.prototype.setConsecutiveHyphenLimit = function(limit)
	{
		if (this.consecutiveHyphenLimit === limit
			|| (undefined === this.consecutiveHyphenLimit && !limit))
			return;
		
		AscCommon.AddAndExecuteChange(new CChangesDocumentSettingsConsecutiveHyphenLimit(this.LogicDocument, this.consecutiveHyphenLimit, limit));
	};
	DocumentSettings.prototype.getHyphenationZone = function()
	{
		return undefined === this.hyphenationZone ? DEFAULT_HYPHENATION_ZONE : this.hyphenationZone;
	};
	DocumentSettings.prototype.setHyphenationZone = function(zone)
	{
		if (this.hyphenationZone === zone
			|| (undefined === this.hyphenationZone && DEFAULT_HYPHENATION_ZONE === zone))
			return;
		
		AscCommon.AddAndExecuteChange(new CChangesDocumentSettingsHyphenationZone(this.LogicDocument, this.hyphenationZone, zone));
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].DocumentSettings          = DocumentSettings;
	window['AscWord'].DEFAULT_DOCUMENT_SETTINGS = new DocumentSettings(null);
	window['AscWord'].DEFAULT_HYPHENATION_ZONE  = DEFAULT_HYPHENATION_ZONE;

})(window);
