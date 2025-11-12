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
	 * Класс представляющий номер сноски внутри сноски.
	 * @param {CFootEndnote} Footnote - Ссылка на сноску.
	 * @constructor
	 * @extends {AscWord.CRunFootnoteReference}
	 */
	function CRunFootnoteRef(Footnote)
	{
		AscWord.CRunFootnoteReference.call(this, Footnote);
	}
	CRunFootnoteRef.prototype = Object.create(AscWord.CRunFootnoteReference.prototype);
	CRunFootnoteRef.prototype.constructor = CRunFootnoteRef;

	CRunFootnoteRef.prototype.Type = para_FootnoteRef;
	CRunFootnoteRef.prototype.Get_Type = function()
	{
		return para_FootnoteRef;
	};
	CRunFootnoteRef.prototype.Copy = function(oPr)
	{
		var oFootnote = this.GetFootnote();
		var oParagraph, oParent, oTopDocument;
		if(oPr && oPr.Paragraph)
		{
			oParagraph = oPr.Paragraph;
			oParent = oParagraph.GetParent();
			if(oParent)
			{
				oTopDocument = oParent.GetTopDocumentContent();
				if(oTopDocument && oTopDocument instanceof CFootEndnote)
				{
					oFootnote = oTopDocument;
				}
			}
		}
		return new CRunFootnoteRef(oFootnote);
	};
	CRunFootnoteRef.prototype.UpdateNumber = function(oFootnote)
	{
		this.Footnote = oFootnote;
		if (this.Footnote && this.Footnote instanceof CFootEndnote)
		{
			this.Number    = this.Footnote.GetNumber();
			this.NumFormat = this.Footnote.GetReferenceSectPr().GetFootnoteNumFormat();
			this.private_Measure();
		}
		else
		{
			this.Number    = 1;
			this.NumFormat = Asc.c_oAscNumberingFormat.Decimal;
			this.private_Measure();
		}
	};
	CRunFootnoteRef.prototype.PreDelete = function()
	{
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunFootnoteRef = CRunFootnoteRef;

})(window);
