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
	 * Класс представляющий ссылку на сноску
	 * @param {CFootEndnote} oEndnote - Ссылка на сноску
	 * @param {string} sCustomMark
	 * @constructor
	 * @extends {AscWord.CRunFootnoteReference}
	 */
	function CRunEndnoteReference(oEndnote, sCustomMark)
	{
		AscWord.CRunFootnoteReference.call(this, oEndnote, sCustomMark);
	}
	CRunEndnoteReference.prototype = Object.create(AscWord.CRunFootnoteReference.prototype);
	CRunEndnoteReference.prototype.constructor = CRunEndnoteReference;

	CRunEndnoteReference.prototype.Type = para_EndnoteReference;
	CRunEndnoteReference.prototype.Get_Type = function()
	{
		return para_EndnoteReference;
	};
	CRunEndnoteReference.prototype.Copy = function(oPr)
	{
		if (this.Footnote)
		{
			var oEndnote;
			if (oPr && oPr.Comparison)
			{
				oEndnote = oPr.Comparison.createEndNote();
			}
			else
			{
				oEndnote = this.Footnote.Parent.CreateEndnote();
			}
			oEndnote.Copy2(this.Footnote, oPr);
		}

		var oRef = new CRunEndnoteReference(oEndnote);

		oRef.Number    = this.Number;
		oRef.NumFormat = this.NumFormat;

		return oRef;
	};
	CRunEndnoteReference.prototype.UpdateNumber = function(PRS, isKeepNumber)
	{
		if (this.Footnote && true !== PRS.IsFastRecalculate() && PRS.TopDocument instanceof CDocument)
		{
			var nPageAbs    = PRS.GetPageAbs();
			var nColumnAbs  = PRS.GetColumnAbs();
			var nNumber     = PRS.GetEndnoteReferenceNumber(this);
			var oSectPr     = PRS.GetSectPr();
			var nNumFormat  = oSectPr.GetEndnoteNumFormat();

			var oLogicDocument      = this.Footnote.GetLogicDocument();
			var oEndnotesController = oLogicDocument.GetEndnotesController();

			if (!isKeepNumber)
			{
				this.NumFormat = nNumFormat;
				this.Number    = -1 === nNumber ? oEndnotesController.GetEndnoteNumberOnPage(nPageAbs, nColumnAbs, oSectPr, this.Footnote) : nNumber;

				// Если данная сноска не участвует в нумерации, просто уменьшаем ей номер на 1, для упрощения работы
				if (this.IsCustomMarkFollows())
					this.Number--;
			}

			this.private_Measure();
			this.Footnote.SetNumber(this.Number, oSectPr, this.IsCustomMarkFollows());
		}
		else
		{
			this.Number    = 1;
			this.NumFormat = Asc.c_oAscNumberingFormat.Decimal;
			this.private_Measure();
		}
	};
	CRunEndnoteReference.prototype.ToSearchElement = function(oProps)
	{
		return new AscCommonWord.CSearchTextSpecialEndnoteMark();
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunEndnoteReference = CRunEndnoteReference;

})(window);
