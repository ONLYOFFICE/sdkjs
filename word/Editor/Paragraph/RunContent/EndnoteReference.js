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
