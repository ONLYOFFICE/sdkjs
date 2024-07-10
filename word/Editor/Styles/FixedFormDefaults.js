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
	const NONE_BORDER = new CDocumentBorder();
	NONE_BORDER.Value = border_None;

	const NONE_SHD = new CDocumentShd();
	NONE_SHD.Value = Asc.c_oAscShd.Nil;

	// NumId=0 означает отсутствие нумерации
	const NONE_NUM = new AscWord.NumPr(0, 0);

	const TEXTFORM_PR = new CParaPr();
	const CHECKBOX_PR = new CParaPr();

	TEXTFORM_PR.Set_FromObject({
		Ind : {
			Left      : 0,
			Right     : 0,
			FirstLine : 0
		},

		Jc : AscCommon.align_Left,

		Spacing : {
			Line              : 1,
			LineRule          : Asc.linerule_Auto,
			Before            : 0,
			BeforeAutoSpacing : false,
			After             : 0,
			AfterAutoSpacing  : false
		},

		Shd : NONE_SHD,

		Brd : {
			Between : NONE_BORDER,
			Bottom  : NONE_BORDER,
			Left    : NONE_BORDER,
			Right   : NONE_BORDER,
			Top     : NONE_BORDER
		},

		NumPr : NONE_NUM
	});

	CHECKBOX_PR.Set_FromObject({
		Ind : {
			Left      : 0,
			Right     : 0,
			FirstLine : 0
		},

		Jc : AscCommon.align_Center,

		Spacing : {
			Line              : 1,
			LineRule          : Asc.linerule_Auto,
			Before            : 0,
			BeforeAutoSpacing : false,
			After             : 0,
			AfterAutoSpacing  : false
		},

		Shd : NONE_SHD,

		Brd : {
			Between : NONE_BORDER,
			Bottom  : NONE_BORDER,
			Left    : NONE_BORDER,
			Right   : NONE_BORDER,
			Top     : NONE_BORDER
		},

		NumPr : NONE_NUM
	});

	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].DEFAULT_PARAPR_FIXED_TEXTFORM     = TEXTFORM_PR;
	window['AscWord'].DEFAULT_PARAPR_FIXED_CHECKBOXFORM = CHECKBOX_PR;

})(window);
