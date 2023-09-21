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

(function(window, undefined)
{
    /**
     * Base class.
     * @global
     * @class
     * @name Api
     */

    var Api = window["Asc"]["spreadsheet_api"];

    /**
	 * @typedef {Object} CommentData
	 * The comment data.
	 * @property {string} UserName - The comment author.
	 * @property {string} Text - The comment text.
	 * @property {string} Time - The time when the comment was posted (in milliseconds).
	 * @property {boolean} Solved - Specifies if the comment is resolved (**true**) or not (**false**).
	 * @property {CommentData[]} Replies - An array containing the comment replies represented as the *CommentData* object.
	 */

	/**
	 * Adds a comment to the workbook.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @alias AddComment
	 * @param {CommentData}  oCommentData - An object which contains the comment data.
	 * @return {string | null} - The comment ID in the string format or null if the comment cannot be added.
	 * @since 7.3.0
	 */
	Api.prototype["pluginMethod_AddComment"] = function(oCommentData)
	{
        if (!oCommentData)
            return null;
            
        let oAscCD = new Asc.asc_CCommentData();
		oAscCD.ReadFromSimpleObject(oCommentData);

		this.asc_addComment(oAscCD);

		if (oAscCD.nId)
			return oAscCD.nId;

		return null;
	};

	/**
	 * Changes the specified comment.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @alias ChangeComment
	 * @param {string} sId - The comment ID.
	 * @param {CommentData} oCommentData - An object which contains the new comment data.
	 * @return {boolean}
	 * @since 7.3.0
	 */
	Api.prototype["pluginMethod_ChangeComment"] = function(sId, oCommentData)
	{
		if (!oCommentData)
			return false;
		var oSourceComm = this.wb.cellCommentator.findComment(sId);
		if (!oSourceComm)
			return false;

		let oAscCD = oSourceComm;
		oAscCD.ReadFromSimpleObject(oCommentData);

		this.asc_changeComment(sId, oAscCD);
		return true;
	};

	/**
	 * Removes the specified comments.
	 * @param {string[]} arrIds - An array which contains the IDs of the specified comments.
	 * @memberof Api
	 * @typeofeditors ["CSE"]
	 * @alias RemoveComments
	 * @since 7.3.0
	 */
	Api.prototype["pluginMethod_RemoveComments"] = function(arrIds)
	{
		for (let comm in arrIds)
		{
			if (arrIds.hasOwnProperty(comm))
			{
				this.asc_removeComment(arrIds[comm]);
			}
		}

	};

})(window);
