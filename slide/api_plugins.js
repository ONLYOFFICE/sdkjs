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

(function(window, undefined)
{
    /**
     * Base class.
     * @global
     * @class
     * @name Api
     */

    var Api = window["asc_docs_api"];

    /**
	 * @typedef {Object} CommentData
	 * The comment data.
	 * @property {string} UserName - The comment author.
	 * @property {string} QuoteText - The quote comment text.
	 * @property {string} Text - The comment text.
	 * @property {string} Time - The time when the comment was posted (in milliseconds).
	 * @property {boolean} Solved - Specifies if the comment is resolved (**true**) or not (**false**).
	 * @property {CommentData[]} Replies - An array containing the comment replies represented as the *CommentData* object.
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/CommentData.js
	 */

	/**
	 * Adds a comment to the presentation.
	 * @memberof Api
	 * @typeofeditors ["CPE"]
	 * @alias AddComment
	 * @param {CommentData}  oCommentData - An object which contains the comment data.
	 * @return {string | null} - The comment ID in the string format or null if the comment cannot be added.
	 * @since 7.3.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/AddComment.js
	 */
	Api.prototype["pluginMethod_AddComment"] = function(oCommentData)
	{
        if (!oCommentData)
            return null;
            
        let oCD = new AscCommon.CCommentData();
		oCD.ReadFromSimpleObject(oCommentData);
		oCD.m_sGuid = AscCommon.CreateGUID();

        let Comment = this.WordControl.m_oLogicDocument.AddComment(oCD, true);
		if (Comment) {
			return Comment.Get_Id();
		}
		return null;
	};

	/**
	 * Changes the specified comment.
	 * @memberof Api
	 * @typeofeditors ["CPE"]
	 * @alias ChangeComment
	 * @param {string} sId - The comment ID.
	 * @param {CommentData} oCommentData - An object which contains the new comment data.
	 * @return {boolean}
	 * @since 7.3.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/ChangeComment.js
	 */
	Api.prototype["pluginMethod_ChangeComment"] = function(sId, oCommentData)
	{
		if (!oCommentData)
			return false;
		var oSourceComm = g_oTableId.Get_ById(sId);
		if (!oSourceComm || !oSourceComm.Parent)
			return false;

		var oCD = oSourceComm.Data;
		oCD.ReadFromSimpleObject(oCommentData);

		this.WordControl.m_oLogicDocument.EditComment(sId, oCD);
		return true;
	};

	/**
	 * Removes the specified comments.
	 * @param {string[]} arrIds - An array which contains the IDs of the specified comments.
	 * @memberof Api
	 * @typeofeditors ["CPE"]
	 * @alias RemoveComments
	 * @since 7.3.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/RemoveComments.js
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

	/**
	 * Returns all the comments from the document.
	 * @memberof Api
	 * @typeofeditors ["CPE"]
	 * @alias GetAllComments
	 * @returns {comment[]} - An array of comment objects containing the comment data.
	 * @since 8.1.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GetAllComments.js
	 */
	Api.prototype["pluginMethod_GetAllComments"] = function()
	{
		const oLogicDocument = this.WordControl.m_oLogicDocument;
		if (!oLogicDocument)
			return;

		const arrResult = [];

		const oComments = oLogicDocument.GetAllComments();
		for (let index = 0; index < oComments.length; index++)
		{
			const oComment = oComments[index].comment;
			arrResult.push({"Id" : oComment.GetId(), "Data" : oComment.GetData().ConvertToSimpleObject()});
		}

		return arrResult;
	};

	/**
	 * Starts the presentation slide show.
	 * @memberof Api
	 * @typeofeditors ["CPE"]
	 * @alias StartSlideShow
	 * @since 8.0.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/StartSlideShow.js
	 */
	Api.prototype["pluginMethod_StartSlideShow"] = function()
	{
		this.sendEvent("asc_onStartDemonstration");
	};

	/**
	 * Pauses the current slide show.
	 * @memberof Api
	 * @typeofeditors ["CPE"]
	 * @alias PauseSlideShow
	 * @since 8.0.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/PauseSlideShow.js
	 */
	Api.prototype["pluginMethod_PauseSlideShow"] = function()
	{
		this.WordControl.DemonstrationManager.Pause();
		this.sendEvent("asc_onDemonstrationStatus", "pause");
	};
	/**
	 * Resumes the current slide show.
	 * @memberof Api
	 * @typeofeditors ["CPE"]
	 * @alias ResumeSlideShow
	 * @since 8.0.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/ResumeSlideShow.js
	 */
	Api.prototype["pluginMethod_ResumeSlideShow"] = function()
	{
		this.WordControl.DemonstrationManager.Play();
		this.sendEvent("asc_onDemonstrationStatus", "play");
	};


	/**
	 * Ends the current slide show.
	 * @memberof Api
	 * @typeofeditors ["CPE"]
	 * @alias EndSlideShow
	 * @since 8.0.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/EndSlideShow.js
	 */
	Api.prototype["pluginMethod_EndSlideShow"] = function()
	{
		this.EndDemonstration();
	};

	/**
	 * Displays the slide following the current slide in the slide show.
	 * @memberof Api
	 * @typeofeditors ["CPE"]
	 * @alias GoToNextSlideInSlideShow
	 * @since 8.0.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GoToNextSlideInSlideShow.js
	 */
	Api.prototype["pluginMethod_GoToNextSlideInSlideShow"] = function()
	{
		this.DemonstrationNextSlide();
	};

	/**
	 * Displays the slide following the current slide in the slide show.
	 * @memberof Api
	 * @typeofeditors ["CPE"]
	 * @alias GoToPreviousSlideInSlideShow
	 * @since 8.0.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GoToPreviousSlideInSlideShow.js
	 */
	Api.prototype["pluginMethod_GoToPreviousSlideInSlideShow"] = function()
	{
		this.DemonstrationPrevSlide();
	};

	/**
	 * Displays the slide with the specific index.
	 * @memberof Api
	 * @typeofeditors ["CPE"]
	 * @alias GoToSlideInSlideShow
	 * @param {number} nSlideIndex - The slide index.
	 * @since 8.0.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GoToSlideInSlideShow.js
	 */
	Api.prototype["pluginMethod_GoToSlideInSlideShow"] = function(nSlideIndex)
	{
		this.DemonstrationGoToSlide(nSlideIndex - 1);
	};

})(window);


