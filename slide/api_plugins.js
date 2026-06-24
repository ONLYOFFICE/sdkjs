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
	 * @typedef {Object} comment
	 * Comment object.
	 * @property {string} Id - The comment ID.
	 * @property {CommentData} Data - An object which contains the comment data.
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/comment.js
	 */
	
    /**
	 * @typedef {Object} CommentData
	 * The comment data.
	 * @property {string} UserName - The comment author.
	 * @property {string} QuoteText - The quote comment text.
	 * @property {string} Text - The comment text.
	 * @property {string} Time - The time when the comment was posted (in milliseconds).
	 * @property {string} UserId - The user ID of the comment author.
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
	 * Removes the OLE object from the presentation by its internal ID.
	 * @memberof Api
	 * @typeofeditors ["CPE"]
	 * @alias RemoveOleObject
	 * @param {string} internalId - The OLE object identifier which is used to work with OLE object added to the presentation.
	 * @since 9.1.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/RemoveOleObject.js
	 */
	Api.prototype["pluginMethod_RemoveOleObject"] = function (internalId) {
		let logicDocument = this.WordControl.m_oLogicDocument;
		if (logicDocument) {
			logicDocument.RemoveDrawingObjectById(internalId);
		}
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
		if (!AscFormat.isRealNumber(nSlideIndex) || nSlideIndex <= 0) {
			return;
		}
		let nTargetIdx = nSlideIndex - 1;
		let oDemo = this.getDemoManager();
		if (oDemo && oDemo.SlideNum === nTargetIdx) {
			return;
		}
		this.DemonstrationGoToSlide(nTargetIdx);
	};

	/**
	 * Sets the active slide in the editor by its index.
	 * @memberof Api
	 * @typeofeditors ["CPE"]
	 * @alias GoToSlide
	 * @param {number} slideNumber - The slide index.
	 * @since 9.0.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Api/Methods/GoToSlide.js
	 */
	Api.prototype["pluginMethod_GoToSlide"] = function (slideNumber) {
		if (!AscFormat.isRealNumber(slideNumber) || slideNumber <= 0) {
			return;
		}
		let nTargetIdx = slideNumber - 1;
		if (this.getCurrentPage() === nTargetIdx) {
			return;
		}
		this.goToPage(nTargetIdx);
	};

})(window);


