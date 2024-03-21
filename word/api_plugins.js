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

    /**
     * @typedef {Object} ContentControl
	 * Content control object.
     * @property {string} Tag - A tag assigned to the content control. The same tag can be assigned to several content controls so that it is possible to make reference to them in your code.
     * @property {string} Id - A unique identifier of the content control. It can be used to search for a certain content control and make reference to it in the code.
     * @property {ContentControlLock} Lock - A value that defines if it is possible to delete and/or edit the content control or not: 0 - only deleting, 1 - no deleting or editing, 2 - only editing, 3 - full access.
     * @property {string} InternalId - A unique internal identifier of the content control. It is used for all operations with content controls.
     */

    /**
     * @typedef {(0 | 1 | 2 | 3)} ContentControlLock
     * A value that defines if it is possible to delete and/or edit the content control or not:
	 * * <b>0</b> - only deleting
	 * * <b>1</b> - disable deleting or editing
	 * * <b>2</b> - only editing
	 * * <b>3</b> - full access
     */

    /**
     * @typedef {(1 | 2 | 3 | 4)} ContentControlType
     * A numeric value that specifies the content control type:
	 * * <b>1</b> - block content control
	 * * <b>2</b> - inline content control
	 * * <b>3</b> - row content control
	 * * <b>4</b> - cell content control
     */

    /**
     * @typedef {Object} ContentControlPropertiesAndContent
     * The content control properties and contents.
     * @property  {ContentControlProperties} [ContentControlProperties = {}] - The content control properties.
     * @property  {string} Script - A script that will be executed to generate the data within the content control (can be replaced with the *Url* parameter).
     * @property  {string} Url - A link to the shared file (can be replaced with the *Script* parameter).
     */

    /**
     * @typedef {Object} ContentControlProperties
	 * The content control properties.
     * @property {string} Id - A unique identifier of the content control. It can be used to search for a certain content control and make reference to it in the code.
     * @property {string} Tag - A tag assigned to the content control. The same tag can be assigned to several content controls so that it is possible to make reference to them in the code.
     * @property {ContentControlLock} Lock - A value that defines if it is possible to delete and/or edit the content control or not.
     * @property {string} InternalId - A unique internal identifier of the content control.
	 * @property {string} Alias - The alias attribute.
	 * @property {string} PlaceHolderText - The content control placeholder text.
     * @property {number} Appearance - Defines if the content control is shown as the bounding box (**1**) or not (**2**).
     * @property {object} Color - The color for the current content control in the RGB format.
     * @property {number} Color.R - Red color component value.
     * @property {number} Color.G - Green color component value.
     * @property {number} Color.B - Blue color component value.
     * @example
     * {"Id": 100, "Tag": "CC_Tag", "Lock": 3}
     */
	
	/**
	 * @typedef {('none' | 'comments' | 'forms' | 'readOnly')} DocumentEditingRestrictions
	 * The document editing restrictions:
	 * * <b>none</b> - no editing restrictions,
	 * * <b>comments</b> - allows editing comments,
	 * * <b>forms</b> - allows editing form fields,
	 * * <b>readOnly</b> - does not allow editing.
	 */
	
	/**
	 * @typedef {("entirely" | "beforeCursor" | "afterCursor")} TextPartType
	 * Specifies if the whole text or only its part will be returned or replaced:
	 * * <b>entirely</b> - replaces/returns the whole text,
	 * * <b>beforeCursor</b> - replaces/returns only the part of the text before the cursor,
	 * * <b>afterCursor</b> - replaces/returns only the part of the text after the cursor.
	 */

    var Api = window["asc_docs_api"];

    /**
     * Opens a file with fields.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias OpenFile
     * @param {Uint8Array} binaryFile - A file in the format of the 8-bit unsigned integer array.
     * @param {string[]} fields - A list of field values.
     */
    window["asc_docs_api"].prototype["pluginMethod_OpenFile"] = function(binaryFile, fields)
    {
        this.asc_CloseFile();

        this.FontLoader.IsLoadDocumentFonts2 = true;
        this.OpenDocumentFromBin(this.DocumentUrl, binaryFile);

        if (fields)
            this.asc_SetBlockChainData(fields);

        this.restrictions = Asc.c_oAscRestrictionType.OnlyForms;
    };
    /**
     * Returns all fields as a text.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias GetFields
     * @returns {string[]} - A list of field values.
     */
    window["asc_docs_api"].prototype["pluginMethod_GetFields"] = function()
    {
        return this.asc_GetBlockChainData();
    };
    /**
     * Inserts the content control containing data. The data is specified by the JS code for {@link /docbuilder/basic Document Builder}, or by a link to the shared document.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias InsertAndReplaceContentControls
     * @param {ContentControlPropertiesAndContent[]} arrDocuments - An array of properties and contents of the content control.
     * @return {ContentControlProperties[]} - An array of created content control properties.
     * @example
     * // Add new content control
     * var arrDocuments = [{
     *  "Props": {
     *       "Id": 100,
     *       "Tag": "CC_Tag",
     *       "Lock": 3
     *   },
     *   "Script": "var oParagraph = Api.CreateParagraph();oParagraph.AddText('Hello world!');Api.GetDocument().InsertContent([oParagraph]);"
     *}]
     * window.Asc.plugin.executeMethod("InsertAndReplaceContentControls", [arrDocuments]);
     *
     * // Change existed content control
     * var arrDocuments = [{
     *  "Props": {
     *       "InternalId": "2_803"
     *   },
     *   "Script": "var oParagraph = Api.CreateParagraph();oParagraph.AddText('New text');Api.GetDocument().InsertContent([oParagraph]);"
     *}]
     * window.Asc.plugin.executeMethod("InsertAndReplaceContentControls", [arrDocuments]);

     */
    window["asc_docs_api"].prototype["pluginMethod_InsertAndReplaceContentControls"] = function(arrDocuments)
    {
        var _worker = new AscCommon.CContentControlPluginWorker(this, arrDocuments);
        return _worker.start();
    };
    /**
     * Removes several content controls.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias RemoveContentControls
     * @param {ContentControl[]} arrDocuments - An array of content control internal IDs. Example: [{"InternalId": "5_556"}].
     * @example
     * window.Asc.plugin.executeMethod("RemoveContentControls", [[{"InternalId": "5_556"}]])
     */
    window["asc_docs_api"].prototype["pluginMethod_RemoveContentControls"] = function(arrDocuments)
    {
        var _worker = new AscCommon.CContentControlPluginWorker(this, arrDocuments);
        return _worker.delete();
    };
    /**
     * Returns information about all the content controls that have been added to the page.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias GetAllContentControls
     * @returns {ContentControl[]} - An array of content control objects.
     * @example
     * window.Asc.plugin.executeMethod("GetAllContentControls");
     */
    window["asc_docs_api"].prototype["pluginMethod_GetAllContentControls"] = function()
    {
        var _blocks = this.WordControl.m_oLogicDocument.GetAllContentControls();
        var _ret = [];
        var _obj = null;
        for (var i = 0; i < _blocks.length; i++)
        {
            _obj = _blocks[i].GetContentControlPr();
            _ret.push({"Tag" : _obj.Tag, "Id" : _obj.Id, "Lock" : _obj.Lock, "InternalId" : _obj.InternalId});
        }
        return _ret;
    };

	/**
     * @typedef {Object} ContentControlParentPr
     * The content control parent properties.
     * @property  {object} Parent - The content control parent. For example, oParagraph.
     * @property  {number} Pos - The content control position within the parent object.
     * @property  {number} Count - A number of elements in the parent object.
     */

    /**
     * Removes the currently selected content control retaining all its contents. The content control where the mouse cursor is currently positioned will be removed.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias RemoveContentControl
     * @param {string} InternalId - A unique internal identifier of the content control.
     * @returns {ContentControlParentPr} - An object which contains the following values: Parent - content control parent, Pos - content control position within the parent object, Count - a number of elements in the parent object.
     * @example
     * window.Asc.plugin.executeMethod("RemoveContentControl", ["InternalId"])
     */
    window["asc_docs_api"].prototype["pluginMethod_RemoveContentControl"] = function(InternalId)
    {
        return this.asc_RemoveContentControlWrapper(InternalId);
    };
    /**
     * Returns an identifier of the selected content control (i.e. the content control where the mouse cursor is currently positioned).
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias GetCurrentContentControl
     * @returns {string} - The content control internal ID.
     * @example
     * window.Asc.plugin.executeMethod("GetCurrentContentControl");
     */
    window["asc_docs_api"].prototype["pluginMethod_GetCurrentContentControl"] = function()
    {
        return this.asc_GetCurrentContentControl();
    };
    /**
     * Returns the current content control properties.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias GetCurrentContentControlPr
	 * @param {string} contentFormat - The content format ("none", "text", "html", "ole" or "desktop").
     * @returns {ContentControlProperties} - The content control properties.
     * @example
     * window.Asc.plugin.executeMethod("GetCurrentContentControlPr")
     */
	window["asc_docs_api"].prototype["pluginMethod_GetCurrentContentControlPr"] = function(contentFormat)
	{
		var oLogicDocument = this.private_GetLogicDocument();

		var oState;
		var prop = this.asc_GetContentControlProperties();
		if (!prop)
			return null;

		if (oLogicDocument && prop.CC && contentFormat)
		{
			oState = oLogicDocument.SaveDocumentState();
			prop.CC.SelectContentControl();
		}

		var result =
		{
			"Tag"        : prop.Tag,
			"Id"         : prop.Id,
			"Lock"       : prop.Lock,
			"Alias"      : prop.Alias,
			"InternalId" : prop.InternalId,
			"Appearance" : prop.Appearance,
		};
		
		if (prop.Color)
		{
			result["Color"] =
			{
				"R" : prop.Color.r,
				"G" : prop.Color.g,
				"B" : prop.Color.b
			}
		}

		if (contentFormat)
		{
			var copy_data = {
				data     : "",
				pushData : function(format, value)
				{
					this.data = value;
				}
			};
			var copy_format = 1;
			if (contentFormat == Asc.EPluginDataType.html)
				copy_format = 2;
			this.asc_CheckCopy(copy_data, copy_format);
			result["content"] = copy_data.data;
		}

		if (oState && contentFormat)
		{
			oLogicDocument.LoadDocumentState(oState);
			oLogicDocument.UpdateSelection();
		}

		return result;
	};
    /**
     * Selects the specified content control.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias SelectContentControl
     * @param {string} id - A unique internal identifier of the content control.
     * @example
     * window.Asc.plugin.executeMethod("SelectContentControl", ["5_665"]);
     */
    window["asc_docs_api"].prototype["pluginMethod_SelectContentControl"] = function(id)
    {
        var oLogicDocument = this.private_GetLogicDocument();
        if (!oLogicDocument)
            return;

        oLogicDocument.SelectContentControl(id);
    };
    /**
     * Moves a cursor to the specified content control.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias MoveCursorToContentControl
     * @param {string} id - A unique internal identifier of the content control.
     * @param {boolean} [isBegin = false] - Defines if the cursor position changes in the content control. By default, a cursor will be placed to the content control begin (**false**).
     * @example
     * window.Asc.plugin.executeMethod("MoveCursorToContentControl", ["2_839", false])
     */
    window["asc_docs_api"].prototype["pluginMethod_MoveCursorToContentControl"] = function(id, isBegin)
    {
        var oLogicDocument = this.private_GetLogicDocument();
        if (!oLogicDocument)
            return;

        oLogicDocument.MoveCursorToContentControl(id, isBegin);
    };
    /**
     * Removes the selected content from the document.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias RemoveSelectedContent
     * @example
     *  window.Asc.plugin.executeMethod("RemoveSelectedContent")
     */
    window["asc_docs_api"].prototype["pluginMethod_RemoveSelectedContent"] = function()
    {
        var oLogicDocument = this.private_GetLogicDocument();
        if (!oLogicDocument || !oLogicDocument.IsSelectionUse())
            return;

        if (false === oLogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_Remove, null, true, oLogicDocument.IsFormFieldEditing()))
        {
            oLogicDocument.StartAction(AscDFH.historydescription_Document_BackSpaceButton);
            oLogicDocument.Remove(-1, true);
            oLogicDocument.FinalizeAction();
        }
    };

	/**
	 * @typedef {Object} comment
	 * Comment object.
	 * @property {string} Id - The comment ID.
	 * @property {CommentData} Data - An object which contains the comment data.
	 */

	/**
	 * @typedef {Object} CommentData
	 * The comment data.
	 * @property {string} UserName - The comment author.
	 * @property {string} QuoteText - The quote comment text.
	 * @property {string} Text - The comment text.
	 * @property {string} Time - The time when the comment was posted (in milliseconds).
	 * @property {boolean} Solved - Specifies if the comment is resolved (**true**) or not (**false**).
	 * @property {CommentData[]} Replies - An array containing the comment replies represented as the *CommentData* object.
	 */

	/**
	 * Adds a comment to the document.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias AddComment
	 * @param {CommentData}  oCommentData - An object which contains the comment data.
	 * @return {string | null} - The comment ID in the string format or null if the comment cannot be added.
	 */
	window["asc_docs_api"].prototype["pluginMethod_AddComment"] = function(oCommentData)
	{
		var oCD = undefined;
		if (oCommentData)
		{
			oCD = new AscCommon.CCommentData();
			oCD.ReadFromSimpleObject(oCommentData);
		}

		return this.asc_addComment(new window['Asc']['asc_CCommentDataWord'](oCD));
	};
    /**
     * Moves a cursor to the beginning of the current editing area (document body, footer/header, footnote, or autoshape).
	 * This method is similar to pressing the <b>Ctrl + Home</b> keyboard shortcut.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias MoveCursorToStart
     * @param {boolean} isMoveToMainContent - This flag ignores the current position and always moves a cursor to the beginning of the document body.
     */
    window["asc_docs_api"].prototype["pluginMethod_MoveCursorToStart"] = function(isMoveToMainContent)
    {
        var oLogicDocument = this.private_GetLogicDocument();
        if (oLogicDocument)
        {
            if (isMoveToMainContent)
                oLogicDocument.MoveCursorToStartOfDocument();
            else
                oLogicDocument.MoveCursorToStartPos(false);
        }
    };
    /**
     * Moves a cursor to the end of the current editing area (document body, footer/header, footnote, or autoshape).
	 * This method is similar to pressing the <b>Ctrl + End</b> keyboard shortcut.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias MoveCursorToEnd
     * @param {boolean} isMoveToMainContent - This flag ignores the current position and always moves a cursor to the end of the document body.
     */
    window["asc_docs_api"].prototype["pluginMethod_MoveCursorToEnd"] = function(isMoveToMainContent)
    {
        var oLogicDocument = this.private_GetLogicDocument();
        if (oLogicDocument)
        {
            if (isMoveToMainContent)
                oLogicDocument.MoveCursorToStartOfDocument();

            oLogicDocument.MoveCursorToEndPos(false);
        }
    };
    /**
     * Finds and replaces the text.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias SearchAndReplace
     * @param {Object} oProperties - An object which contains the search and replacement strings.
     * @param {string} oProperties.searchString - The search string.
     * @param {string} oProperties.replaceString - The replacement string.
     * @param {boolean} [oProperties.matchCase=true] - Case sensitive or not.
     */
    window["asc_docs_api"].prototype["pluginMethod_SearchAndReplace"] = function(oProperties)
    {
        var sReplace    = oProperties["replaceString"];

        let oProps = new AscCommon.CSearchSettings();
        oProps.SetText(oProperties["searchString"]);
        oProps.SetMatchCase(undefined !== oProperties["matchCase"] ? oProperties.matchCase : true);

        var oSearchEngine = this.WordControl.m_oLogicDocument.Search(oProps);
        if (!oSearchEngine)
            return;

        this.WordControl.m_oLogicDocument.ReplaceSearchElement(sReplace, true, null, false);
    };
	/**
	 * Finds and selects the next occurrence of the text starting at the current position.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias SearchNext
	 * @param {Object} oProperties - An object which contains the search string.
	 * @param {string} oProperties.searchString - The search string.
	 * @param {boolean} [oProperties.matchCase=true] - Case sensitive or not.
	 * @param {boolean} [isForward=true] - Search direction.
	 * @returns {boolean} returns false if text was not found
	 */
	window["asc_docs_api"].prototype["pluginMethod_SearchNext"] = function(oProperties, isForward)
	{
		let logicDocument = this.WordControl.m_oLogicDocument;
		if (!logicDocument)
			return false;
		
		let searchProps = new AscCommon.CSearchSettings();
		searchProps.SetText(oProperties["searchString"]);
		searchProps.SetMatchCase(undefined !== oProperties["matchCase"] ? oProperties["matchCase"] : true);
		
		logicDocument.Search(searchProps);
		let elementId = logicDocument.GetSearchElementId(!(false === isForward || 0 === isForward));
		if (null === elementId)
			return false;
		
		logicDocument.SelectSearchElement(elementId);
		return true;
	};
    /**
     * Returns file content in the HTML format.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias GetFileHTML
     * @return {string} - The HTML file content in the string format.
     * @example
     * window.Asc.plugin.executeMethod("GetFileHTML")
     */
    window["asc_docs_api"].prototype["pluginMethod_GetFileHTML"] = function()
    {
        return this.ContentToHTML(true);
    };
	/**
	 * Returns all the comments from the document.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias GetAllComments
	 * @returns {comment[]} - An array of comment objects containing the comment data.
	 */
	window["asc_docs_api"].prototype["pluginMethod_GetAllComments"] = function()
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var arrResult = [];

		var oComments = oLogicDocument.Comments.GetAllComments();
		for (var sId in oComments)
		{
			var oComment = oComments[sId];
			arrResult.push({"Id" : oComment.GetId(), "Data" : oComment.GetData().ConvertToSimpleObject()});
		}

		return arrResult;
	};
	/**
	 * Removes the specified comments.
	 * @param {string[]} arrIds - An array which contains the IDs of the specified comments.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias RemoveComments
	 */
	window["asc_docs_api"].prototype["pluginMethod_RemoveComments"] = function(arrIds)
	{
		this.asc_RemoveAllComments(false, false, arrIds);
	};
	/**
	 * Changes the specified comment.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias ChangeComment
	 * @param {string} sId - The comment ID.
	 * @param {CommentData} oCommentData - An object which contains the new comment data.
	 */
	window["asc_docs_api"].prototype["pluginMethod_ChangeComment"] = function(sId, oCommentData)
	{
		var oCD = undefined;
		if (oCommentData)
		{
			oCD = new AscCommon.CCommentData();
			oCD.ReadFromSimpleObject(oCommentData);

			var oLogicDocument = this.private_GetLogicDocument();
			if (oLogicDocument && AscCommonWord && AscCommonWord.CDocument && oLogicDocument instanceof AscCommonWord.CDocument)
			{
				var oComment = oLogicDocument.Comments.Get_ById(sId);
				if (oComment)
				{
					var sQuotedText = oComment.GetData().GetQuoteText();
					if (sQuotedText)
						oCD.SetQuoteText(sQuotedText);
				}
			}
		}

		this.asc_changeComment(sId, new window['Asc']['asc_CCommentDataWord'](oCD));
	};
	/**
	 * Moves a cursor to the specified comment.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias MoveToComment
	 * @param {string} sId - The comment ID.
	 */
	window["asc_docs_api"].prototype["pluginMethod_MoveToComment"] = function(sId)
	{
		this.asc_selectComment(sId);
		this.asc_showComment(sId);
	};
	/**
	 * Sets the display mode for track changes.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias SetDisplayModeInReview
	 * @param {string} [sMode="edit"] - The display mode:
	 * * <b>edit</b> - all changes are displayed,
	 * * <b>simple</b> - all changes are displayed but the balloons are turned off,
	 * * <b>final</b> - all accepted changes are displayed,
	 * * <b>original</b> - all rejected changes are displayed.
	 */
	window["asc_docs_api"].prototype["pluginMethod_SetDisplayModeInReview"] = function(sMode)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		if ("final" === sMode)
			oLogicDocument.SetDisplayModeInReview(Asc.c_oAscDisplayModeInReview.Final, true);
		else if ("original" === sMode)
			oLogicDocument.SetDisplayModeInReview(Asc.c_oAscDisplayModeInReview.Original, true);
		else if ("simple" === sMode)
			oLogicDocument.SetDisplayModeInReview(Asc.c_oAscDisplayModeInReview.Simple, true);
		else
			oLogicDocument.SetDisplayModeInReview(Asc.c_oAscDisplayModeInReview.Edit, true);
	};
	/**
	 * Adds an empty content control to the document.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias AddContentControl
	 * @param {ContentControlType} type - A numeric value that specifies the content control type. It can have one of the following values: <b>1</b> (block), <b>2</b> (inline), <b>3</b> (row), or <b>4</b> (cell).
	 * @param {ContentControlProperties}  [commonPr = {}] - The common content control properties.
	 * @returns {ContentControl} - A JSON object containing the data about the created content control.
	 * @example
	 * var type = 1;
	 * var properties = {"Id": 100, "Tag": "CC_Tag", "Lock": 3};
	 * window.Asc.plugin.executeMethod("AddContentControl", [type, properties]);
	 */
	window["asc_docs_api"].prototype["pluginMethod_AddContentControl"] = function(type, commonPr)
	{
		var _content_control_pr = private_ReadContentControlCommonPr(commonPr);

		var _obj = this.asc_AddContentControl(type, _content_control_pr);
		if (!_obj)
			return undefined;
		return {"Tag" : _obj.Tag, "Id" : _obj.Id, "Lock" : _obj.Lock, "InternalId" : _obj.InternalId};
	};

	/**
	 * @typedef {Object} ContentControlCheckBoxProperties
	 * The content control checkbox properties.
	 * @property {boolean} Checked - Defines if the content control checkbox is checked or not.
	 * @property {number} CheckedSymbol - A symbol in the HTML code format that is used when the checkbox is checked.
	 * @property {number} UncheckedSymbol - A symbol in the HTML code format that is used when the checkbox is not checked.
	 */

	/**
	 * Adds an empty content control checkbox to the document.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias AddContentControlCheckBox
	 * @param {ContentControlCheckBoxProperties}  [checkBoxPr = {}] - The content control checkbox properties.
	 * @param {ContentControlProperties}  [commonPr = {}] - The common content control properties.
	 * @example
	 * var checkBoxPr = {"Checked": false, "CheckedSymbol": 9746, "UncheckedSymbol": 9744};
	 * var commonPr = {"Id": 100, "Tag": "CC_Tag", "Lock": 3};
	 * window.Asc.plugin.executeMethod("AddContentControlCheckBox", [checkBoxPr, commonPr]);
	 */
	window["asc_docs_api"].prototype["pluginMethod_AddContentControlCheckBox"] = function(checkBoxPr, commonPr)
	{
		var oPr;
		if (checkBoxPr)
		{
			oPr = new AscWord.CSdtCheckBoxPr()
			if (checkBoxPr["Checked"])
				oPr.SetChecked(checkBoxPr["Checked"]);
			if (checkBoxPr["CheckedSymbol"])
				oPr.SetCheckedSymbol(checkBoxPr["CheckedSymbol"]);
			if (checkBoxPr["UncheckedSymbol"])
				oPr.SetUncheckedSymbol(checkBoxPr["UncheckedSymbol"]);
		}

		var _content_control_pr = private_ReadContentControlCommonPr(commonPr);

		this.asc_AddContentControlCheckBox(oPr, null, _content_control_pr);
	};

	/**
	 * Adds an empty content control picture to the document.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias AddContentControlPicture
	 * @param {ContentControlProperties}  [commonPr = {}] - The common content control properties.
	 * @example
	 * var commonPr = {"Id": 100, "Tag": "CC_Tag", "Lock": 3};
	 * window.Asc.plugin.executeMethod("AddContentControlPicture", [commonPr]);
	 */
	window["asc_docs_api"].prototype["pluginMethod_AddContentControlPicture"] = function(commonPr)
	{
		var _content_control_pr = private_ReadContentControlCommonPr(commonPr);

		this.asc_AddContentControlPicture(null, _content_control_pr);
	};
	/**
	 * Adds an empty content control list to the document.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias AddContentControlList
	 * @param {ContentControlType} type - A numeric value that specifies the content control type. It can have one of the following values: <b>1</b> (combo box), <b>0</b> (dropdown list).
	 * @param {Array<String, String>}  [List = [{Display, Value}]] - A list of the content control elements that consists of two items: <b>Display</b> - an item that will be displayed to the user in the content control list, <b>Value</b> - a value of each item from the content control list.
	 * @param {ContentControlProperties}  [commonPr = {}] - The common content control properties.
	 * @example
	 * var type = 1; //1 - ComboBox  0 - DropDownList
	 * var List = [{Display: "Item1_D", Value: "Item1_V"}, {Display: "Item2_D", Value: "Item2_V"}];
	 * var commonPr = {"Id": 100, "Tag": "CC_Tag", "Lock": 3};
	 * window.Asc.plugin.executeMethod("AddContentControlList", [type, List, commonPr]);
	 */
	window["asc_docs_api"].prototype["pluginMethod_AddContentControlList"] = function(type, List, commonPr)
	{
		var oPr;
		if (List)
		{
			oPr = new AscWord.CSdtComboBoxPr();
			List.forEach(function(el) {
				oPr.AddItem(el["Display"], el["Value"]);
			});
		}

		var _content_control_pr = private_ReadContentControlCommonPr(commonPr);

		this.asc_AddContentControlList(type, oPr, null, _content_control_pr);
	};

	/**
	 * @typedef {Object} ContentControlDatePickerProperties
	 * The content control datepicker properties.
	 * @property {string} DateFormat - A format in which the date will be displayed.
	 * For example: *"MM/DD/YYYY", "dddd\,\ mmmm\ dd\,\ yyyy", "DD\ MMMM\ YYYY", "MMMM\ DD\,\ YYYY", "DD-MMM-YY", "MMMM\ YY", "MMM-YY", "MM/DD/YYYY\ hh:mm\ AM/PM", "MM/DD/YYYY\ hh:mm:ss\ AM/PM", "hh:mm", "hh:mm:ss", "hh:mm\ AM/PM", "hh:mm:ss:\ AM/PM"*.
	 * @property {object} Date - The current date and time.
	 */

	/**
	 * Adds an empty content control datepicker to the document.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias AddContentControlDatePicker
	 * @param {ContentControlDatePickerProperties}  [datePickerPr = {}] - The content control datepicker properties.
	 * @param {ContentControlProperties}  [commonPr = {}] - The common content control properties.
	 * @example
	 * var DateFormats = [
	 * "MM/DD/YYYY",
	 * "dddd\,\ mmmm\ dd\,\ yyyy",
	 * "DD\ MMMM\ YYYY",
	 * "MMMM\ DD\,\ YYYY",
	 * "DD-MMM-YY",
	 * "MMMM\ YY",
	 * "MMM-YY",
	 * "MM/DD/YYYY\ hh:mm\ AM/PM",
	 * "MM/DD/YYYY\ hh:mm:ss\ AM/PM",
	 * "hh:mm",
	 * "hh:mm:ss",
	 * "hh:mm\ AM/PM",
	 * "hh:mm:ss:\ AM/PM"
	 * ];
	 * var Date = new window.Date();
	 * var datePickerPr = {"DateFormat" : DateFormats[2], "Date" : Date};
	 * var commonPr = {"Id": 100, "Tag": "CC_Tag", "Lock": 3};
	 * window.Asc.plugin.executeMethod("AddContentControlDatePicker", [datePickerPr, commonPr]);
	 */
	window["asc_docs_api"].prototype["pluginMethod_AddContentControlDatePicker"] = function(datePickerPr, commonPr)
	{
		var oPr;
		if (datePickerPr)
		{
			oPr = new AscWord.CSdtDatePickerPr();
			if (datePickerPr["Date"])
				oPr.SetFullDate(datePickerPr["Date"]);
			if (datePickerPr["DateFormat"])
				oPr.SetDateFormat(datePickerPr["DateFormat"]);
		}

		var _content_control_pr = private_ReadContentControlCommonPr(commonPr);

		this.asc_AddContentControlDatePicker(oPr, _content_control_pr);
	};


	/**
	 * @typedef {Object} OLEObjectData
	 * The OLE object data.
	 * @property {string} Data - OLE object data (internal format).
	 * @property {string} ImageData - An image in the base64 format stored in the OLE object and used by the plugin.
	 * @property {string} ApplicationId - An identifier of the plugin which can edit the current OLE object and must be of the *asc.{UUID}* type.
	 * @property {string} InternalId - The OLE object identifier which is used to work with OLE object added to the document.
	 * @property {string} ParaDrawingId - An identifier of the drawing object containing the current OLE object.
	 * @property {number} Width - The OLE object width measured in millimeters.
	 * @property {number} Height - The OLE object height measured in millimeters.
	 * @property {?number} WidthPix - The OLE object image width in pixels.
	 * @property {?number} HeightPix - The OLE object image height in pixels.
	 */
	
	/**
	 * @typedef {Object} AddinFieldData
	 * The addin field data.
	 * @property {string} FieldId - Field identifier.
	 * @property {string} Value - Field value.
	 * @property {string} Content - Field text content.
	 */

	/**
	 * Returns all OLE object data for objects which can be opened by the specified plugin.
	 * If *sPluginId* is not defined, this method returns all OLE objects contained in the currrent document.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias GetAllOleObjects
	 * @param {?string} sPluginId - Plugin identifier. It must be of the *asc.{UUID}* type.
	 * @returns {OLEObjectData[]} - An array of the OLEObjectData objects containing the data about the OLE object parameters.
	 * @since 7.1.0
	 * */
	window["asc_docs_api"].prototype["pluginMethod_GetAllOleObjects"] = function (sPluginId)
	{
		let aDataObjects = [];
		let oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument)
			return aDataObjects;
		let aOleObjects = oLogicDocument.GetAllOleObjects(sPluginId, []);
		for(let nObj = 0; nObj < aOleObjects.length; ++nObj)
		{
			aDataObjects.push(aOleObjects[nObj].getDataObject());
		}
		return aDataObjects;
	};

	/**
	 * Removes the OLE object from the document by its internal ID.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias RemoveOleObject
	 * @param {string} sInternalId - The OLE object identifier which is used to work with OLE object added to the document.
	 * @since 7.1.0
	 * */
	window["asc_docs_api"].prototype["pluginMethod_RemoveOleObject"] = function (sInternalId)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument)
		{
			return;
		}
		oLogicDocument.RemoveDrawingObjectById(sInternalId);
	};

	/**
	 * Removes several OLE objects from the document by their internal IDs.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias RemoveOleObjects
	 * @param {OLEObjectData[]} arrObjects An array of the identifiers which are used to work with OLE objects added to the document. Example: [{"InternalId": "5_556"}].
	 * @since 7.1.0
	 * @example
	 * window.Asc.plugin.executeMethod("RemoveOleObjects", [[{"InternalId": "5_556"}]])
	 */
	window["asc_docs_api"].prototype["pluginMethod_RemoveOleObjects"] = function (arrObjects)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if(!oLogicDocument)
		{
			return;
		}
		var arrIds = [];
		for(var nIdx = 0; nIdx < arrObjects.length; ++nIdx)
		{
			let oOleObject = arrObjects[nIdx];
			arrIds.push(oOleObject["InternalId"]);
		}
		oLogicDocument.RemoveDrawingObjects(arrIds);
	};

	/**
	 * Selects the specified OLE object.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias SelectOleObject
	 * @param {string} id - The OLE object identifier which is used to work with OLE object added to the document.
	 * @since 7.1.0
	 * @example
	 * window.Asc.plugin.executeMethod("SelectOleObject", ["5_665"]);
	 */
	window["asc_docs_api"].prototype["pluginMethod_SelectOleObject"] = function(id)
	{
		var oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;

		var oDrawing = AscCommon.g_oTableId.Get_ById(id);
		if(!oDrawing)
		{
			return;
		}
		oDrawing.Set_CurrentElement(true, null);
	};

	/**
	 * Inserts the OLE object at the current document position.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias InsertOleObject
	 * @param {OLEObjectData} NewObject - The OLEObjectData object.
	 * @param {?boolean} bSelect - Defines if the OLE object will be selected after inserting into the document (**true**) or not (**false**).
	 * @since 7.1.0
	 */
	window["asc_docs_api"].prototype["pluginMethod_InsertOleObject"] = function(NewObject, bSelect)
	{
		var oPluginData = {};
		oPluginData["imgSrc"] = NewObject["ImageData"];
		oPluginData["widthPix"] = NewObject["WidthPix"];
		oPluginData["heightPix"] = NewObject["HeightPix"];
		oPluginData["width"] = NewObject["Width"];
		oPluginData["height"] = NewObject["Height"];
		oPluginData["data"] = NewObject["Data"];
		oPluginData["guid"] = NewObject["ApplicationId"];
		oPluginData["select"] = bSelect;
		oPluginData["plugin"] = true;
		this.asc_addOleObject(oPluginData);
	};


	/**
	 * Changes the OLE object with the *InternalId* specified in OLE object data.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias ChangeOleObject
	 * @param {OLEObjectData} ObjectData - The OLEObjectData object.
	 * @since 7.1.0
	 */
	window["asc_docs_api"].prototype["pluginMethod_ChangeOleObject"] = function(ObjectData)
	{
		this["pluginMethod_ChangeOleObjects"]([ObjectData]);
	};
	/**
	 * Changes multiple OLE objects with the *InternalIds* specified in OLE object data.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias ChangeOleObjects
	 * @param {OLEObjectData[]} arrObjectData - An array of OLE object data.
	 * @since 7.1.0
	 */
	window["asc_docs_api"].prototype["pluginMethod_ChangeOleObjects"] = function(arrObjectData)
	{
		let oLogicDocument = this.private_GetLogicDocument();
		if (!oLogicDocument)
			return;
		let oParaDrawing;
		let oParaDrawingsMap = {};
		let nDrawing;
		let oDrawing;
		let oMainGroup;
		let aDrawings = [];
		let aParaDrawings = [];
		let oDataMap = {};
		let oData;
		for (nDrawing = 0; nDrawing < arrObjectData.length; ++nDrawing)
		{
			oData = arrObjectData[nDrawing];
			oDrawing = AscCommon.g_oTableId.Get_ById(oData["InternalId"]);
			oDataMap[oData["InternalId"]] = oData;
			if (oDrawing
				&& oDrawing.getObjectType
				&& oDrawing.getObjectType() === AscDFH.historyitem_type_OleObject
				&& oDrawing.IsUseInDocument())
			{
				aDrawings.push(oDrawing);
			}
		}
		for (nDrawing = 0; nDrawing < aDrawings.length; ++nDrawing)
		{
			oDrawing = aDrawings[nDrawing];
			if (oDrawing.group)
			{
				oMainGroup = oDrawing.getMainGroup();
				if (oMainGroup && oMainGroup.parent)
					oParaDrawingsMap[oMainGroup.parent.Id] = oMainGroup.parent;
			}
			else if (oDrawing.parent)
			{
				oParaDrawingsMap[oDrawing.parent.Id] = oDrawing.parent;
			}
		}
		for(let sId in oParaDrawingsMap)
		{
			if(oParaDrawingsMap.hasOwnProperty(sId))
			{
				oParaDrawing = oParaDrawingsMap[sId];
				aParaDrawings.push(oParaDrawing);
			}
		}
		if(aParaDrawings.length > 0)
		{
			let oStartState = oLogicDocument.SaveDocumentState();
			oLogicDocument.Start_SilentMode();
			oLogicDocument.SelectDrawings(aParaDrawings, oLogicDocument);
			if (!oLogicDocument.IsSelectionLocked(AscCommon.changestype_Drawing_Props))
			{
				oLogicDocument.StartAction()
				let oImagesMap = {};
				for(nDrawing = 0; nDrawing < aDrawings.length; ++nDrawing)
				{
					oDrawing = aDrawings[nDrawing];
					oData = oDataMap[oDrawing.Id];
					oDrawing.editExternal(oData["Data"], oData["ImageData"], oData["Width"], oData["Height"], oData["WidthPix"], oData["HeightPix"]);
					oImagesMap[oData["ImageData"]] = oData["ImageData"];
				}

				window.g_asc_plugins && window.g_asc_plugins.setPluginMethodReturnAsync();
				AscCommon.Check_LoadingDataBeforePrepaste(this, {}, oImagesMap, function() {
					oLogicDocument.Reassign_ImageUrls(oImagesMap);
					oLogicDocument.Recalculate();
					oLogicDocument.End_SilentMode();
					oLogicDocument.LoadDocumentState(oStartState);
					oLogicDocument.UpdateSelection();
					oLogicDocument.FinalizeAction();

					window.g_asc_plugins && window.g_asc_plugins.onPluginMethodReturn();
				});
			}
			else
			{
				oLogicDocument.End_SilentMode();
				oLogicDocument.LoadDocumentState(oStartState);
				oLogicDocument.UpdateSelection();
			}

		}
	};
	/**
	 * Accepts review changes.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias AcceptReviewChanges
	 * @param {boolean} [isAll=false] Specifies if all changes will be accepted (**true**) or only changes from the current selection (**false**).
	 * @since 7.2.1
	 * @example
	 * window.Asc.plugin.executeMethod("AcceptReviewChanges");
	 */
	window["asc_docs_api"].prototype["pluginMethod_AcceptReviewChanges"] = function(isAll)
	{
		if (isAll)
			this.asc_AcceptAllChanges();
		else
			this.asc_AcceptChangesBySelection(false);
	};
	/**
	 * Rejects review changes.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias RejectReviewChanges
	 * @param {boolean} [isAll=false] Specifies if all changes will be rejected (**true**) or only changes from the current selection (**false**).
	 * @since 7.2.1
	 * @example
	 * window.Asc.plugin.executeMethod("RejectReviewChanges");
	 */
	window["asc_docs_api"].prototype["pluginMethod_RejectReviewChanges"] = function(isAll)
	{
		if (isAll)
			this.asc_RejectAllChanges();
		else
			this.asc_RejectChangesBySelection(false);
	};
	/**
	 * Navigates through the review changes.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias MoveToNextReviewChange
	 * @param {boolean} [isForward=true] Specifies whether to navigate to the next (**true**) or previous (**false**) review change.
	 * @since 7.2.1
	 * @example
	 * window.Asc.plugin.executeMethod("MoveToNextReviewChange");
	 */
	window["asc_docs_api"].prototype["pluginMethod_MoveToNextReviewChange"] = function(isForward)
	{
		if (undefined !== isForward && !isForward)
			this.asc_GetPrevRevisionsChange();
		else
			this.asc_GetNextRevisionsChange();
	};
	/**
	 * Returns all addin fields from the current document.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias GetAllAddinFields
	 * @returns {AddinFieldData[]} - An array of the AddinFieldData objects containing the data about the addin fields.
	 * @since 7.3.3
	 * @example
	 * window.Asc.plugin.executeMethod("GetAllAddinFields");
	 */
	window["asc_docs_api"].prototype["pluginMethod_GetAllAddinFields"] = function()
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return [];
		
		let result = [];
		let fields = logicDocument.GetAllAddinFields();
		fields.forEach(function(field)
		{
			let fieldData = AscWord.CAddinFieldData.FromField(field);
			if (fieldData)
				result.push(fieldData.ToJson());
		});
		
		return result;
	};
	/**
	 * Updates the addin fields with the specified data.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias UpdateAddinFields
	 * @param {AddinFieldData[]} arrData - An array of addin field data.
	 * @since 7.3.3
	 * @example
	 * window.Asc.plugin.executeMethod("UpdateAddinFields");
	 */
	window["asc_docs_api"].prototype["pluginMethod_UpdateAddinFields"] = function(arrData)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument || !Array.isArray(arrData))
			return;
		
		let arrAddinData = [];
		arrData.forEach(function(data)
		{
			arrAddinData.push(AscWord.CAddinFieldData.FromJson(data));
		})
		
		logicDocument.UpdateAddinFieldsByData(arrAddinData);
	};
	/**
	 * Creates a new addin field with the data specified in the request.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias AddAddinField
	 * @param {AddinFieldData} data - Addin field data.
	 * @since 7.3.3
	 * @example
	 * window.Asc.plugin.executeMethod("AddAddinField");
	 */
	window["asc_docs_api"].prototype["pluginMethod_AddAddinField"] = function(data)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		logicDocument.AddAddinField(AscWord.CAddinFieldData.FromJson(data));
	};
	/**
	 * Removes a field wrapper, leaving only the field content.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias RemoveFieldWrapper
	 * @param {string} [fieldId=undefined] - Field ID. If it is not specified, then the wrapper of the current field is removed.
	 * @since 7.3.3
	 * @example
	 * window.Asc.plugin.executeMethod("RemoveFieldWrapper");
	 */
	window["asc_docs_api"].prototype["pluginMethod_RemoveFieldWrapper"] = function(fieldId)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		logicDocument.RemoveComplexFieldWrapper(fieldId);
	};
	/**
	 * Sets the document editing restrictions.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias SetEditingRestrictions
	 * @param {DocumentEditingRestrictions} restrictions - The document editing restrictions.
	 * @since 7.3.3
	 * @example
	 * window.Asc.plugin.executeMethod("SetEditingRestrictions");
	 */
	window["asc_docs_api"].prototype["pluginMethod_SetEditingRestrictions"] = function(restrictions)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return;
		
		let _restrictions = null;
		switch (restrictions)
		{
			case "comments": _restrictions = Asc.c_oAscRestrictionType.OnlyComments; break;
			case "forms": _restrictions = Asc.c_oAscRestrictionType.OnlyForms; break;
			case "readOnly": _restrictions = Asc.c_oAscRestrictionType.View; break;
			case "none": _restrictions = Asc.c_oAscRestrictionType.None; break;
		}
		
		if (null === _restrictions)
			return;
		
		this.asc_setRestriction(_restrictions);
	};
	/**
	 * Returns the current word.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias GetCurrentWord
	 * @param {TextPartType} [type="entirely"] - Specifies if the whole word or only its part will be returned.
	 * @returns {string} - A word or its part.
	 * @since 7.4.0
	 * @example
	 * window.Asc.plugin.executeMethod("GetCurrentWord");
	 */
	window["asc_docs_api"].prototype["pluginMethod_GetCurrentWord"] = function(type)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return "";
		
		return logicDocument.GetCurrentWord(private_GetTextDirection(type));
	};
	/**
	 * Replaces the current word with the specified string.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias ReplaceCurrentWord
	 * @param {string} replaceString - Replacement string.
	 * @param {TextPartType} [type="entirely"] - Specifies if the whole word or only its part will be replaced.
	 * @since 7.4.0
	 * @example
	 * window.Asc.plugin.executeMethod("ReplaceCurrentWord");
	 */
	window["asc_docs_api"].prototype["pluginMethod_ReplaceCurrentWord"] = function(replaceString, type)
	{
		let _replaceString = "" === replaceString ? "" : AscBuilder.GetStringParameter(replaceString, null);

		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument || null === _replaceString)
			return;
		
		logicDocument.ReplaceCurrentWord(private_GetTextDirection(type), _replaceString);
	};
	/**
	 * Returns the current sentence.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias GetCurrentSentence
	 * @param {TextPartType} [type="entirely"] - Specifies if the whole sentence or only its part will be returned.
	 * @returns {string} - A sentence or its part.
	 * @since 7.4.0
	 * @example
	 * window.Asc.plugin.executeMethod("GetCurrentSentence");
	 */
	window["asc_docs_api"].prototype["pluginMethod_GetCurrentSentence"] = function(type)
	{
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument)
			return "";
		
		return logicDocument.GetCurrentSentence(private_GetTextDirection(type));
	};
	/**
	 * Replaces the current sentence with the specified string.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias ReplaceCurrentSentence
	 * @param {string} replaceString - Replacement string.
	 * @param {TextPartType} [type="entirely"] - Specifies if the whole sentence or only its part will be replaced.
	 * @since 7.4.0
	 * @example
	 * window.Asc.plugin.executeMethod("ReplaceCurrentSentence");
	 */
	window["asc_docs_api"].prototype["pluginMethod_ReplaceCurrentSentence"] = function(replaceString, type)
	{
		let _replaceString = "" === replaceString ? "" : AscBuilder.GetStringParameter(replaceString, null);
		
		let logicDocument = this.private_GetLogicDocument();
		if (!logicDocument || null === _replaceString)
			return;

		
		return logicDocument.ReplaceCurrentSentence(private_GetTextDirection(type), _replaceString);
	};

	function private_ReadContentControlCommonPr(commonPr)
	{
		var resultPr;
		if (commonPr)
		{
			resultPr = new AscCommon.CContentControlPr();

			resultPr.Id    = commonPr["Id"];
			resultPr.Tag   = commonPr["Tag"];
			resultPr.Lock  = commonPr["Lock"];
			resultPr.Alias = commonPr["Alias"];

			if (undefined !== commonPr["Appearance"])
				resultPr.Appearance = commonPr["Appearance"];

			if (undefined !== commonPr["Color"])
				resultPr.Color = new Asc.asc_CColor(commonPr["Color"]["R"], commonPr["Color"]["G"], commonPr["Color"]["B"]);

			if (undefined !== commonPr["PlaceHolderText"])
				resultPr.SetPlaceholderText(commonPr["PlaceHolderText"]);
		}

		return resultPr;
	}
	function private_GetTextDirection(type)
	{
		let direction = 0;
		switch (AscBuilder.GetStringParameter(type, "entirely"))
		{
			case "beforeCursor":
				direction = -1;
				break;
			case "afterCursor":
				direction = 1;
				break;
			case "entirely":
			default:
				direction = 0;
				break;
		}
		return direction;
	}
	
})(window);
