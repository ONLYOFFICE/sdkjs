/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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
     * @typedef {Object} ContentControlLock
     * A value that defines if it is possible to delete and/or edit the content control or not.
     *
     * **0** - only deleting
     * **1** - disable deleting or editing
     * **2** - only editing
     * **3** - full access
     * @property {(0 | 1 | 2 | 3)} Lock
     */

    /**
     * @typedef {Object} ContentControlType
     * A numeric value that specifies the content control type.

     * @property  {(1 | 2 | 3 | 4)} type **1** - block content control, **2** - inline content control, **3** - row content control, **4** - cell content control
     */

    /**
     * @typedef {Object} ContentControlPropertiesAndContent
     * The content control properties and contents.
     * @property  {ContentControlProperties} [ContentControlProperties = {}] - The content control properties (ID, Tag, Lock, Alias, Appearance, Color).
     * @property  {string} Script - A script that will be executed to generate the data within the content control.
     * @property  {string} Url - A link to the shared file.
     */

    /**
     * @typedef {Object} ContentControlProperties
	 * The content control properties.
     * @property {string} Id - A unique identifier of the content control. It can be used to search for a certain content control and make reference to it in the code.
     * @property {string} Tag - A tag assigned to the content control. The same tag can be assigned to several content controls so that it is possible to make reference to them in your code.
     * @property {ContentControlLock} Lock - A value that defines if it is possible to delete and/or edit the content control or not: 0 - only deleting, 1 - no deleting or editing, 2 - only editing, 3 - full access.
     * @property {string} Alias - The alias attribute.
     * @property {string} Appearance - Defines if the content control is shown as the bounding box (1) or not (2).
     * @property {object} Color - The color for the current content control in the RGB format.
     * @property {number} Color.R - Red color component value.
     * @property {number} Color.G - Green color component value.
     * @property {number} Color.B - Blue color component value.
     * @example
     * {"Id": 100, "Tag": "CC_Tag", "Lock": 3}
     */

    var Api = window["asc_docs_api"];

    /**
     * Opens a file with fields.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias OpenFile
     * @param {Uint8Array} binaryFile - A file in the format of the 8-bit unsigned integer array.
     * @param {string[]} fields - A list of field values.
	 * @return {undefined}
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
     * Inserts the content control containing data. The data is specified by the js code for Document Builder, or by a link to the shared document.
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
     * @return {undefined}
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
     * Removes the currently selected content control retaining all its contents. The content control where the mouse cursor is currently positioned will be removed.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias RemoveContentControl
     * @param {string} InternalId - A unique internal identifier of the content control.
     * @returns {Object} - An object which contains the following values: Parent - content control parent, Pos - content control position within the parent object, Count - a number of elements in the parent object.
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
     * @returns {string} - InternalId of the selected content control.
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
     * @returns {ContentControlProperties} - Content control properties.
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

		if (prop && prop.CC) delete prop.CC;

		prop["Tag"] = prop.Tag;
		prop["Id"] = prop.Id;
		prop["Lock"] = prop.Lock;
		prop["InternalId"] = prop.InternalId;
		prop["Appearance"] = prop.Appearance;

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
			prop["content"] = copy_data.data;
		}

		if (oState && contentFormat)
		{
			oLogicDocument.LoadDocumentState(oState);
			oLogicDocument.UpdateSelection();
		}

		return prop;
	};
    /**
     * Selects the specified content control.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias SelectContentControl
     * @param {string} id - A unique internal identifier of the content control.
	 * @return {undefined}
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
     * @param {boolean} [isBegin = false] - Defines if the cursor position changes in the content control. By default, the cursor will be placed to the content control begin (false).
     * @return {undefined}
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
     * @return {undefined}
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
	 * Adds a comment to the document.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias AddComment
	 * @param {object}  oCommentData - An object which contains the comment data
	 * @param {string}  oCommentData.UserName - the comment author
	 * @param {string}  oCommentData.Text - the comment text
	 * @param {string}  oCommentData.Time - the comment time
	 * @param {boolean}  oCommentData.Solved - is the comment resolved
	 * @param {undefined | array} oCommentData.Replies - an array of replies, they are in the same format as oCommentData
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
     * Moves a cursor to the start position.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias MoveCursorToStart
     * @param {boolean} isMoveToMainContent - Defines if a cursor moves to the document start (true) or to the current element start (false).
	 * @return {undefined}
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
     * Moves a cursor to the end position.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias MoveCursorToEnd
     * @param {boolean} isMoveToMainContent - Defines if a cursor moves to the document end (true) or to the current element end (false).
	 * @return {undefined}
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
	 * @return {undefined}
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
	 * @returns {object[]} - An array which contains all the comments from the document.
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
	 * @return {undefined}
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
	 * @param {object} oCommentData - An object which contains the new comment data: "comment" - the comment text, "author" - the comment author.
	 * @return {undefined}
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
	 * @return {undefined}
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
	 * @param {"final" | "original" | "edit" | "simple"} [sMode="edit"] - The display mode: "edit" - all changes are displayed, "simple" - all changes are displayed but the balloons are turned off, "final" - all accepted changes are displayed, "original" - all rejected changes are displayed.
	 * @return {undefined}
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
	 * @param {ContentControlType} type - A numeric value that specifies the content control type. It can have one of the following values: 1 (block), 2 (inline), 3 (row) or 4 (cell).
	 * @param {ContentControlProperties}  [commonPr = {}] - The common content control properties.
	 * @returns {ContentControl} - A JSON object containing the data about the created content control: "Tag", "Id", "Lock" and "InternalId".
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
	 * Adds an empty content control checkbox to the document.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias AddContentControlCheckBox
	 * @param {ContentControlCkeckBoxProperties}  [checkBoxPr = {}] - The content control checkbox properties: "Checked" - defines if the content control checkbox is checked or not, "CheckedSymbol" - a symbol in the HTML code format that is used when the checkbox is checked, "UncheckedSymbol" - a symbol in the HTML code format that is used when the checkbox is not checked.
	 * @param {ContentControlProperties}  [commonPr = {}] - The common content control properties.
	 * @return {undefined}
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
	 * @return {undefined}
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
	 * @param {ContentControlType} type - A numeric value that specifies the content control type. It can have one of the following values: 1 (comboBox) or 0 (drop-down list).
	 * @param {Array<String, String>}  [List = [{Display, Value}]] - A list of the content control elements that consists of two items: "Display" - an item that will be displayed to the user in the content control list, "Value" - a value of each item from the content control list.
	 * @param {ContentControlProperties}  [commonPr = {}] - The common content control properties.
	 * @return {undefined}
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
				oPr.AddItem(el.Display, el.Value);
			});
		}

		var _content_control_pr = private_ReadContentControlCommonPr(commonPr);

		this.asc_AddContentControlList(type, oPr, null, _content_control_pr);
	};
	/**
	 * Adds an empty content control datepicker to the document.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias AddContentControlDatePicker
	 * @param {ContentControlDatePickerProperties}  [datePickerPr = {}] - The content control datepicker properties: "DateFormat" - a format in which the date will be displayed, "Date" - the current date and time.
	 * @param {ContentControlProperties}  [commonPr = {}] - The common content control properties.
	 * @return {undefined}
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
			if (datePickerPr.Date)
				oPr.SetFullDate(datePickerPr.Date);
			if (datePickerPr.DateFormat)
				oPr.SetDateFormat(datePickerPr.DateFormat);
		}

		var _content_control_pr = private_ReadContentControlCommonPr(commonPr);

		this.asc_AddContentControlDatePicker(oPr, _content_control_pr);
	};


	/**
	 * @typedef {Object} OLEObjectData
	 * The OLE object data.
	 * @property {string} Data - OLE object data (internal format).
	 * @property {string} ImageData - An image in the base64 format stored in the OLE object and used by the plugin.
	 * @property {string} ApplicationId - An identifier of the plugin which can edit the current OLE object and must be of the asc.{UUID} type.
	 * @property {string} InternalId - The OLE object identifier which is used to work with OLE object added to the document.
	 * @property {string} ParaDrawingId - An identifier of the drawing object containing the current OLE object.
	 * @property {number} Width - The OLE object width measured in millimeters.
	 * @property {number} Height - The OLE object height measured in millimeters.
	 * @property {?number} WidthPix - The OLE object image width in pixels.
	 * @property {?number} HeightPix - The OLE object image height in pixels.
	 */

	/**
	 * Returns all OLE object data for objects which can be opened by the specified plugin.
	 * If sPluginId is not defined, this method returns all OLE objects contained in the currrent document.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias GetAllOleObjects
	 * @param {?string} sPluginId - Plugin identifier. It must be of the asc.{UUID} type.
	 * @returns {OLEObjectData[]} - OLE object data.
	 * */
	window["asc_docs_api"].prototype["pluginMethod_GetAllOleObjects"] = function (sPluginId)
	{
		let aDataObjects = [];
		let aOleObjects = this.WordControl.m_oLogicDocument.GetAllOleObjects(sPluginId, []);
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
	 * @return {undefined}
	 * */
	window["asc_docs_api"].prototype["pluginMethod_RemoveOleObject"] = function (sInternalId)
	{
		this.WordControl.m_oLogicDocument.RemoveDrawingObjectById(sInternalId);
	};

	/**
	 * Removes several OLE objects from the document by their internal IDs.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias RemoveOleObjects
	 * @param {OLEObjectData[]} arrObjects An array of the identifiers which are used to work with OLE objects added to the document. Example: [{"InternalId": "5_556"}].
	 * @return {undefined}
	 * @example
	 * window.Asc.plugin.executeMethod("RemoveOleObjects", [[{"InternalId": "5_556"}]])
	 */
	window["asc_docs_api"].prototype["pluginMethod_RemoveOleObjects"] = function (arrObjects)
	{
		var arrIds = [];
		for(var nIdx = 0; nIdx < arrObjects.length; ++nIdx)
		{
			arrIds.push(arrObjects[nIdx].InternalId);
		}
		this.WordControl.m_oLogicDocument.RemoveDrawingObjects(arrIds);
	};

	/**
	 * Selects the specified OLE object.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias SelectOleObject
	 * @param {string} id - The OLE object identifier which is used to work with OLE object added to the document.
	 * @return {undefined}
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
	 * Adds the OLE object to the current document position.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias InsertOleObject
	 * @param {OLEObjectData} NewObject - The new OLE object data.
	 * @param {?boolean} bSelect - Defines if the OLE object will be selected after inserting into the document (true) or not (false).
	 * @return {undefined}
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
		this.asc_addOleObject(oPluginData);
	};


	/**
	 * Changes the OLE object with the InternalId specified in OLE object data.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias ChangeOleObject
	 * @param {OLEObjectData} ObjectData - OLE object data.
	 * @return {undefined}
	 */
	window["asc_docs_api"].prototype["pluginMethod_ChangeOleObject"] = function(ObjectData)
	{
		this["pluginMethod_ChangeOleObjects"]([ObjectData]);
	};
	/**
	 * Changes multiple OLE objects with the InternalIds specified in OLE object data.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias ChangeOleObjects
	 * @param {OLEObjectData[]} arrObjectData - An array of OLE object data.
	 * @return {undefined}
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
		for(nDrawing = 0; nDrawing < arrObjectData.length; ++nDrawing)
		{
			oData = arrObjectData[nDrawing];
			oDrawing = AscCommon.g_oTableId.Get_ById(oData.InternalId);
			oDataMap[oData.InternalId] = oData;
			if(oDrawing
				&& oDrawing.getObjectType
				&& oDrawing.getObjectType() === AscDFH.historyitem_type_OleObject)
			{
				if(oDrawing.IsUseInDocument())
				{
					aDrawings.push(oDrawing);
				}
			}
		}
		for(nDrawing = 0; nDrawing < aDrawings.length; ++nDrawing)
		{
			oDrawing = aDrawings[nDrawing];
			if(oDrawing.group)
			{
				oMainGroup = oDrawing.getMainGroup();
				if(oMainGroup)
				{
					if(oMainGroup.parent)
					{
						oParaDrawingsMap[oMainGroup.parent.Id] = oMainGroup.parent;
					}
				}
			}
			else
			{
				if(oDrawing.parent)
				{
					oParaDrawingsMap[oDrawing.parent.Id] = oDrawing.parent;
				}
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
				let oApi = this;
				AscCommon.Check_LoadingDataBeforePrepaste(this, {}, oImagesMap, function() {
					oLogicDocument.Reassign_ImageUrls(oImagesMap);
					oLogicDocument.Recalculate();
					oLogicDocument.End_SilentMode();
					oLogicDocument.LoadDocumentState(oStartState);
					oLogicDocument.UpdateSelection();
					oLogicDocument.FinalizeAction();
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
	 * Accept review changes
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias AcceptReviewChanges
	 * @param {boolean} [isAll=false] Accept all changes or only in the current selection
	 * @example
	 * window.Asc.plugin.executeMethod("AcceptReviewChanges");
	 *
	 */
	window["asc_docs_api"].prototype["pluginMethod_AcceptReviewChanges"] = function(isAll)
	{
		if (isAll)
			this.asc_AcceptAllChanges();
		else
			this.asc_AcceptChangesBySelection(false);
	};
	/**
	 * Reject review changes in the current selection
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias RejectReviewChanges
	 * @param {boolean} [isAll=false] Accept all changes or only in the current selection
	 * @example
	 * window.Asc.plugin.executeMethod("RejectReviewChanges");
	 *
	 */
	window["asc_docs_api"].prototype["pluginMethod_RejectReviewChanges"] = function(isAll)
	{
		if (isAll)
			this.asc_RejectAllChanges();
		else
			this.asc_RejectChangesBySelection(false);
	};
	/**
	 * Move to next review change
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @alias MoveToNextReviewChange
	 * @param {boolean} [isForward=true] Move forward or backward
	 * @example
	 * window.Asc.plugin.executeMethod("MoveToNextReviewChange");
	 *
	 */
	window["asc_docs_api"].prototype["pluginMethod_MoveToNextReviewChange"] = function(isForward)
	{
		if (undefined !== isForward && !isForward)
			this.asc_GetPrevRevisionsChange();
		else
			this.asc_GetNextRevisionsChange();
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

})(window);
