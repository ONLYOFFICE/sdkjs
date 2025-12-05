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

/**
 * @typedef {Object} ContentControl
 * The content control object.
 * @property {string} Tag - A tag assigned to the content control. The same tag can be assigned to several content controls so that you can make reference to them in your code.
 * @property {string} Id - A unique content control identifier. It can be used to search for a certain content control and make reference to it in your code.
 * @property {ContentControlLock} Lock - A value that defines if it is possible to delete and/or edit the content control or not.
 * @property {string} InternalId - A unique internal identifier of the content control. It is used for all operations with content controls.
 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/ContentControl.js
 */

/**
 * @typedef {(0 | 1 | 2 | 3)} ContentControlLock
 * Defines the access restrictions for a content control.
 * Possible values:
 * <b>0</b> - only deleting,
 * <b>1</b> - disable deleting or editing,
 * <b>2</b> - only editing,
 * <b>3</b> - full access.
 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/ContentControlLock.js
 */

/**
 * @typedef {Object} comment
 * The comment object.
 * @property {string} Id - The comment ID.
 * @property {CommentData} Data - An object which contains the comment data.
 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/comment.js
 */

/**
 * @typedef {Object} CommentData
 * The comment data.
 * @property {string} UserName - The comment author.
 * @property {string} Text - The comment text.
 * @property {string} Time - The time when the comment was posted (in milliseconds).
 * @property {boolean} Solved - Specifies if the comment is resolved (**true**) or not (**false**).
 * @property {CommentData[]} Replies - An array containing the comment replies represented as the *CommentData* object.
 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/CommentData.js
 */

/**
 * COMMENTS
 */

/**
 * Event: onAddComment
 * @event Plugin#onAddComment
 * @memberof Plugin
 * @typeofeditors ["CDE"]
 * @alias onAddComment
 * @description The function called when a comment is added to the document with the {@link /docs/plugins/text-document-api/Methods/AddComment AddComment} method.
 * @param {comment} comment - Defines the comment object containing the comment data.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onAddComment.js
 */

/**
 * Event: onChangeCommentData
 * @event Plugin#onChangeCommentData
 * @memberof Plugin
 * @typeofeditors ["CDE"]
 * @alias onChangeCommentData
 * @description The function called when the specified comment is changed with the {@link /docs/plugins/text-document-api/Methods/ChangeComment ChangeComment} method.
 * @param {comment} comment - Defines the comment object containing the comment data.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onChangeCommentData.js
 */

/**
 * Event: onChangeCurrentPage
 * @event Plugin#onChangeCurrentPage
 * @memberof Plugin
 * @typeofeditors ["CDE"]
 * @alias onChangeCurrentPage
 * @description The function called when the current page has changed.
 * @param {number} index - The index of the newly activated page.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onChangeCurrentPage.js
 */

/**
 * Event: onRemoveComment
 * @event Plugin#onRemoveComment
 * @memberof Plugin
 * @typeofeditors ["CDE"]
 * @alias onRemoveComment
 * @description The function called when the specified comment is removed with the {@link /docs/plugins/text-document-api/Methods/RemoveComments RemoveComments} method.
 * @param {comment} comment - Defines the comment object containing the comment data.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onRemoveComment.js
 */

/**
 * FORMS
 */

/**
 * Event: onSubmitForm
 * @event Plugin#onSubmitForm
 * @memberof Plugin
 * @typeofeditors ["CDE"]
 * @alias onSubmitForm
 * @description The function called when the user clicks the "Complete & Submit" button.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onSubmitForm.js
 */

/**
 * CC
 */

/**
 * Event: onFocusContentControl
 * @event Plugin#onFocusContentControl
 * @memberof Plugin
 * @typeofeditors ["CDE"]
 * @alias onFocusContentControl
 * @description The function called to show which content control has been focused.
 * @param {ContentControl} control - Defines the content control that has been focused.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onFocusContentControl.js
 */

/**
 * Event: onBlurContentControl
 * @event Plugin#onBlurContentControl
 * @memberof Plugin
 * @typeofeditors ["CDE"]
 * @alias onBlurContentControl
 * @description The function called to show which content control has been blurred.
 * @param {ContentControl} control - Defines the content control that has been blurred.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onBlurContentControl.js
 */

/**
 * Event: onChangeContentControl
 * @event Plugin#onChangeContentControl
 * @memberof Plugin
 * @typeofeditors ["CDE"]
 * @alias onChangeContentControl
 * @description The function called to show which content control has been changed.
 * @param {ContentControl} control - Defines the content control that has been changed.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onChangeContentControl.js
 */

/**
 * Event: onHideContentControlTrack
 * @event Plugin#onHideContentControlTrack
 * @memberof Plugin
 * @typeofeditors ["CDE"]
 * @alias onHideContentControlTrack
 * @description The function called when the content control loses focus in the document.
 * @param {string[]} ids - An array of content control IDs that have lost focus.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onHideContentControlTrack.js
 */

/**
 * Event: onShowContentControlTrack
 * @event Plugin#onShowContentControlTrack
 * @memberof Plugin
 * @typeofeditors ["CDE"]
 * @alias onShowContentControlTrack
 * @description The function called when the content control receives focus and its track appears.
 * @param {string[]} ids - An array of content control IDs that have received focus.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onShowContentControlTrack.js
 */

/**
 * Event: onInsertOleObjects
 * @event Plugin#onInsertOleObjects
 * @memberof Plugin
 * @typeofeditors ["CDE"]
 * @alias onInsertOleObjects
 * @description The function called when one or more OLE objects are inserted into the document.
 * @param {object[]} data - An array containing information about the inserted OLE objects.
 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Events/onInsertOleObjects.js
 */

