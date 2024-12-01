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

window.startPluginApi = function() {

	/***********************************************************************
	 * CONFIG
	 */

	/**
	 * Translations for the text field. The object keys are the two letter language codes (ru, de, it, etc.) and the values are the button label translation for each language.
	 * Example: { "en" : "name", "ru" : "имя" }
	 * @typedef { Object.<string, string> } localeTranslate
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/localeTranslate.js
	 */

	/**
	 * The editors which the plugin is available for:
	 * <b>word</b> - text document editor,
	 * <b>cell</b> - spreadsheet editor,
	 * <b>slide</b> - presentation editor,
	 * <b>pdf</b> - pdf editor.
	 * @typedef {("word" | "cell" | "slide" | "pdf")} editorType
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/editorType.js
	 */

	/**
	 * The data type selected in the editor and sent to the plugin:
     * <b>text</b> - the text data,
	 * <b>html</b> - HTML formatted code,
	 * <b>ole</b> - OLE object data,
     * <b>desktop</b> - the desktop editor data,
     * <b>destop-external</b> - the main page data of the desktop app (system messages),
     * <b>none</b> - no data will be send to the plugin from the editor,
	 * <b>sign</b> - the sign for the keychain plugin.
	 * @typedef {("text" | "html" | "ole" | "desktop" | "destop-external" | "none" | "sign")} initDataType
     * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/initDataType.js
	 */

	/**
	 * The skinnable plugin button used in the plugin interface (used for visual plugins with their own window only, i.e. isVisual == true and isInsideMode == false).
	 * @typedef { Object } Button
	 * @property {string} text - The label which is displayed on the button.
	 * @property {boolean} [primary] - Defines if the button is primary or not. The primary flag affects the button skin only.
	 * @property {boolean} [isViewer] - Defines if the button is shown in the viewer mode only or not.
	 * @property {localeTranslate} [textLocale] - Translations for the text field. The object keys are the two letter language codes (ru, de, it, etc.) and the values are the button label translation for each language.
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/Button.js
	 */

	/**
	 * @typedef {Object} ContentControl
	 * Content control object.
	 * @property {string} Tag - A tag assigned to the content control. The same tag can be assigned to several content controls so that you can make reference to them in your code.
	 * @property {string} Id - A unique content control identifier. It can be used to search for a certain content control and make reference to it in your code.
	 * @property {ContentControlLock} Lock - A value that defines if it is possible to delete and/or edit the content control or not.
	 * @property {string} InternalId - A unique internal identifier of the content control. It is used for all operations with content controls.
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/ContentControl.js
	 */

	/**
     * @typedef {(0 | 1 | 2 | 3)} ContentControlLock
     * A value that defines if it is possible to delete and/or edit the content control or not:
	 * <b>0</b> - only deleting,
	 * <b>1</b> - disable deleting or editing,
	 * <b>2</b> - only editing,
	 * <b>3</b> - full access.
     * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/ContentControlLock.js
	 */

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
	 * @property {string} Text - The comment text.
	 * @property {string} Time - The time when the comment was posted (in milliseconds).
	 * @property {boolean} Solved - Specifies if the comment is resolved (**true**) or not (**false**).
	 * @property {CommentData[]} Replies - An array containing the comment replies represented as the *CommentData* object.
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/CommentData.js
	 */


	/**
	 * Plugin variations, or subplugins, that are created inside the origin plugin.
	 * @typed { Object } variation
	 * @descr Plugin variations can be created for the following purposes:
	 * to perform the main plugin actions,
	 * to contain plugin settings,
	 * to display an About window, etc.
	 * For example, the Translation plugin: the plugin itself does not need a visual window for translation as it can be done just pressing a single button, but its settings (the translation direction) and an 'About' window must be visual. So we will need to have at least two plugin variations (translation itself and settings), or three, in case we want to add an 'About' window with the information about the plugin and its authors or the software used for the plugin creation.
	 *
	 * @pr {string} description - The description, i.e. what describes your plugin the best way.
	 * @pr {localeTranslate} [descriptionLocale] - Translations for the description field. The object keys are the two letter language codes (ru, de, it, etc.) and the values are the plugin description translation for each language.
	 *
	 * @pr {string} url - Plugin entry point, i.e. an HTML file which connects the plugin.js file (the base file needed for work with plugins) and launches the plugin code.
	 *
	 * @pr {string[]} icons - Plugin icon image files used in the editors. There can be several scaling types for plugin icons: 100%, 125%, 150%, 175%, 200%, etc.
	 *
	 * @pr {editorType[]} [EditorsSupport=Array.<string>("word","cell","slide")] - The editors which the plugin is available for ("word" - text document editor, "cell" - spreadsheet editor, "slide" - presentation editor).
	 * @pr {boolean} [isViewer=false] - Specifies if the plugin is working when the document is available in the viewer mode only or not.
	 * @pr {boolean} [isDisplayedInViewer=true] - Specifies if the plugin will be displayed in the viewer mode as well as in the editor mode (isDisplayedInViewer == true) or in the editor mode only (isDisplayedInViewer == false).
	 *
	 * @pr {boolean} isVisual - Specifies if the plugin is visual (will open a window for some action, or introduce some additions to the editor panel interface) or non-visual (will provide a button (or buttons) which is going to apply some transformations or manipulations to the document).
	 * @pr {boolean} isModal - Specifies if the opened plugin window is modal (used for visual plugins only, and if isInsideMode is not true).
	 * @pr {boolean} isInsideMode - Specifies if the plugin must be displayed inside the editor panel instead of its own window.
	 * @pr {boolean} isSystem - Specifies if the plugin is not displayed in the editor interface and is started in the background with the server (or desktop editors start) not interfering with the other plugins, so that they can work simultaneously.
	 *
	 * @pr {boolean} initOnSelectionChanged - Specifies if the plugin watches the text selection events in the editor window.
 	 *
	 * @pr {boolean} isUpdateOleOnResize - Specifies if an OLE object must be redrawn when resized in the editor using the vector object draw type or not (used for OLE objects only, i.e. initDataType == "ole").
	 *
	 * @pr {initDataType} initDataType - The data type selected in the editor and sent to the plugin: "text" - the text data, "html" - HTML formatted code, "ole" - OLE object data, "desktop" - the desktop editor data, "destop-external" - the main page data of the desktop app (system messages), "none" - no data will be send to the plugin from the editor.
	 * @pr {string} initData - Is usually equal to "" - this is the data which is sent from the editor to the plugin at the plugin start (e.g. if initDataType == "text", the plugin will receive the selected text when run). It may also be equal to encryption in the encryption plugins.
	 *
	 * @pr {number[]} [size] - Plugin window size.
	 *
	 * @pr {Button[]} [buttons] - The list of skinnable plugin buttons used in the plugin interface (used for visual plugins with their own window only, i.e. isVisual == true && isInsideMode == false).
	 *
	 * @pr {string[]} events - Plugin events.
	 */


	/**
	 * @typed {Object} Config
	 * @pr {string} [basePath=""] - Path to the plugin. All the other paths are calculated relative to this path. In case the plugin is installed on the server, an additional parameter (path to the plugins) is added there. If baseUrl == "", the path to all plugins will be used.
	 * @pr {string} guid - Plugin identifier. It must be of the asc.{UUID} type.
	 * @pr {string} minVersion - The minimum supported editors version.
	 *
	 * @pr {string} name - Plugin name which will be visible at the plugin toolbar.
	 * @pr {localeTranslate} [nameLocale] - Translations for the name field. The object keys are the two letter language codes (ru, de, it, etc.) and the values are the plugin name translation for each language.
	 *
	 * @pr {variation[]} variations - Plugin variations, or subplugins, that are created inside the origin plugin.
	 */

	/***********************************************************************
	 * EVENTS
	 */

	/**
	 * @global
	 * @class
	 * @name Plugin
	 * @hideconstructor
	 */

	/***********************************************************************
	 * EVENTS
	 */

	/**
	 * Event: init
	 * @event Plugin#init
	 * @memberof Plugin
	 * @alias init
	 * @description The function called when the plugin is launched. It defines the data sent to the plugin describing what actions are to be performed and how they must be performed.
	 * @param {string} text - Defines the data parameter that depends on the {@link /plugin/config#initDataType initDataType} setting specified in the *config.json* file.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/init.js
	 */

	/**
	 * Event: button
	 * @event Plugin#button
	 * @memberof Plugin
	 * @alias button
	 * @description The function called when any of the plugin buttons is clicked. It defines the buttons used with the plugin and the plugin behavior when they are clicked.
	 * @param {number} buttonIndex - Defines the button index in the {@link /plugin/config#buttons buttons} array of the *config.json* file. If *id == -1*, then the plugin considers that the <b>Close</b> window cross button has been clicked or its operation has been somehow interrupted.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/button.js
	 */

	/**
	 * Event: onTargetPositionChanged
	 * @event Plugin#onTargetPositionChanged
	 * @memberof Plugin
	 * @alias onTargetPositionChanged
	 * @description The function called when the target position in the editor is changed.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onTargetPositionChanged.js
	 */

	/**
	 * Event: onDocumentContentReady
	 * @event Plugin#onDocumentContentReady
	 * @memberof Plugin
	 * @alias onDocumentContentReady
	 * @description The function called when the document is completely loaded.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onDocumentContentReady.js
	 */

	/**
	 * Event: onClick
	 * @event Plugin#onClick
	 * @memberof Plugin
	 * @alias onClick
	 * @description The function called when the user clicks on the element.
	 * @param {boolean} isSelectionUse - Defines if the selection is used or not.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onClick.js
	 */

	/**
	 * Event: inputHelper_onSelectItem
	 * @event Plugin#inputHelper_onSelectItem
	 * @memberof Plugin
	 * @alias inputHelper_onSelectItem
	 * @description The function called when the user is trying to select an item from the input helper.
	 * @param {object} item - Defines the selected item:
	 * <b>text</b> - the item text,  
	 * <b>type</b>: string,  
	 * <b>example</b>: "name";
	 * <b>id</b> - the item index,  
	 * <b>type</b>: string,  
	 * <b>example</b>: "1".
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/inputHelper_onSelectItem.js
	 */

	/**
	 * Event: onInputHelperClear
	 * @event Plugin#onInputHelperClear
	 * @memberof Plugin
	 * @alias onInputHelperClear
	 * @description The function called when the user is trying to clear the text and the input helper disappears.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onInputHelperClear.js
	 */

	/**
	 * Event: onInputHelperInput
	 * @event Plugin#onInputHelperInput
	 * @memberof Plugin
	 * @alias onInputHelperInput
	 * @description The function called when the user is trying to input the text and the input helper appears.
	 * @param {object} data - Defines the text which the user inputs:
	 * <b>add</b> - defines if the text is added to the current text (**true**) or this is the beginning of the text (**false**),  
	 * <b>type</b>: boolean,  
	 * <b>example</b>: true;
	 * <b>text</b> - the text which the user inputs,  
	 * <b>type</b>: string,  
	 * <b>example</b>: "text".
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onInputHelperInput.js
	 */

	/**
	 * Event: onTranslate
	 * @event Plugin#onTranslate
	 * @memberof Plugin
	 * @alias onTranslate
	 * @description The function called right after the plugin startup or later in case the plugin language is changed.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onTranslate.js
	 */

    /**
     * Event: onEnableMouseEvent
     * @event Plugin#onEnableMouseEvent
     * @memberof Plugin
     * @alias onEnableMouseEvent
	 * @description The function called to turn the mouse or touchpad events on/off.
	 * @param {boolean} isEnabled - Defines if the mouse or touchpad is enabled (**true**) or not (**false**).
     * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onEnableMouseEvent.js
	 */

    /**
     * Event: onExternalMouseUp
     * @event Plugin#onExternalMouseUp
     * @memberof Plugin
     * @alias onExternalMouseUp
	 * @description The function called when the mouse button is released outside the plugin iframe.
     * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onExternalMouseUp.js
	 */

    /**
     * Event: onExternalPluginMessage
     * @event Plugin#onExternalPluginMessage
     * @memberof Plugin
     * @alias onExternalPluginMessage
     * @description The function called to show the editor integrator message.
     * @param {Object} data - Defines the editor integrator message:
	 * <b>type</b> - the message type,  
	 * <b>type</b>: string,  
	 * <b>example</b>: "close";
	 * <b>text</b> - the message text,  
	 * <b>type</b>: string,  
	 * <b>example</b>: "text".
     * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onExternalPluginMessage.js
	 */

	/**
	 * Event: onFocusContentControl
	 * @event Plugin#onFocusContentControl
	 * @memberof Plugin
	 * @typeofeditors ["CDE"]
	 * @alias onFocusContentControl
	 * @description The function called to show which content control has been focused.
	 * @param {ContentControl} control - Defines the content control that has been focused.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onFocusContentControl.js
	 */

	/**
	 * Event: onBlurContentControl
	 * @event Plugin#onBlurContentControl
	 * @memberof Plugin
	 * @typeofeditors ["CDE"]
	 * @alias onBlurContentControl
	 * @description The function called to show which content control has been blurred.
	 * @param {ContentControl} control - Defines the content control that has been blurred.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onBlurContentControl.js
	 */

	/**
	 * Event: onChangeContentControl
	 * @event Plugin#onChangeContentControl
	 * @memberof Plugin
	 * @typeofeditors ["CDE"]
	 * @alias onChangeContentControl
	 * @description The function called to show which content control has been changed.
	 * @param {ContentControl} control - Defines the content control that has been changed.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onChangeContentControl.js
	 */

	/**
	 * The context menu type:
	 * <b>None</b> - not used,
	 * <b>Target</b> - nothing is selected,
	 * <b>Selection</b> - text is selected,
	 * <b>Image</b> - image is selected,
	 * <b>Shape</b> - shape is selected,
	 * <b>OleObject</b> - OLE object is selected.
	 * @typedef {("None" | "Target" | "Selection" | "Image" | "Shape" | "OleObject")} ContextMenuType
	 */

	/**
	 * @typedef {Object} ContextMenuOptions
	 * @description Defines the context menu options.
	 * @property {ContextMenuType} Type - The context menu type.
	 * @property {boolean} [header] - Specifies if the context menu is opened inside the header.
	 * @property {boolean} [footer] - Specifies if the context menu is opened inside the footer.
	 * @property {boolean} [headerArea] - Specifies if the context menu is opened over the header.
	 * @property {boolean} [footerArea] - Specifies if the context menu is opened over the footer.
	 */

	/**
	 * Event: onContextMenuShow
	 * WARNING! If plugin is listening this event, it MUST call AddContextMenuItem method (synchronously or not),
	 * because editor wait answers from ALL plugins and then and only then fill contextmenu.
	 * @event Plugin#onContextMenuShow
	 * @memberof Plugin
	 * @alias onContextMenuShow
	 * @description The function called when the context menu has been shown.
	 * 
	 * <note>If a plugin is listening for this event, it must call the {@link /plugin/executeMethod/common/addcontextmenuitem AddContextMenuItem} method (synchronously or not),
	 * because the editor waits for responses from all plugins before filling the context menu.</note>
	 * @param {ContextMenuOptions} options - Defines the context menu information.
	 * @since 7.4.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onContextMenuShow.js
	 */

	/**
	 * Event: onContextMenuClick
	 * @event Plugin#onContextMenuClick
	 * @memberof Plugin
	 * @alias onContextMenuClick
	 * @description The function called when the context menu item has been clicked.
	 * @param {string} id - Item ID.
	 * @since 7.4.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onContextMenuClick.js
	 */

	/**
	 * Event: onToolbarMenuClick
	 * @event Plugin#onToolbarMenuClick
	 * @memberof Plugin
	 * @alias onToolbarMenuClick
	 * @description The function called when the toolbar menu item has been clicked.
	 * @param {string} id - Item ID.
	 * @since 8.1.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onToolbarMenuClick.js
	 */

	/**
	 * Event: onCommandCallback
	 * @event Plugin#onCommandCallback
	 * @memberof Plugin
	 * @typeofeditors ["CDE"]
	 * @alias onCommandCallback
	 * @description The function called to return the result of the previously executed command. It can be used to return data after executing the {@link Plugin#callCommand callCommand} method.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onCommandCallback.js
	 */

	/**
	 * Event: onMethodReturn
	 * @event Plugin#onMethodReturn
	 * @memberof Plugin
	 * @typeofeditors ["CDE"]
	 * @alias onMethodReturn
	 * @description The function called to return the result of the previously executed method. It can be used to return data after executing the {@link Plugin#executeMethod executeMethod} method.
	 * @param returnValue - Defines the value that will be returned.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onMethodReturn.js
	 */

	/**
	 * Event: onAddComment
	 * @event Plugin#onAddComment
	 * @memberof Plugin
	 * @typeofeditors ["CDE"]
	 * @alias onAddComment
	 * @description The function called when a comment is added to the document with the {@link /plugin/executeMethod/text/addcomment AddComment} method.
	 * @param {comment} comment - Defines the comment object containing the comment data.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onAddComment.js
	 */

	/**
	 * Event: onChangeCommentData
	 * @event Plugin#onChangeCommentData
	 * @memberof Plugin
	 * @typeofeditors ["CDE"]
	 * @alias onChangeCommentData
	 * @description The function called when the specified comment is changed with the {@link /plugin/executeMethod/text/changecomment ChangeComment} method.
	 * @param {comment} comment - Defines the comment object containing the comment data.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onChangeCommentData.js
	 */

	/**
	 * Event: onRemoveComment
	 * @event Plugin#onRemoveComment
	 * @memberof Plugin
	 * @typeofeditors ["CDE"]
	 * @alias onRemoveComment
	 * @description The function called when the specified comment is removed with the {@link /plugin/executeMethod/text/removecomments RemoveComments} method.
	 * @param {comment} comment - Defines the comment object containing the comment data.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onRemoveComment.js
	 */

	/**
	 * Event: onSubmitForm
	 * @event Plugin#onSubmitForm
	 * @memberof Plugin
	 * @typeofeditors ["CDE"]
	 * @alias onSubmitForm
	 * @description The function called when the user clicks the "Complete & Submit" button.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/onSubmitForm.js
	 */

    var Plugin = window["Asc"]["plugin"];

	Plugin._checkPluginOnWindow = function(isWindowSupport)
	{
		if (this.windowID && !isWindowSupport)
		{
			console.log("This method does not allow in window frame");
			return true;
		}
		if (!this.windowID && true === isWindowSupport)
		{
			console.log("This method is allow only in window frame");
			return true;
		}
		return false;
	};

	Plugin._pushWindowMethodCommandCallback = function(callback)
	{
		if (this.windowCallbacks === undefined)
		{
			this.windowCallbacks = [];
			this.attachEvent("on_private_window_method", function(retValue) {
				var _retCallback = window.Asc.plugin.windowCallbacks.shift();
				if (_retCallback)
					_retCallback(retValue);
			});
			this.attachEvent("on_private_window_command", function(retValue) {
				var _retCallback = window.Asc.plugin.windowCallbacks.shift();
				if (_retCallback)
					_retCallback(retValue);
			});
		}
		this.windowCallbacks.push(callback);
	};

	/***********************************************************************
	 * METHODS
	 */

	/**
	 * executeCommand
	 * @undocumented
	 * @memberof Plugin
	 * @alias executeCommand
	 * @deprecated Please use callCommand method.
	 * @description Defines the method used to send the data back to the editor.
	 * <note>This method is deprecated, please use the {@link Plugin#callCommand callCommand} method which runs the code from the *data* string parameter.</note>
	 * 
	 * Now this method is mainly used to work with the OLE objects or close the plugin without any other commands.
	 * It is also retained for using with text so that the previous versions of the plugin remain compatible.
	 * 
	 * The *callback* is the result that the command returns. It is an optional parameter. In case it is missing, the window.Asc.plugin.onCommandCallback function will be used to return the result of the command execution.
	 * 
	 * The second parameter is the JavaScript code for working with <b>ONLYOFFICE Document Builder</b> API 
	 * that allows the plugin to send structured data inserted to the resulting document file (formatted paragraphs, tables, text parts, and separate words, etc.).
	 * <note><b>ONLYOFFICE Document Builder</b> commands can be only used to create content and insert it to the document editor
	 * (using the *Api.GetDocument().InsertContent(...)*). This limitation exists due to the co-editing feature in the online editors.
	 * If it is necessary to create a plugin for the desktop editors to work with local files, no such limitation is applied.</note>
	 * 
	 * When creating/editing OLE objects, two extensions are used to work with them:
	 * Api.asc_addOleObject (window.Asc.plugin.info)* - used to create an OLE object in the document;
	 * Api.asc_editOleObject (window.Asc.plugin.info)* - used to edit the created OLE object.
	 * 
	 * When creating/editing the objects, their properties can be passed to the window.Asc.plugin.info object that defines how the object should look.
	 * @param {string} type - Defines the type of the command. The *close* is used to close the plugin window after executing the function in the *data* parameter.
	 * The *command* is used to execute the command and leave the window open waiting for the next command.
     * @param {string} data - Defines the command written in JavaScript code which purpose is to form the structured data which can be inserted to the resulting document file
	 * (formatted paragraphs, tables, text parts, and separate words, etc.). Then the data is sent to the editors.
	 * The command must be compatible with {@link /docbuilder/basic ONLYOFFICE Document Builder} syntax.
     * @param {Function} callback - The result that the method returns.
	 */
	Plugin.executeCommand = function(type, data, callback)
    {
		if (this._checkPluginOnWindow() && 0 !== type.indexOf("onmouse")) return;

        window.Asc.plugin.info.type = type;
        window.Asc.plugin.info.data = data;

        var _message = "";
        try
        {
            _message = JSON.stringify(window.Asc.plugin.info);
        }
        catch(err)
        {
            _message = JSON.stringify({ type : data });
        }

        window.Asc.plugin.onCallCommandCallback = callback;
        window.plugin_sendMessage(_message);
    };

	/**
	 * executeMethod
	 * @memberof Plugin
	 * @alias executeMethod
	 * @description Defines the method used to execute certain editor methods using the plugin.
	 * 
	 * The callback is the result that the method returns. It is an optional parameter. In case it is missing, the {@link Plugin#onMethodReturn window.Asc.plugin.onMethodReturn} function will be used to return the result of the method execution.
	 * @param {string} name - The name of the specific method that must be executed.
	 * @param {Array} params - The arguments that the method in use has (if it has any).
     * @param {Function} callback - The result that the method returns.
	 * @returns {boolean}
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/executeMethod.js
	 */
	Plugin.executeMethod = function(name, params, callback)
    {
		if (this.windowID)
		{
			this._pushWindowMethodCommandCallback(callback);
			this.sendToPlugin("private_window_method", { name : name, params : params });
			return;
		}

        if (window.Asc.plugin.isWaitMethod === true)
        {
            if (undefined === this.executeMethodStack)
                this.executeMethodStack = [];

            this.executeMethodStack.push({ name : name, params : params, callback : callback });
            return false;
        }

        window.Asc.plugin.isWaitMethod = true;
        window.Asc.plugin.methodCallback = callback;

        window.Asc.plugin.info.type = "method";
        window.Asc.plugin.info.methodName = name;
        window.Asc.plugin.info.data = params;

        var _message = "";
        try
        {
            _message = JSON.stringify(window.Asc.plugin.info);
        }
        catch(err)
        {
            return false;
        }
        window.plugin_sendMessage(_message);
        return true;
    };

	/**
	 * resizeWindow (only for visual modal plugins)
	 * @memberof Plugin
	 * @alias resizeWindow
	 * @description Defines the method used to change the window size updating the minimum/maximum sizes.
	 * <note>This method is used for visual modal plugins only.</note>
	 * @param {number} width - The window width.
     * @param {number} height - The window height.
     * @param {number} minW - The window minimum width.
     * @param {number} minH - The window minimum height.
     * @param {number} maxW - The window maximum width.
	 * @param {number} maxH - The window maximum height.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/resizeWindow.js
	 */
	Plugin.resizeWindow = function(width, height, minW, minH, maxW, maxH)
    {
		// TODO: resize with window ID
		if (this._checkPluginOnWindow()) return;

        if (undefined === minW) minW = 0;
        if (undefined === minH) minH = 0;
        if (undefined === maxW) maxW = 0;
        if (undefined === maxH) maxH = 0;

        var data = JSON.stringify({ width : width, height : height, minw : minW, minh : minH, maxw : maxW, maxh : maxH });

        window.Asc.plugin.info.type = "resize";
        window.Asc.plugin.info.data = data;

        var _message = "";
        try
        {
            _message = JSON.stringify(window.Asc.plugin.info);
        }
        catch(err)
        {
            _message = JSON.stringify({ type : data });
        }
        window.plugin_sendMessage(_message);
    };

	/**
	 * callCommand
	 * @memberof Plugin
	 * @alias callCommand
	 * @description Defines the method used to send the data back to the editor.
	 * It allows the plugin to send structured data that can be inserted to the resulting document file (formatted paragraphs, tables, text parts, and separate words, etc.).
	 * 
	 * The *callback* is the result that the command returns. It is an optional parameter. In case it is missing, the {@link Plugin#onCommandCallback window.Asc.plugin.onCommandCallback} function will be used to return the result of the command execution.
     * <note><b>ONLYOFFICE Document Builder</b> commands can be only used to create content and insert it to the document editor (using the *Api.GetDocument().InsertContent(...)*).
	 * This limitation exists due to the co-editing feature in the online editors. If it is necessary to create a plugin for desktop editors to work with local files, no such limitation is applied.</note>
     * 
	 * This method is executed in its own context isolated from other JavaScript data. If some parameters or other data need to be passed to this method, use {@link /plugin/scope Asc.scope} object.
	 * @param {Function} func - Defines the command written in JavaScript which purpose is to form structured data which can be inserted to the resulting document file
	 * (formatted paragraphs, tables, text parts, and separate words, etc.). Then the data is sent to the editors.
	 * The command must be compatible with {@link /docbuilder/basic ONLYOFFICE Document Builder} syntax.
	 * @param {boolean} isClose - Defines whether the plugin window must be closed after the code is executed or left open waiting for another command or action.
	 * The *true* value is used to close the plugin window after executing the function in the *func* parameter.
	 * The *false* value is used to execute the command and leave the window open waiting for the next command.
     * @param {boolean} isCalc - Defines whether the document will be recalculated or not.
	 * The *true* value is used to recalculate the document after executing the function in the *func* parameter.
	 * The *false* value will not recalculate the document (use it only when your edits surely will not require document recalculation).
	 * @param {Function} callback - The result that the method returns. Only the js standart types are available (any objects will be replaced with *undefined*).
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/callCommand.js
	 */
	Plugin.callCommand = function(func, isClose, isCalc, callback)
    {
		var _txtFunc = "var Asc = {}; Asc.scope = " + JSON.stringify(window.Asc.scope) + "; var scope = Asc.scope; (" + func.toString() + ")();";
		if (this.windowID)
		{
			this._pushWindowMethodCommandCallback(callback);
			this.sendToPlugin("private_window_command", { code : _txtFunc, isCalc : isCalc });
			return;
		}

        var _type = (isClose === true) ? "close" : "command";
        window.Asc.plugin.info.recalculate = (false === isCalc) ? false : true;
        window.Asc.plugin.executeCommand(_type, _txtFunc, callback);
    };

	/**
	 * callModule
	 * @memberof Plugin
	 * @alias callModule
	 * @description Defines the method used to execute a remotely located script following a link.
	 * @param {string} url - The resource code URL.
	 * @param {Function} callback - The result that the method returns.
	 * @param {boolean} isClose - Defines whether the plugin window must be closed after the code is executed or left open waiting for another action.
	 * The *true* value is used to close the plugin window after executing a remotely located script.
	 * The *false* value is used to execute the code and leave the window open waiting for the next action.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/callModule.js
	 */
	Plugin.callModule = function(url, callback, isClose)
    {
		if (this._checkPluginOnWindow()) return;

        var _isClose = isClose;
        var _client = new XMLHttpRequest();
        _client.open("GET", url);

        _client.onreadystatechange = function() {
            if (_client.readyState == 4 && (_client.status == 200 || location.href.indexOf("file:") == 0))
            {
                var _type = (_isClose === true) ? "close" : "command";
                window.Asc.plugin.info.recalculate = true;
                window.Asc.plugin.executeCommand(_type, _client.responseText);
                if (callback)
                    callback(_client.responseText);
            }
        };
        _client.send();
    };

	/**
	 * loadModule
	 * @memberof Plugin
	 * @alias loadModule
	 * @description Defines the method used to load a remotely located text resource.
     * @param {string} url - The resource code URL.
	 * @param {Function} callback - The result that the method returns.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/loadModule.js
	 */
	Plugin.loadModule = function(url, callback)
    {
		if (this._checkPluginOnWindow()) return;

        var _client = new XMLHttpRequest();
        _client.open("GET", url);

        _client.onreadystatechange = function() {
            if (_client.readyState == 4 && (_client.status == 200 || location.href.indexOf("file:") == 0))
            {
                if (callback)
                    callback(_client.responseText);
            }
        };
        _client.send();
    };

	/**
	 * @function attachEvent
	 * @memberof Plugin
	 * @alias attachEvent
	 * @description Defines the method to add an event listener, a function that will be called whenever the specified event is delivered to the target.
	 * The list of all the available events can be found {@link /plugin/events here}.
     * @param {string} id - The event name.
	 * @param {Function} action - The event listener.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/attachEvent.js
	 */

	/**
	 * @function attachContextMenuClickEvent
	 * @memberof Plugin
	 * @alias attachContextMenuClickEvent
	 * @description Defines the method to add an event listener, a function that will be called whenever the specified event is clicked in the context menu.
     * @param {string} id - The event name.
	 * @param {Function} action - The event listener.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/attachContextMenuClickEvent.js
	 */

	/***********************************************************************
	 * INPUT HELPERS
 	 */

	/**
	 * @typedef {Object} InputHelperItem
	 * @description Defines the input helper item.
	 * @property {string} id - The item index.
	 * @property {string} text - The item text.
	 * @see office-js-api/Examples/Plugins/{Editor}/Enumeration/InputHelperItem.js
	 */

	/**
	 * Class representing an input helper - a window that appears and disappears when you type text. Its location is tied to the cursor.
	 * @global
	 * @class
	 * @name InputHelper
	 * @hideconstructor
	 */

	/**
	 * @function createWindow
	 * @memberof InputHelper
	 * @alias createWindow
	 * @description Creates an input helper window.
	 * @see office-js-api/Examples/Plugins/{Editor}/InputHelper/Methods/createWindow.js
	 */

	/**
	 * @function getItems
	 * @memberof InputHelper
	 * @alias getItems
	 * @description Returns an array of the {@link global#InputHelperItem InputHelperItem} objects that contain all the items from the input helper.
	 * @return {InputHelperItem[]}
	 * @see office-js-api/Examples/Plugins/{Editor}/InputHelper/Methods/getItems.js
	 */

	/**
	 * @function setItems
	 * @memberof InputHelper
	 * @alias setItems
	 * @description Sets the items to the input helper.
	 * @param {InputHelperItem[]} items - Defines an array of the {@link global#InputHelperItem InputHelperItem} objects which contain all the items for the input helper.
	 * @see office-js-api/Examples/Plugins/{Editor}/InputHelper/Methods/setItems.js
	 */

	/**
	 * @function show
	 * @memberof InputHelper
	 * @alias show
	 * @description Shows an input helper.
	 * @param {number} width - The input helper window width measured in millimeters.
	 * @param {number} height - The input helper window height measured in millimeters.
	 * @param {boolean} isCaptureKeyboard - Defines if the keyboard is caught (**true**) or not (**false**).
	 * @see office-js-api/Examples/Plugins/{Editor}/InputHelper/Methods/show.js
	 */

	/**
	 * @function unShow
	 * @memberof InputHelper
	 * @alias unShow
	 * @description Hides an input helper.
	 * @see office-js-api/Examples/Plugins/{Editor}/InputHelper/Methods/unShow.js
	 */

	/**
	 * @function getScrollSizes
	 * @memberof InputHelper
	 * @alias getScrollSizes
	 * @description Returns the sizes of the input helper scrolled window. Returns an object with width and height parameters.
	 * @return {number}
	 * @see office-js-api/Examples/Plugins/{Editor}/InputHelper/Methods/getScrollSizes.js
	 */

	/**
	 * createInputHelper
	 * @memberof Plugin
	 * @alias createInputHelper
	 * @description Defines the method used to create an {@link inputhelper input helper} - a window that appears and disappears when you type text. Its location is tied to the cursor.
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/createInputHelper.js
	 */
	Plugin.createInputHelper = function()
    {
		if (this._checkPluginOnWindow()) return;

        window.Asc.plugin.ih = new window.Asc.inputHelper(window.Asc.plugin);
    };
	/**
	 * getInputHelper
	 * @memberof Plugin
	 * @alias getInputHelper
	 * @description Defines the method used to get the {@link inputhelper InputHelper object}.
	 * @return {InputHelper} Input helper object
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/getInputHelper.js
	 */
	Plugin.getInputHelper = function()
	{
		if (this._checkPluginOnWindow()) return;

		return window.Asc.plugin.ih;
	};

	/**
	 * sendToPlugin
	 * @memberof Plugin
	 * @alias sendToPlugin
	 * @description Sends a message from the modal window to the plugin.
	 * @param {string} name - The event name.
	 * @param {object} data - The event data.
	 * @return {boolean} Returns true if the operation is successful.
	 * @since 7.4.0
	 * @see office-js-api/Examples/Plugins/{Editor}/Plugin/Methods/sendToPlugin.js
	 */
	Plugin.sendToPlugin = function(name, data)
	{
		if (this._checkPluginOnWindow(true)) return;

		window.Asc.plugin.info.type = "messageToPlugin";
		window.Asc.plugin.info.eventName = name;
		window.Asc.plugin.info.data = data;
		window.Asc.plugin.info.windowID = this.windowID;

		var _message = "";
		try
		{
			_message = JSON.stringify(window.Asc.plugin.info);
		}
		catch(err)
		{
			return false;
		}
		window.plugin_sendMessage(_message);
		return true;
	};

};


