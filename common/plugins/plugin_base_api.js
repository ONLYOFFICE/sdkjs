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

window.startPluginApi = function() {

	/***********************************************************************
	 * CONFIG
	 */

	/**
	 * Object for translations
	 * example: { "en" : "name", "ru" : "имя" }
	 * @typedef { Object.<string, string> } localeTranslate
	 */

	/**
	 * EditorType
	 * @typedef {("word" | "cell" | "slide")} editorType
	 * */

	/**
	 * Variation init data type
	 * @typedef {("none" | "text" | "ole" | "html")} initDataType
	 * */

	/**
	 * @typedef { Object } variationButton
	 * @property {string} text Label
	 * @property {boolean} [primary] Is button use primary css style
	 * @proverty {boolean} [isViewer] Is button shown in viewer mode
	 * @property {localeTranslate} [textLocale] Translations for text field
	 */

	/**
	 * @typedef { Object } variation
	 * @description Why would one plugin might need some variations? The answer is simple enough: the plugin can not only perform some actions but also contain some settings, or 'About' window, or something like that. For example, translation plugin: the plugin itself does not need a visual window for translation as it can be done just pressing a single button, but its settings (the translation direction) and 'About' window must be visual. So we will need to have at least two plugin variations (translation itself and settings), or three, in case we want to add an 'About' window with the information about the plugin and its authors or the software used for the plugin creation.
	 *
	 * @property {string} description Name
	 * @property {localeTranslate} [descriptionName] Translations for description field
	 *
	 * @property {string} url Plugin entry point, i.e. HTML file which connects the plugin.js (the base file needed for work with plugins) file and launches the plugin code.
	 *
	 * @proverty {string[]} icons Plugin icon image files used in the editors: for common screens and with doubled resolution for retina screens.
	 *
	 * @property {boolean} isViewer=false Specifies if the plugin is available when the document is available in viewer mode only or not.
	 * @property {editorType[]} EditorsSupport=Array.<string>("word","cell","slide") The editors which the plugin is available for ("word" - text document editor, "cell" - spreadsheet editor, "slide" - presentation editor).
	 *
	 * @property {boolean} isVisual Specifies if the plugin is visual (will open a window for some action, or introduce some additions to the editor panel interface) or non-visual (will provide a button (or buttons) which is going to apply some transformations or manipulations to the document).
	 * @property {boolean} isModal Specifies if the opened plugin window is modal, i.e. a separate modal window must be opened, or not (used for visual plugins only). The following rule must be observed at all times: isModal != isInsideMode.
	 * @property {boolean} isInsideMode Specifies if the plugin must be displayed inside the editor panel instead of its own window (used for visual non-modal plugins only). The following rule must be observed at all times: isModal != isInsideMode.
	 * @property {boolean} [isSystem] Specifies if the plugin is not displayed in the editor interface and is started in background with the server (or desktop editors start) not interfering with the other plugins, so that they can work simultaneously.
	 * @property {boolean} [isDisplayedInViewer] Specifies if the plugin will be displayed in viewer mode as well as in editor mode (isDisplayedInViewer == true) or in the editor mode only (isDisplayedInViewer == false).
	 *
	 * @property {boolean} [initOnSelectionChanged] Specifies if the plugin watches the text selection events in the editor window.
 	 *
	 * @property {boolean} [isUpdateOleOnResize] Specifies if the OLE object must be redrawn when resized in the editor using the vector object draw type or not (used for OLE objects only, i.e. initDataType == "ole").
	 *
	 * @property {initDataType} initDataType The data type selected in the editor and sent to the plugin: "text" - the text data, "html" - HTML formatted code, "ole" - OLE object data, "none" - no data will be send to the plugin from the editor.
	 * @property {string} initData Is always equal to "" - this is the data which is sent from the editor to the plugin at the plugin start (e.g. if initDataType == "text", the plugin will receive the selected text when run).
	 *
	 * @property {number[]} [size] Size of the plugin window
	 *
	 * @property {variationButton[]} [buttons] Buttons
	 *
	 * @property {string[]} events Events
	 */


	/**
	 * @typedef {Object} Config
	 * @property {string} basePath="" Path to the plugin. All the other paths are calculated relative to this path. In case the plugin is installed on the server, an additional parameter (path to the plugins) is added there. If baseUrl == "" the path to all plugins will be used.
	 * @property {string} guid Plugin identifier. It <b>must<b> be of the asc.{UUID} type.
	 *
	 * @property {string} name Plugin name which will be visible at the plugin toolbar.
	 * @property {localeTranslate} [nameLocale] Translations for name field
	 *
	 * @property {variation[]} variations Plugin variations or "subplugins"
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
	 * @param {string} text
	 */

	/**
	 * Event: button
	 * @event Plugin#button
	 * @memberof Plugin
	 * @alias button
	 * @param {number} buttonIndex
	 */

	/**
	 * Event: onTargetPositionChanged
	 * @event Plugin#event_onTargetPositionChanged
	 * @memberof Plugin
	 * @alias event_onTargetPositionChanged
	 */

	/**
	 * Event: onDocumentContentReady
	 * @event Plugin#event_onDocumentContentReady
	 * @memberof Plugin
	 * @alias event_onDocumentContentReady
	 */

	/**
	 * Event: onClick
	 * @event Plugin#event_onClick
	 * @memberof Plugin
	 * @alias event_onClick
	 */

	/**
	 * Event: inputHelper_onSelectItem
	 * @event Plugin#event_inputHelper_onSelectItem
	 * @memberof Plugin
	 * @alias event_inputHelper_onSelectItem
	 */

	/**
	 * Event: onInputHelperClear
	 * @event Plugin#event_onInputHelperClear
	 * @memberof Plugin
	 * @alias event_onInputHelperClear
	 */

	/**
	 * Event: onInputHelperInput
	 * @event Plugin#event_onInputHelperInput
	 * @memberof Plugin
	 * @alias event_onInputHelperInput
	 */

	/**
	 * Event: onTranslate
	 * @event Plugin#onTranslate
	 * @memberof Plugin
	 * @alias onTranslate
	 */

	var Plugin = window["Asc"]["plugin"];

	/***********************************************************************
	 * METHODS
	 */

	/**
	 * executeCommand
	 * @memberof Plugin
	 * @alias executeCommand
	 * @deprecated Please use callCommand method
	 * @param {string} type "close" or "command"
     * @param {string} data Script code
     * @param {Function} callback
	 */
	Plugin.executeCommand = function(type, data, callback)
    {
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
	 * @param {string} name Name of the method
	 * @param {Array} params Parameters of the method
	 * @param {Function} callback Callback function
	 */
	Plugin.executeMethod = function(name, params, callback)
    {
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
	 * @param {number} width New width of the window
     * @param {number} height New height of the window
     * @param {number} minW New min-width of the window
     * @param {number} minH New min-height of the window
     * @param {number} maxW New max-width of the window
	 * @param {number} maxH New max-height of the window
	 */
	Plugin.resizeWindow = function(width, height, minW, minH, maxW, maxH)
    {
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
	 * @param {Function} func Function to call
	 * @param {boolean} isClose
     * @param {boolean} isCalc
	 * @param {Function} callback Callback function
	 */
	Plugin.callCommand = function(func, isClose, isCalc, callback)
    {
        var _txtFunc = "var Asc = {}; Asc.scope = " + JSON.stringify(window.Asc.scope) + "; var scope = Asc.scope; (" + func.toString() + ")();";
        var _type = (isClose === true) ? "close" : "command";
        window.Asc.plugin.info.recalculate = (false === isCalc) ? false : true;
        window.Asc.plugin.executeCommand(_type, _txtFunc, callback);
    };

	/**
	 * callModule
	 * @memberof Plugin
	 * @alias callModule
	 * @param {string} url Url to resource code
	 * @param {Function} callback Callback function
	 * @param {boolean} isClose
	 */
	Plugin.callModule = function(url, callback, isClose)
    {
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
	 * @param {string} url Url to resource code
	 * @param {Function} callback Callback function
	 */
	Plugin.loadModule = function(url, callback)
    {
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

	/***********************************************************************
	 * INPUT HELPERS
 	 */

	/**
	 * @typedef {Object} InputHelperItem
	 * @property {string} id
	 * @property {string} text
	 */

	/**
	 * @global
	 * @class
	 * @name InputHelper
	 * @hideconstructor
	 */

	/**
	 * @function createWindow
	 * @memberof InputHelper
	 * @alias createWindow
	 */

	/**
	 * @function getItems
	 * @memberof InputHelper
	 * @alias getItems
	 * @return {InputHelperItem[]}
	 */

	/**
	 * @function setItems
	 * @memberof InputHelper
	 * @alias setItems
	 * @param {InputHelperItem[]}
	 */

	/**
	 * @function show
	 * @memberof InputHelper
	 * @alias show
	 * @param {number} width
	 * @param {number} height
	 * @param {boolean} isCaptureKeyboard
	 */

	/**
	 * @function unShow
	 * @memberof InputHelper
	 * @alias unShow
	 */

	/**
	 * @function getScrollSizes
	 * @memberof InputHelper
	 * @alias getScrollSizes
	 * @return {number}
	 */

	/**
	 * createInputHelper
	 * @memberof Plugin
	 * @alias createInputHelper
	 */
	Plugin.createInputHelper = function()
    {
        window.Asc.plugin.ih = new window.Asc.inputHelper(window.Asc.plugin);
    };
	/**
	 * getInputHelper
	 * @memberof Plugin
	 * @alias getInputHelper
	 * @return {InputHelper} Input helper object
	 */
	Plugin.getInputHelper = function()
	{
		return window.Asc.plugin.ih;
	};

};
