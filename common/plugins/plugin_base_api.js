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
	 * @property {boolean} isModal Specifies if the opened plugin window is modal (used for visual plugins only, and if isInsideMode is not true).
	 * @property {boolean} isInsideMode Specifies if the plugin must be displayed inside the editor panel instead of its own window.
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
	 * @property {string} [minVersion] The minimum supported editors version.
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
	 * @description The method invokes pressing of a button in a window with a plugin. The button is an index in the list (see config.json).
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

    /**
     * Event: onEnableMouseEvent
     * @event Plugin#onEnableMouseEvent
     * @memberof Plugin
     * @alias onEnableMouseEvent
	 * @description The method turns off/on mouse/touchpad events.
	 * @param {boolean} isEnabled
     */

    /**
     * Event: onExternalMouseUp
     * @event Plugin#onExternalMouseUp
     * @memberof Plugin
     * @alias onExternalMouseUp
	 * @description The method indicates that the mouse/touchpad event was completed outside the plugin while started inside the plugin.
     */

    /**
     * Event: onExternalPluginMessage
     * @event Plugin#onExternalPluginMessage
     * @memberof Plugin
     * @alias onExternalPluginMessage
     * @description The method shows the editor integrator message (see externallistener plugin).
     * @param {Object} data
     */


    var Plugin = window["Asc"]["plugin"];

	/***********************************************************************
	 * METHODS
	 */

	/**
	 * executeCommand
	 * @memberof Plugin
	 * @alias executeCommand
	 * @deprecated Please use callCommand method.
	 * @description THE METHOD IS OBSOLETE. Use window.Asc.plugin.callCommand which runs the code from "data" string parameter.
     * After the code is executed, callback is called. If callback === undefined,
	 * then window.Asc.plugin.onCommandCallback is called in case it is implemented in the plugin.
	 * Script in the "data" variable is a builder script (see documentation LINK).
	 * Сurrently, it is allowed to insert content created with the help of builder methods
	 * only with InsertContent method, for the plugin not to handle images/fonts loading (as they are handled in the editor asynchronously).
	 * In the future this restriction will be eliminated.
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
	 * @description The function calls the editor method.
	 * After execution of the function, callback with a parameter is called - the value returned by the method. If callback === undefined,
	 * then window.Asc.plugin.onMethodReturn with a parameter is called - the value returned by the method in case it is implemented in the plugin.
	 * @param {string} name Name of the method
	 * @param {Array} params Array with calling parameters
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
	 * @description The funcion is intended for visual modal plugins -
	 * it changes the window size updating the minimum/maximum sizes
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
	 * @description The method performs the "func" function.
     * After that callback is called. If callback === undefined,
     * then window.Asc.plugin.onCommandCallback is called, in case it is executed in the plugin
     * if isClose === true, then after the plugin is executed, it will close.
     * if isCalc === false, then the editor will not recalculate the document (use this parameter carefully -
     * only when you are sure that your edits will not require document recalculation)
     * Code in the "func" function is a builder script (see documentation LINK)
     * Currently, it is allowed to insert content created with the help of builder methods
     * only with InsertContent method, for the plugin not to handle images/fonts loading (as they are run in the editor asynchronously)
     * In the future this restriction will be eliminated
     *
     * Note: in the "func" code it is not allowed to use external plugin variables.
     * If it is necessary, you need to export the variable in Asc.scope object, and use it the function code
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
	 * @description The function executes a remotely located script following a link.
	 * After the execution callback is called, if it is not undefined.
	 * If isClose === true, then after execution the plugin will close.
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
	 * @description The function allows to load a text resource located remotely on a link.
	 * After the execution callback is called with resource content if it is not undefined.
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
