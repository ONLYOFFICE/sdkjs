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

(function(window, undefined){

	/***********************************************************************
	 * CONFIG
	 */

	/**
	 * @typedef {Object} Config
	 * @property {string} guid
	 * @property {string} name
	 */

	/***********************************************************************
	 * API
	 */

	/**
	 * @global
	 * @class
	 * @name Plugin
	 */
	var Plugin = window["Asc"]["plugin"];

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
	 * @typedef {Function} IH_createWindow
	 */
	/**
	 * @typedef {Function} IH_getItems
	 * @returns {InputHelperItem[]}
	 */
	/**
	 * @typedef {Function} IH_setItems
	 * @param {InputHelperItem[]}
	 */
	/**
	 * @typedef {Function} IH_show
	 * @param {number} width
	 * @param {number} height
	 * @param {boolean} isCaptureKeyboard
	 */
	/**
	 * @typedef {Function} IH_unShow
	 */
	/**
	 * @typedef {Function} IH_getScrollSizes
	 * @return {number}
	 */

	/**
	 * @typedef {Object} InputHelper
	 * @property {IH_createWindow} createWindow Create window of the input helper
	 * @property {IH_getItems} getItems Return array of items
	 * @property {IH_setItems} setItems Set array of items
	 * @property {IH_show} show Show window
	 * @property {IH_unShow} unShow Hide window
	 * @property {IH_getScrollSizes} getScrollSizes Get full size of the input helper
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
		window.Asc.plugin.ih = new window.Asc.inputHelper(window.Asc.plugin);
	};

	/***********************************************************************
	 * EVENTS
	 */

	function events()
	{
		/**
		 * Event: init
		 * @memberof Plugin
		 * @alias init
		 */
		Plugin.init = function() {};

		/**
		 * @memberof Plugin
		 * @property {event_init} init
		 * @property {event_button} button
		 *
		 * @property {event_onTargetPositionChanged} event_onTargetPositionChanged
		 * @property {event_onClick} event_onClick
		 * @property {event_onDocumentContentReady} event_onDocumentContentReady
		 *
		 * @property {event_inputHelper_onSelectItem} inputHelper_onSelectItem
		 * @property {event_onInputHelperClear} event_onInputHelperClear
		 * @property {event_onInputHelperInput} event_onInputHelperInput
		 */
	}

})(window, undefined);
