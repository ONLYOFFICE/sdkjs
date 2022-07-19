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

(function(exports) {

	function generateGuid() {
		if (!window.crypto || !window.crypto.getRandomValues) {
			function s4() {
				return Math.floor((1 + Math.random()) * 0x10000).toString(16).substring(1);
			}

			return s4() + s4() + '-' + s4() + '-' + s4() + '-' + s4() + '-' + s4() + s4() + s4();
		} else {
			var array = new Uint16Array(8);
			window.crypto.getRandomValues(array);
			var index = 0;

			function s4() {
				var value = 0x10000 + array[index++];
				return value.toString(16).substring(1);
			}

			return s4() + s4() + '-' + s4() + '-' + s4() + '-' + s4() + '-' + s4() + s4() + s4();
		}
	}

	function EditorConnector(config) {
		this.frame = config.frame;
		this.guid = "asc.{" + generateGuid() + "}";
		this.isConnected = false;

		this.callbacks = [];
		this.events = {};
		this.tasks = [];

		this.onMessageBound = this.onMessage.bind(this);

		if (config.autoconnect)
			this.connect();
	}

	EditorConnector.prototype.onMessage = function (e) {

		if (typeof (e.data) == "string") {
			var pluginData = {};
			try {
				pluginData = JSON.parse(e.data);
			} catch (err) {
				pluginData = {};
			}

			if ("onExternalPluginMessageCallback" !== pluginData.type)
				return;

			pluginData = pluginData.data;

			if (this.guid !== pluginData.guid)
				return;

			switch (pluginData.type) {
				case "onMethodReturn": {
					if (this.callbacks.length > 0) {
						var callback = this.callbacks.shift();
						callback && callback(pluginData.methodReturnData);
					}

					if (this.tasks.length > 0) {
						this.sendMessage(this.tasks.shift());
					}
					break;
				}
				case "onCommandCallback": {
					if (this.callbacks.length > 0) {
						var callback = this.callbacks.shift();
						callback && callback();
					}

					if (this.tasks.length > 0) {
						this.sendMessage(this.tasks.shift());
					}
					break;
				}
				case "onEvent": {
					if (pluginData.eventName && this.events[pluginData.eventName])
						this.events[pluginData.eventName](pluginData.eventData);
					break;
				}
				default:
					break;
			}
		}

	};

	EditorConnector.prototype.sendMessage = function (message) {
		var e = {};
		e.frameEditorId = "iframeEditor";
		e.type = "onExternalPluginMessage";
		e.subType = "connector";
		e.data = message;
		e.data.guid = this.guid;

		var frame = this.frame;
		if (typeof frame === "string")
			frame = document.getElementById(this.frame);

		if (frame)
			frame.contentWindow.postMessage(JSON.stringify(e), "*");
	};

	EditorConnector.prototype.connect = function () {
		if (window.addEventListener)
			window.addEventListener("message", this.onMessageBound, false);
		else if (window.attachEvent)
			window.attachEvent("onmessage", this.onMessageBound);
		this.isConnected = true;

		this.sendMessage({type: "register"});
	};

	EditorConnector.prototype.disconnect = function () {
		if (window.removeEventListener)
			window.removeEventListener("message", this.onMessageBound, false);
		else if (window.detachEvent)
			window.detachEvent("onmessage", this.onMessageBound);
		this.isConnected = false;

		this.sendMessage({type: "unregister"});
	};

	/**
	 * callCommand
	 * @memberof EditorConnector
	 * @alias callCommand
	 * @description Defines the method used to send the data back to the editor.
	 * It allows the plugin to send structured data that can be inserted to the resulting document file (formatted paragraphs, tables, text parts and separate words, etc.).
	 * The callback is the result that the command returns. It is an optional parameter.
	 * ONLYOFFICE Document Builder commands can be only used to create content and insert it to the document editor (using the Api.GetDocument().InsertContent(...)).
	 * This limitation exists due to the co-editing feature in the online editors. If it is necessary to create a plugin for desktop editors to work with local files, no such limitation is applied.
	 * This method is executed in its own context isolated from other JavaScript data. If some parameters or other data need to be passed to this method, use Asc.scope object.
	 * @param {Function} command - Defines the command written in JavaScript which purpose is to form structured data which can be inserted to the resulting document file
	 * (formatted paragraphs, tables, text parts and separate words, etc.). Then the data is sent to the editors.
	 * The command must be compatible with ONLYOFFICE Document Builder syntax.
	 * @param {Function} callback - The result that the method returns.
	 * @param {boolean} isNoCalc - Defines whether the document will be recalculated or not.
	 * The true value is used to recalculate the document after executing the function in the command parameter.
	 * The false value will not recalculate the document (use it only when your edits surely will not require document recalculation).
	 */
	EditorConnector.prototype.callCommand = function (command, callback, isNoCalc) {
		if (!this.isConnected) {
			console.log("Connector is not connected with editor");
			return;
		}

		this.callbacks.push(callback);
		var txtFunc = "var Asc = {}; Asc.scope = " + JSON.stringify(window.Asc.scope || {}) + "; var scope = Asc.scope; (" + command.toString() + ")();";

		var message = {
			type : "command",
			recalculate : (isNoCalc === true) ? false : true,
			data : txtFunc
		};

		if (this.callbacks.length !== 1) {
			this.tasks.push(message);
			return;
		}

		this.sendMessage(message);
	};

	/**
	 * executeMethod
	 * @memberof EditorConnector
	 * @alias executeMethod
	 * @description Defines the method used to execute certain editor methods using the connector.
	 * The callback is the result that the method returns. It is an optional parameter.
	 * @param {string} name - The name of the specific method that must be executed.
	 * @param {Array} params - The arguments that the method in use has (if it has any).
	 * @param {Function} callback - The result that the method returns.
	 */
	EditorConnector.prototype.executeMethod = function(name, params, callback) {
		if (!this.isConnected) {
			console.log("Connector is not connected with editor");
			return;
		}

		this.callbacks.push(callback);

		var message = {
			type : "method",
			methodName : name,
			data : params
		};

		if (this.callbacks.length !== 1) {
			this.tasks.push(message);
			return;
		}

		this.sendMessage(message);
	};

	/**
	 * attachEvent
	 * @memberof EditorConnector
	 * @alias attachEvent
	 * @description Adds an event listener, a function that will be called whenever the specified event is delivered to the target.
	 * @param {string} name - The event name.
	 * @param {Function} callback - The event listener.
	 */
	EditorConnector.prototype.attachEvent = function(name, callback) {
		if (!this.isConnected) {
			console.log("Connector is not connected with editor");
			return;
		}

		this.events[name] = callback;
		this.sendMessage({
			type : "attachEvent",
			name : name
		});
	};

	/**
	 * detachEvent
	 * @memberof EditorConnector
	 * @alias detachEvent
	 * @description Removes an event listener.
	 * @param {string} name - The event name.
	 */
	EditorConnector.prototype.detachEvent = function(name) {
		if (!this.events[name])
			return;

		delete this.events[name];

		if (!this.isConnected) {
			console.log("Connector is not connected with editor");
			return;
		}

		this.sendMessage({
			type : "detachEvent",
			name : name
		});
	};

	exports.Asc = exports.Asc ? exports.Asc : {};
	exports.Asc.EditorConnector = EditorConnector;

})(window);
