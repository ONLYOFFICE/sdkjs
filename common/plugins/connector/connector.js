/**
 *
 * (c) Copyright Ascensio System SIA 2020
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */

(function(exports){

	function generateGuid() {
		if (!window.crypto || !window.crypto.getRandomValues) {
			function s4() { return Math.floor((1 + Math.random()) * 0x10000).toString(16).substring(1); }
			return s4() + s4() + '-' + s4() + '-' + s4() + '-' + s4() + '-' + s4() + s4() + s4();
		} else {
			var array = new Uint16Array(8);
			window.crypto.getRandomValues(array);
			var index = 0;
			function s4() { var value = 0x10000 + array[index++]; return value.toString(16).substring(1); }
			return s4() + s4() + '-' + s4() + '-' + s4() + '-' + s4() + '-' + s4() + s4() + s4();
		}
	}

	function EditorConnector(frameId, autoconnect) {
		this.frameId = frameId;
		this.guid = "asc.{" + generateGuid() + "}";

		if (autoconnect)
			this.connect();

		this.callbacks = [];
		this.events = {};
		this.tasks = [];

		this.onMessageBound = this.onMessage.bind(this);
	}

	EditorConnector.prototype.onMessage = function(e) {

		if (typeof(e.data) == "string") {
			var pluginData = {};
			try {
				pluginData = JSON.parse(e.data);
			} catch(err) {
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

	EditorConnector.prototype.sendMessage = function(message) {
		var e = {};
		e.frameEditorId = "iframeEditor";
		e.type = "onExternalPluginMessage";
		e.subType = "connector";
		e.data = message;
		e.data.guid = this.guid;

		var frame = document.getElementById(this.frameId);
		if (frame)
			frame.contentWindow.postMessage(JSON.stringify(e), "*");
	};

	EditorConnector.prototype.connect = function() {
		if (window.addEventListener)
			window.addEventListener("message", this.onMessageBound, false);
		else if (window.attachEvent)
			window.attachEvent("onmessage", this.onMessageBound);

		this.sendMessage({ type : "register" });
	};

	EditorConnector.prototype.disconnect = function() {
		if (window.removeEventListener)
			window.removeEventListener("message", this.onMessageBound, false);
		else if (window.detachEvent)
			window.detachEvent("onmessage", this.onMessageBound);

		this.sendMessage({ type : "unregister" });
	};

	EditorConnector.prototype.callCommand = function(command, callback, scope, isNoCalc) {

		this.callbacks.push(callback);
		var txtFunc = "var Asc = {}; Asc.scope = " + JSON.stringify(scope || {}) + "; var scope = Asc.scope; (" + command.toString() + ")();";

		var message = {
			type : "command",
			recalculate : isNoCalc ? false : true,
			data : txtFunc
		};

		if (this.callbacks.length !== 1) {
			this.tasks.push(message);
			return;
		}

		this.sendMessage(message);
	};

	EditorConnector.prototype.callMethod = function(name, params, callback) {

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

	EditorConnector.prototype.attachEvent = function(name, callback) {
		this.events[name] = callback;
		this.sendMessage({
			type : "attachEvent",
			name : name
		});
	};

	EditorConnector.prototype.detachEvent = function(name) {
		if (!this.events[name])
			return;

		delete this.events[name];
		this.sendMessage({
			type : "detachEvent",
			name : name
		});
	};

	window.Asc = window.Asc ? window.Asc : {};
	window.Asc.EditorConnector = EditorConnector;

})(window);
