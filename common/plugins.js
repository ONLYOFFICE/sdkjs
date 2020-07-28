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

(function(window, undefined)
{

	function CPluginData()
	{
		this.privateData = {};
	}

	CPluginData.prototype =
	{
		setAttribute : function(name, value)
		{
			this.privateData[name] = value;
		},

		getAttribute : function(name)
		{
			return this.privateData[name];
		},

		serialize : function()
		{
			var _data = "";
			try
			{
				_data = JSON.stringify(this.privateData);
			}
			catch (err)
			{
				_data = "{ \"data\" : \"\" }";
			}
			return _data;
		},

		deserialize : function(_data)
		{
			try
			{
				this.privateData = JSON.parse(_data);
			}
			catch (err)
			{
				this.privateData = {"data" : ""};
			}
		}
	};

	function CPluginsManager(api)
	{
		this.plugins          = [];
		this.systemPlugins	  = [];

		this.runnedPluginsMap = {}; // guid => { iframeId: "", currentVariation: 0, currentInit: false, isSystem: false, startData: {}, closeAttackTimer: -1, methodReturnAsync: false }
		this.pluginsMap = {};		// guid => { isSystem: false }

		this.path             = "../../../../sdkjs-plugins/";
		this.systemPath 	  = "";
		this.api              = api;
		this["api"]			  = this.api;

		this.runAndCloseData = null;

		this.isNoSystemPluginsOnlyOne = true;

		this.countEventDocContOrPluginsReady = 0;

		this.guidAsyncMethod = "";

		this.sendsToInterface = {};

		this.sendEncryptionDataCounter = 0;

		this.language = "en-EN";

		if (this.api.isCheckCryptoReporter)
			this.checkCryptoReporter();
	}

	CPluginsManager.prototype =
	{
		unregisterAll : function()
		{
			// удаляем все, кроме запущенного
			var i = 0;
			for (i = 0; i < this.plugins.length; i++)
			{
				if (!this.runnedPluginsMap[this.plugins[i].guid])
				{
					delete this.pluginsMap[this.plugins[i].guid];
					this.plugins.splice(i, 1);
					i--;
				}
			}
		},

		register : function(basePath, plugins)
		{
			this.path = basePath;

			for (var i = 0; i < plugins.length; i++)
			{
				var guid = plugins[i].guid;
				var isSystem = false;
				if (plugins[i].variations && plugins[i].variations[0] && plugins[i].variations[0].isSystem)
					isSystem = true;

				if (this.runnedPluginsMap[guid])
				{
					// не меняем запущенный
					continue;
				}
				else if (this.pluginsMap[guid])
				{
					// заменяем новым
					for (var j = 0; j < this.plugins.length; j++)
					{
						if (this.plugins[j].guid == guid)
						{
							this.plugins[j] = plugins[i];
							break;
						}
					}
				}
				else
				{
					if (!isSystem)
						this.plugins.push(plugins[i]);
					else
						this.systemPlugins.push(plugins[i]);

					this.pluginsMap[guid] = { isSystem : isSystem };
				}

				if (isSystem)
				{
					this.run(guid, 0, "");
				}
			}
		},
		registerSystem : function(basePath, plugins)
		{
			this.systemPath = basePath;

			for (var i = 0; i < plugins.length; i++)
			{
				var guid = plugins[i].guid;

				// системные не обновляем
				if (this.pluginsMap[guid])
				{
					continue;
				}

				this.systemPlugins.push(plugins[i]);
				this.pluginsMap[guid] = { isSystem : true };
			}
		},
		runAllSystem : function()
		{
			for (var i = 0; i < this.systemPlugins.length; i++)
			{
				this.run(this.systemPlugins[i].guid, 0, "");
			}
		},
		// pointer events methods -------------------
		enablePointerEvents : function()
		{
			for (var guid in this.runnedPluginsMap)
			{
				var _frame = document.getElementById(this.runnedPluginsMap[guid].frameId);
				if (_frame)
					_frame.style.pointerEvents = "";
			}
		},
		disablePointerEvents : function()
		{
			for (var guid in this.runnedPluginsMap)
			{
				var _frame = document.getElementById(this.runnedPluginsMap[guid].frameId);
				if (_frame)
					_frame.style.pointerEvents = "none";
			}
		},
		// ------------------------------------------
		checkRunnedFrameId : function(id)
		{
			for (var guid in this.runnedPluginsMap)
			{
				if (this.runnedPluginsMap[guid].frameId == id)
					return true;
			}
			return false;
		},
		sendToAllPlugins : function(data)
		{
			for (var guid in this.runnedPluginsMap)
			{
				var _frame = document.getElementById(this.runnedPluginsMap[guid].frameId);
				if (_frame)
					_frame.contentWindow.postMessage(data, "*");
			}
		},
		getPluginByGuid : function(guid)
		{
			if (undefined === this.pluginsMap[guid])
				return null;

			var _array = (this.pluginsMap[guid].isSystem) ? this.systemPlugins : this.plugins;
			for (var i = _array.length - 1; i >= 0; i--)
			{
				if (_array[i].guid == guid)
					return _array[i];
			}
			return null;
		},
		isWorked : function()
		{
			for (var i in this.runnedPluginsMap)
			{
				if (this.pluginsMap[i] && !this.pluginsMap[i].isSystem)
				{
					return true;
				}
			}
			return false;
		},
		stopWorked : function()
		{
		   for (var i in this.runnedPluginsMap)
		   {
			   if (this.pluginsMap[i] && !this.pluginsMap[i].isSystem)
			   {
					this.close(i);
			   }
		   }
		},
		isRunned : function(guid)
		{
			return (undefined !== this.runnedPluginsMap[guid]);
		},
		checkEditorSupport : function(plugin, variation)
		{
			var typeEditor = this.api.getEditorId();
			var typeEditorString = "";
			switch (typeEditor)
			{
				case AscCommon.c_oEditorId.Word:
					typeEditorString = "word";
					break;
				case AscCommon.c_oEditorId.Presentation:
					typeEditorString = "slide";
					break;
				case AscCommon.c_oEditorId.Spreadsheet:
					typeEditorString = "cell";
					break;
				default:
					break;
			}
			var runnedVariation = variation ? variation : 0;
			if (!plugin.variations[runnedVariation] ||
				!plugin.variations[runnedVariation].EditorsSupport ||
				!plugin.variations[runnedVariation].EditorsSupport.includes(typeEditorString))
				return false;
			return true;
		},
		run : function(guid, variation, data, isNoUse_isNoSystemPluginsOnlyOne)
		{
            var isEnabled = this.api.DocInfo ? this.api.DocInfo.get_IsEnabledPlugins() : true;
			if (false === isEnabled)
				return;

			if (this.runAndCloseData) // run only on close!!!
				return;

			if (this.pluginsMap[guid] === undefined)
				return;

			var plugin = this.getPluginByGuid(guid);
			if (!plugin)
				return;

			if (!this.checkEditorSupport(plugin, variation))
				return;

			var isSystem = this.pluginsMap[guid].isSystem;
			var isRunned = (this.runnedPluginsMap[guid] !== undefined) ? true : false;

			if (isRunned && ((variation == null) || variation == this.runnedPluginsMap[guid].currentVariation))
			{
				// запуск запущенного => закрытие
				this.close(guid);
				return false;
			}

			if ((isNoUse_isNoSystemPluginsOnlyOne !== true) && !isSystem && this.isNoSystemPluginsOnlyOne)
			{
				// смотрим, есть ли запущенный несистемный плагин
				var guidOther = "";
				for (var i in this.runnedPluginsMap)
				{
					if (this.pluginsMap[i] && !this.pluginsMap[i].isSystem)
					{
						guidOther = i;
						break;
					}
				}

				if (guidOther != "")
				{
					// стопим текущий, а после закрытия - стартуем новый.
					this.runAndCloseData = {};
					this.runAndCloseData.guid = guid;
					this.runAndCloseData.variation = variation;
					this.runAndCloseData.data = data;

					this.close(guidOther);
					return;
				}
			}

			var _startData = (data == null || data == "") ? new CPluginData() : data;
			_startData.setAttribute("guid", guid);
			this.correctData(_startData);

			this.runnedPluginsMap[guid] = {
				frameId: "iframe_" + guid,
				currentVariation: Math.min(variation, plugin.variations.length - 1),
				currentInit: false,
				isSystem: isSystem,
				startData: _startData,
				closeAttackTimer: -1,
				methodReturnAsync: false
			};

			this.show(guid);
		},
		runResize : function(data)
		{
			var guid = data.getAttribute("guid");
			var plugin = this.getPluginByGuid(guid);

			if (!plugin)
				return;

			if (true !== plugin.variations[0].isUpdateOleOnResize)
				return;

			data.setAttribute("resize", true);
			return this.run(guid, 0, data, true);
		},
		close : function(guid)
		{
			var plugin = this.getPluginByGuid(guid);
			var runObject = this.runnedPluginsMap[guid];
			if (!plugin || !runObject)
				return;

			if (runObject.startData && runObject.startData.getAttribute("resize") === true)
				this.endLongAction();

			runObject.startData = null;

			if (true)
			{
				if (this.sendsToInterface[plugin.guid])
				{
					this.api.sendEvent("asc_onPluginClose", plugin, runObject.currentVariation);
					delete this.sendsToInterface[plugin.guid];
				}
				var _div = document.getElementById(runObject.frameId);
				if (_div)
					_div.parentNode.removeChild(_div);
			}

			delete this.runnedPluginsMap[guid];

			if (this.runAndCloseData)
			{
				var _tmp = this.runAndCloseData;
				this.runAndCloseData = null;
				this.run(_tmp.guid, _tmp.variation, _tmp.data);
			}
		},

		show : function(guid)
		{
			var plugin = this.getPluginByGuid(guid);
			var runObject = this.runnedPluginsMap[guid];

			if (!plugin || !runObject)
				return;

			if (runObject.startData.getAttribute("resize") === true)
				this.startLongAction();

		    if (this.api.WordControl && this.api.WordControl.m_oTimerScrollSelect != -1)
		    {
		        clearInterval(this.api.WordControl.m_oTimerScrollSelect);
                this.api.WordControl.m_oTimerScrollSelect = -1;
		    }

			if (plugin.variations[runObject.currentVariation].isVisual && runObject.startData.getAttribute("resize") !== true)
			{
				this.api.sendEvent("asc_onPluginShow", plugin, runObject.currentVariation, runObject.frameId, "?lang=" + this.language);
				this.sendsToInterface[plugin.guid] = true;
			}
			else
			{
				var ifr            = document.createElement("iframe");
				ifr.name           = runObject.frameId;
				ifr.id             = runObject.frameId;
				var _add           = plugin.baseUrl == "" ? this.path : plugin.baseUrl;
				ifr.src            = _add + plugin.variations[runObject.currentVariation].url + "?lang=" + this.language;
				ifr.style.position = (AscCommon.AscBrowser.isIE || AscCommon.AscBrowser.isMozilla) ? 'fixed' : "absolute";
				ifr.style.top      = '-100px';
				ifr.style.left     = '0px';
				ifr.style.width    = '10000px';
				ifr.style.height   = '100px';
				ifr.style.overflow = 'hidden';
				ifr.style.zIndex   = -1000;
				ifr.setAttribute("frameBorder", "0");
				ifr.setAttribute("allow", "autoplay");
				document.body.appendChild(ifr);

				if (runObject.startData.getAttribute("resize") !== true)
				{
					this.api.sendEvent("asc_onPluginShow", plugin, runObject.currentVariation);
					this.sendsToInterface[plugin.guid] = true;
				}
			}

			runObject.currentInit = false;
		},

		buttonClick : function(id, guid)
		{
			if (guid === undefined)
			{
				// old version support
				for (var i in this.runnedPluginsMap)
				{
					if (this.runnedPluginsMap[i].isSystem)
						continue;
					
					if (this.pluginsMap[i])
					{
						guid = i;
						break;
					}
				}
			}

			if (undefined === guid)
				return;

			var plugin = this.getPluginByGuid(guid);
			var runObject = this.runnedPluginsMap[guid];

			if (!plugin || !runObject)
				return;

			if (runObject.closeAttackTimer != -1)
			{
				clearTimeout(runObject.closeAttackTimer);
				runObject.closeAttackTimer = -1;
			}

			if (-1 == id)
			{
				if (!runObject.currentInit)
				{
					this.close(guid);
				}

				// защита от плохого плагина
				runObject.closeAttackTimer = setTimeout(function()
				{
					window.g_asc_plugins.close();
				}, 5000);
			}
			var _iframe = document.getElementById(runObject.frameId);
			if (_iframe)
			{
				var pluginData = new CPluginData();
				pluginData.setAttribute("guid", plugin.guid);
				pluginData.setAttribute("type", "button");
				pluginData.setAttribute("button", "" + id);
				_iframe.contentWindow.postMessage(pluginData.serialize(), "*");
			}
		},

		init : function(guid, raw_data)
		{
			var plugin = this.getPluginByGuid(guid);
			var runObject = this.runnedPluginsMap[guid];

			if (!plugin || !runObject || !runObject.startData)
				return;

			if (undefined === raw_data)
			{
				switch (plugin.variations[runObject.currentVariation].initDataType)
				{
					case Asc.EPluginDataType.text:
					{
						var text_data = {
							data:     "",
							pushData: function (format, value)
									  {
										  this.data = value;
									  }
						};

						this.api.asc_CheckCopy(text_data, 1);
						if (text_data.data == null)
							text_data.data = "";
						runObject.startData.setAttribute("data", text_data.data);
						break;
					}
					case Asc.EPluginDataType.html:
					{
						var text_data = {
							data:     "",
							pushData: function (format, value)
									  {
										  this.data = value;
									  }
						};

						this.api.asc_CheckCopy(text_data, 2);
						if (text_data.data == null)
							text_data.data = "";
						runObject.startData.setAttribute("data", text_data.data);
						break;
					}
					case Asc.EPluginDataType.ole:
					{
						// теперь выше задается
						break;
					}
					case Asc.EPluginDataType.desktop:
					{
						if (plugin.variations[runObject.currentVariation].initData == "encryption")
						{
							if (this.api.isReporterMode)
							{
                                this.sendEncryptionDataCounter++;
                                if (2 <= this.sendEncryptionDataCounter)
                                {
                                    runObject.startData.setAttribute("data", {
                                        "type": "setPassword",
                                        "password": this.api.currentPassword,
                                        "hash": this.api.currentDocumentHash,
                                        "docinfo": this.api.currentDocumentInfo
                                    });
                                }
                            }

                            // for crypt mode (end waiting all system plugins)
                            if (this.api.asc_initAdvancedOptions_params)
                            {
                            	window["asc_initAdvancedOptions"].apply(window, this.api.asc_initAdvancedOptions_params);
                                delete this.api.asc_initAdvancedOptions_params;
                                // already sended in asc_initAdvancedOptions
                                return;
                            }
						}
						break;
					}
				}
			}
			else
			{
				runObject.startData.setAttribute("data", raw_data);
			}

			var _iframe = document.getElementById(runObject.frameId);
			if (_iframe)
			{
				runObject.startData.setAttribute("type", "init");
				_iframe.contentWindow.postMessage(runObject.startData.serialize(), "*");
			}

			runObject.currentInit = true;
		},
		correctData : function(pluginData)
		{
			pluginData.setAttribute("editorType", this.api._editorNameById());

			var _mmToPx = AscCommon.g_dKoef_mm_to_pix;
			if (this.api.WordControl && this.api.WordControl.m_nZoomValue)
				_mmToPx *= this.api.WordControl.m_nZoomValue / 100;

			pluginData.setAttribute("mmToPx", _mmToPx);

			if (undefined == pluginData.getAttribute("data"))
				pluginData.setAttribute("data", "");

            pluginData.setAttribute("isViewMode", this.api.isViewMode);
            pluginData.setAttribute("isMobileMode", this.api.isMobileVersion);
            pluginData.setAttribute("isEmbedMode", this.api.isEmbedVersion);
            pluginData.setAttribute("lang", this.language);
            pluginData.setAttribute("documentId", this.api.documentId);
            pluginData.setAttribute("documentTitle", this.api.documentTitle);

            if (this.api.User)
            {
                pluginData.setAttribute("userId", this.api.User.id);
                pluginData.setAttribute("userName", this.api.User.userName);
            }
		},
		loadExtensionPlugins : function(_plugins)
		{
			if (!_plugins || _plugins.length < 1)
				return;

			var _map = {};
			for (var i = 0; i < this.plugins.length; i++)
				_map[this.plugins[i].guid] = true;

			var _new = [];
			for (var i = 0; i < _plugins.length; i++)
			{
				var _p = new Asc.CPlugin();
				_p["deserialize"](_plugins[i]);

				if (_map[_p.guid] === true)
					continue;

				_new.push(_p);
			}

			this.register(this.path, _new);

			var _pluginsInstall = {"url" : this.path, "pluginsData" : []};
			for (var i = 0; i < this.plugins.length; i++)
			{
				_pluginsInstall["pluginsData"].push(this.plugins[i].serialize());
			}

			this.api.sendEvent("asc_onPluginsInit", _pluginsInstall);
		},

		startLongAction : function()
		{
			//console.log("startLongAction");
			this.api.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);
		},
		endLongAction   : function()
		{
			//console.log("endLongAction");
			this.api.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);
		},

		onChangedSelectionData : function()
		{
			for (var guid in this.runnedPluginsMap)
			{
				var plugin = this.getPluginByGuid(guid);
				var runObject = this.runnedPluginsMap[guid];

				if (plugin && plugin.variations[runObject.currentVariation].initOnSelectionChanged === true)
				{
					// re-init
					this.init(guid);
				}
			}
		},

        onPluginEvent : function(name, data)
        {
            for (var guid in this.runnedPluginsMap)
            {
                var plugin = this.getPluginByGuid(guid);
                var runObject = this.runnedPluginsMap[guid];

                if (plugin && plugin.variations[runObject.currentVariation].eventsMap[name])
                {
                    if (!runObject.isInitReceive)
                    {
                        if (!runObject.waitEvents)
                            runObject.waitEvents = [];
                        runObject.waitEvents.push({ n : name, d : data });
                        continue;
                    }
                    var pluginData = new CPluginData();
                    pluginData.setAttribute("guid", plugin.guid);
                    pluginData.setAttribute("type", "onEvent");
                    pluginData.setAttribute("eventName", name);
                    pluginData.setAttribute("eventData", data);
                    var _iframe = document.getElementById(runObject.frameId);
                    if (_iframe)
                        _iframe.contentWindow.postMessage(pluginData.serialize(), "*");
                }
            }
        },

        onPluginEvent2 : function(name, data, guids)
        {
            for (var guid in this.runnedPluginsMap)
            {
                var plugin = this.getPluginByGuid(guid);
                var runObject = this.runnedPluginsMap[guid];

                if (plugin && guids[guid])
                {
                    if (!runObject.isInitReceive)
                    {
                        if (!runObject.waitEvents)
                            runObject.waitEvents = [];
                        runObject.waitEvents.push({ n : name, d : data });
                        continue;
                    }
                    var pluginData = new CPluginData();
                    pluginData.setAttribute("guid", plugin.guid);
                    pluginData.setAttribute("type", "onEvent");
                    pluginData.setAttribute("eventName", name);
                    pluginData.setAttribute("eventData", data);
                    var _iframe = document.getElementById(runObject.frameId);
                    if (_iframe)
                        _iframe.contentWindow.postMessage(pluginData.serialize(), "*");
                }
            }
        },

		onExternalMouseUp : function()
		{
			for (var guid in this.runnedPluginsMap)
			{
				var runObject = this.runnedPluginsMap[guid];
				runObject.startData.setAttribute("type", "onExternalMouseUp");
				this.correctData(runObject.startData);

				var _iframe = document.getElementById(runObject.frameId);
				if (_iframe)
				{
					runObject.startData.setAttribute("guid", guid);
					_iframe.contentWindow.postMessage(runObject.startData.serialize(), "*");
				}
			}
		},

		onEnableMouseEvents : function(isEnable)
		{
			for (var guid in this.runnedPluginsMap)
			{
				var runObject = this.runnedPluginsMap[guid];

				var _pluginData = new Asc.CPluginData();
				_pluginData.setAttribute("type", "enableMouseEvent");
				_pluginData.setAttribute("isEnabled", isEnable);
				this.correctData(_pluginData);

				var _iframe = document.getElementById(runObject.frameId);
				if (_iframe)
				{
					_pluginData.setAttribute("guid", guid);
					_iframe.contentWindow.postMessage(_pluginData.serialize(), "*");
				}
			}
		},

		onPluginMethodReturn : function(guid, _return)
		{
			var plugin = this.getPluginByGuid(guid);
			var runObject = this.runnedPluginsMap[guid];

			if (!plugin || !runObject)
				return;

			var pluginData = new CPluginData();
			pluginData.setAttribute("guid", plugin.guid);
			pluginData.setAttribute("type", "onMethodReturn");
			pluginData.setAttribute("methodReturnData", _return);
			var _iframe = document.getElementById(runObject.frameId);
			if (_iframe)
				_iframe.contentWindow.postMessage(pluginData.serialize(), "*");
		},

		setPluginMethodReturnAsync : function()
		{
			if (this.runnedPluginsMap[this.guidAsyncMethod])
				this.runnedPluginsMap[this.guidAsyncMethod].methodReturnAsync = true;
			return this.guidAsyncMethod;
		},

        /* encryption methods ------------- */
        getEncryption : function()
        {
            var _count = this.plugins.length;
            var i = 0;
            for (i = 0; i < _count; i++)
            {
                var _variation = this.plugins[i].variations[0];
                if (_variation)
                {
                    if ("desktop" == _variation.initDataType && "encryption" == _variation.initData)
                        return this.plugins[i];
                }
            }

            _count = this.systemPlugins.length;
            for (i = 0; i < _count; i++)
            {
                var _variation = this.systemPlugins[i].variations[0];
                if (_variation)
                {
                    if ("desktop" == _variation.initDataType && "encryption" == _variation.initData)
                        return this.systemPlugins[i];
                }
            }

            return null;
        },
        isRunnedEncryption : function()
        {
            var _plugin = this.getEncryption();
            if (!_plugin)
            	return false;
            return this.isRunned(_plugin.guid);
        },
        sendToEncryption : function(data)
        {
            var _plugin = this.getEncryption();
            if (!_plugin)
            	return;
            this.init(_plugin.guid, data);
        },
        checkCryptoReporter : function()
        {
            this.sendEncryptionDataCounter++;
            if (2 <= this.sendEncryptionDataCounter)
            {
                this.sendToEncryption({
                    "type" : "setPassword",
                    "password" : this.api.currentPassword,
                    "hash" : this.api.currentDocumentHash,
                    "docinfo" : this.api.currentDocumentInfo
                });
            }
        },

		checkOrigin : function(guid, event)
		{
			if (event.origin === window.origin)
				return true;

			// allow chrome extensions
			if (0 === event.origin.indexOf("chrome-extension://"))
				return true;

			// external plugins
			var plugin = this.getPluginByGuid(guid);
			if (plugin && 0 === plugin.baseUrl.indexOf(event.origin))
				return true;

			return false;
		}
        /* -------------------------------- */
	};

	// export
	CPluginsManager.prototype["buttonClick"] = CPluginsManager.prototype.buttonClick;

	function onMessage(event)
	{
		if (!window.g_asc_plugins)
			return;

		if (typeof(event.data) != "string")
			return;

		var pluginData = new CPluginData();
		pluginData.deserialize(event.data);

		var guid = pluginData.getAttribute("guid");
		var runObject = window.g_asc_plugins.runnedPluginsMap[guid];

		if (!runObject)
			return;

		// check origin
		if (!window.g_asc_plugins.checkOrigin(guid, event))
			return;

		var name  = pluginData.getAttribute("type");
		var value = pluginData.getAttribute("data");

		if ("initialize_internal" == name)
		{
			window.g_asc_plugins.init(guid);

			runObject.isInitReceive = true;

			setTimeout(function() {
				if (runObject.waitEvents)
				{
					for (var i = 0; i < runObject.waitEvents.length; i++)
					{
						var pluginData = new CPluginData();
						pluginData.setAttribute("guid", guid);
						pluginData.setAttribute("type", "onEvent");
						pluginData.setAttribute("eventName", runObject.waitEvents[i].n);
						pluginData.setAttribute("eventData", runObject.waitEvents[i].d);
						var _iframe = document.getElementById(runObject.frameId);
						if (_iframe)
							_iframe.contentWindow.postMessage(pluginData.serialize(), "*");
					}
					runObject.waitEvents = null;
				}
			}, 100);
		}
		else if ("initialize" == name)
		{
			var pluginData = new CPluginData();
			pluginData.setAttribute("guid", guid);
			pluginData.setAttribute("type", "plugin_init");
			pluginData.setAttribute("data", /*<code>*/"(function(a,g){function f(a){this.plugin=a;this.ps;this.items=[];this.isCurrentVisible=this.isVisible=!1}f.prototype.createWindow=function(){var a=document.body,b=document.getElementsByTagName(\"head\")[0];a&&b&&(a=document.createElement(\"style\"),a.type=\"text/css\",a.innerHTML='.ih_main { margin: 0px; padding: 0px; width: 100%; height: 100%; display: inline-block; overflow: hidden; box-sizing: border-box; user-select: none; position: fixed; border: 1px solid #cfcfcf; } ul { margin: 0px; padding: 0px; width: 100%; height: 100%; list-style-type: none; outline:none; } li { padding: 5px; font-family: \"Helvetica Neue\", Helvetica, Arial, sans-serif; font-size: 12px; font-weight: 400; color: #373737; } li:hover { background-color: #D8DADC; } .li_selected { background-color: #D8DADC; color: #373737; }.li_selected:hover { background-color: #D8DADC; color: #373737; }',b.appendChild(a),document.body.style.background=\"#FFFFFF\",document.body.style.width=\"100%\",document.body.style.height=\"100%\",document.body.style.margin=\"0\",document.body.style.padding=\"0\",document.body.innerHTML='<div class=\"ih_main\" id=\"ih_area\"><ul id=\"ih_elements_id\" role=\"listbox\"></ul></div>',this.ps=new PerfectScrollbar(document.getElementById(\"ih_area\"),{minScrollbarLength:20}),this.updateScrolls(),this.createDefaultEvents())};f.prototype.setItems=function(a){this.items=a;for(var b=\"\",d=a.length,c=0;c<d;c++)g===a[c].id&&(a[c].id=\"\"+c),b+='<li role=\"option\"',0==c&&(b+=' class=\"li_selected\"'),b+=' id=\"'+a[c].id+'\"',b+=' onclick=\"_private_on_ih_click(event)\">',b+=a[c].text,b+=\"</li>\";document.getElementById(\"ih_elements_id\").innerHTML=b;this.updateScrolls();this.scrollToSelected()};f.prototype.createDefaultEvents=function(){this.plugin.onExternalMouseUp=function(){var b=document.createEvent(\"MouseEvents\");b.initMouseEvent(\"mouseup\",!0,!0,a,1,0,0,0,0,!1,!1,!1,!1,0,null);document.dispatchEvent(b)};var e=this;a.onkeydown=function(a){switch(a.keyCode){case 27:e.isVisible&&(e.isVisible=!1,e.plugin.executeMethod(\"UnShowInputHelper\",[e.plugin.info.guid,!0]));break;case 38:case 40:case 9:case 36:case 35:case 33:case 34:for(var d=document.getElementsByTagName(\"li\"),c=-1,b=0;b<d.length;b++)if(\"li_selected\"==d[b].className){c=b;d[b].className=\"\";break}if(-1==c)c=0;else switch(a.keyCode){case 38:c--;0>c&&(c=0);break;case 40:c++;c>=d.length&&(c=d.length-1);break;case 9:c++;c>=d.length&&(c=0);break;case 36:c=0;break;case 35:c=d.length-1;break;case 33:case 34:b=1;var f=document.getElementById(\"ih_area\").clientHeight/24>>0;1<f&&(b=f);33==a.keyCode?(c-=b,0>c&&(c=0)):(c+=b,c>=d.length&&(c=c=d.length-1))}c<d.length&&(d[c].className=\"li_selected\",b=d[c].offsetTop,d=d[c].offsetHeight,c=document.getElementById(\"ih_area\"),f=c.scrollTop,b<f?c.scrollTo?c.scrollTo(0,b):c.scrollTop=b:f+c.offsetHeight<b+d&&(c.scrollTo?c.scrollTo(0,b-(c.offsetHeight-d)):c.scrollTop=b-(c.offsetHeight-d)));break;case 13:e.onSelectedItem()}a.preventDefault&&a.preventDefault();a.stopPropagation&&a.stopPropagation();return!1};a.onresize=function(a){e.updateScrolls()};a._private_on_ih_click=function(a){for(var d=document.getElementsByTagName(\"li\"),c=0;c<d.length;c++)d[c].className=\"\";a.target.className=\"li_selected\";a.target.getAttribute(\"id\");e.onSelectedItem()};this.plugin.event_onKeyDown=function(b){a.onkeydown({keyCode:b.keyCode})}};f.prototype.updateScrolls=function(){this.ps.update();this.ps.update();var e=document.getElementsByClassName(\"ps__rail-y\")[0],b=document.getElementsByClassName(\"ps__rail-x\")[0];if(b&&e){var d=a.getComputedStyle(e),c=a.getComputedStyle(b);d=d&&\"none\"==d.display?!1:!0;c&&\"none\"==c.display||!d?(\"2px\"!=e.style.marginBottom&&(e.style.marginBottom=\"2px\"),\"2px\"!=b.style.marginRight&&(b.style.marginRight=\"2px\")):(\"13px\"!=e.style.marginBottom&&(e.style.marginBottom=\"13px\"),\"13px\"!=b.style.marginRight&&(b.style.marginRight=\"13px\"))}};f.prototype.scrollToSelected=function(){for(var a=document.getElementsByTagName(\"li\"),b=0;b<a.length;b++)if(\"li_selected\"==a[b].className){var d=document.getElementById(\"ih_area\");d.scrollTo?d.scrollTo(0,a[b].offsetTop):d.scrollTop=a[b].offsetTop;break}};f.prototype.getSelectedItem=function(){for(var a=document.getElementsByTagName(\"li\"),b=-1,d=0;d<a.length;d++)if(\"li_selected\"==a[d].className){b=a[d].getAttribute(\"id\");break}if(-1==b)return null;a=this.items.length;for(d=0;d<a;d++)if(b==this.items[d].id)return this.items[d];return null};f.prototype.onSelectedItem=function(){this.plugin.inputHelper_onSelectItem&&this.plugin.inputHelper_onSelectItem(this.getSelectedItem())};f.prototype.show=function(e,b,d){this.isCurrentVisible=!0;this.plugin.executeMethod(\"ShowInputHelper\",[this.plugin.info.guid,e,b,d],function(){a.Asc.plugin.ih.isVisible=!0})};f.prototype.unShow=function(){if(this.isCurrentVisible||this.isVisible)this.isCurrentVisible=!1,a.Asc.plugin.executeMethod(\"UnShowInputHelper\",[this.plugin.info.guid],function(){a.Asc.plugin.ih.isVisible=!1})};f.prototype.getItemHeight=function(){var a=24,b=document.getElementsByTagName(\"li\");0<b.length&&0<b[0].offsetHeight&&(a=b[0].offsetHeight);return a};f.prototype.getItemsHeight=function(a){return 2+a*this.getItemHeight()};f.prototype.getItems=function(){return this.items};f.prototype.getScrollSizes=function(){var a={w:0,h:0},b=this.getItemHeight(),d=document.getElementById(\"ih_elements_id\");d&&(a.w=d.scrollWidth,a.h=2+this.items.length*b);return a};a.Asc=a.Asc||{};a.Asc.inputHelper=f})(window,void 0);(function(a,g){function f(e){var b=e.metaKey||e.ctrlKey?!0:!1;if(116==e.keyCode)return a.parent.postMessage(JSON.stringify({type:\"reload\",guid:a.Asc.plugin.guid,ctrl:b}),\"*\"),e.preventDefault&&e.preventDefault(),e.stopPropagation&&e.stopPropagation(),!1}a.addEventListener?a.addEventListener(\"keydown\",f,!1):a.attachEvent(\"keydown\",f)})(window,void 0);(function(a,g){function f(){a.Asc.plugin.isStarted||(a.Asc.plugin.isStarted=!0,a.startPluginApi(),a.Asc.plugin.checkPixelRatio=function(d){if(!a.Asc.plugin.checkedPixelRatio||!0===d){a.Asc.plugin.checkedPixelRatio=!0;var c=navigator.userAgent.toLowerCase(),b=-1<c.indexOf(\"msie\")||-1<c.indexOf(\"trident\")||-1<c.indexOf(\"edge\");d=!b&&-1<c.indexOf(\"chrome\");!b&&c.indexOf(\"firefox\");b=c=1;var e=a.Asc.plugin.info?a.Asc.plugin.info.isMobileMode:!1;d&&document&&document.firstElementChild&&document.body&&!e?.1<a.devicePixelRatio?(1.99>a.devicePixelRatio?c=a.devicePixelRatio:(c=a.devicePixelRatio/2,b=2),document.firstElementChild.style.zoom=1/c):document.firstElementChild.style.zoom=\"normal\":((d=.01>Math.abs(2-a.devicePixelRatio))&&(b=2),e&&(b=a.devicePixelRatio));a.Asc.plugin.zoom=c;a.Asc.plugin.retinaPixelRatio=b}},a.Asc.plugin.checkPixelRatio())}var e=!1,b=\"\";a.plugin_sendMessage=function(b){a.parent.postMessage(b,\"*\")};a.plugin_onMessage=function(d){if(a.Asc.plugin&&\"string\"==typeof d.data){var c={};try{c=JSON.parse(d.data)}catch(l){c={}}d=c.type;if(c.guid!=a.Asc.plugin.guid){if(g!==c.guid)return;switch(d){case \"onExternalPluginMessage\":break;default:return}}\"init\"==d&&(a.Asc.plugin.info=c);a.Asc.plugin.tr&&a.Asc.plugin.tr_init||(a.Asc.plugin.tr_init=!0,a.Asc.plugin.tr=function(b){return a.Asc.plugin.translateManager&&a.Asc.plugin.translateManager[b]?a.Asc.plugin.translateManager[b]:b});var k=\"\";a.Asc.plugin.info&&(k=a.Asc.plugin.info.lang);if(\"\"==k||k!=b)if(b=k,\"en-EN\"==b||\"\"==b){if(a.Asc.plugin.translateManager={},a.Asc.plugin.onTranslate)a.Asc.plugin.onTranslate()}else{var h=new XMLHttpRequest;h.open(\"GET\",\"./translations/\"+b+\".json\");h.onreadystatechange=function(){if(4==h.readyState){if(200==h.status||0==location.href.indexOf(\"file:\"))try{a.Asc.plugin.translateManager=JSON.parse(h.responseText)}catch(l){a.Asc.plugin.translateManager={}}else a.Asc.plugin.translateManager={};if(a.Asc.plugin.onTranslate)a.Asc.plugin.onTranslate()}};h.send()}switch(d){case \"init\":f();a.Asc.plugin.init(a.Asc.plugin.info.data);break;case \"button\":c=parseInt(c.button);a.Asc.plugin.button||-1!=c?a.Asc.plugin.button(c):a.Asc.plugin.executeCommand(\"close\",\"\");break;case \"enableMouseEvent\":e=c.isEnabled;if(a.Asc.plugin.onEnableMouseEvent)a.Asc.plugin.onEnableMouseEvent(e);break;case \"onExternalMouseUp\":if(a.Asc.plugin.onExternalMouseUp)a.Asc.plugin.onExternalMouseUp();break;case \"onMethodReturn\":a.Asc.plugin.isWaitMethod=!1;if(a.Asc.plugin.methodCallback)d=a.Asc.plugin.methodCallback,a.Asc.plugin.methodCallback=null,d(c.methodReturnData),d=null;else if(a.Asc.plugin.onMethodReturn)a.Asc.plugin.onMethodReturn(c.methodReturnData);a.Asc.plugin.executeMethodStack&&0<a.Asc.plugin.executeMethodStack.length&&(c=a.Asc.plugin.executeMethodStack.shift(),a.Asc.plugin.executeMethod(c.name,c.params,c.callback));break;case \"onCommandCallback\":if(a.Asc.plugin.onCallCommandCallback)a.Asc.plugin.onCallCommandCallback(),a.Asc.plugin.onCallCommandCallback=null;else if(a.Asc.plugin.onCommandCallback)a.Asc.plugin.onCommandCallback();break;case \"onExternalPluginMessage\":if(a.Asc.plugin.onExternalPluginMessage&&c.data&&c.data.type)a.Asc.plugin.onExternalPluginMessage(c.data);break;case \"onEvent\":if(a.Asc.plugin[\"event_\"+c.eventName])a.Asc.plugin[\"event_\"+c.eventName](c.eventData)}}};a.onmousemove=function(b){e&&a.Asc.plugin&&a.Asc.plugin.executeCommand&&a.Asc.plugin.executeCommand(\"onmousemove\",JSON.stringify({x:g===b.clientX?b.pageX:b.clientX,y:g===b.clientY?b.pageY:b.clientY}))};a.onmouseup=function(b){e&&a.Asc.plugin&&a.Asc.plugin.executeCommand&&a.Asc.plugin.executeCommand(\"onmouseup\",JSON.stringify({x:g===b.clientX?b.pageX:b.clientX,y:g===b.clientY?b.pageY:b.clientY}))};a.plugin_sendMessage(JSON.stringify({guid:a.Asc.plugin.guid,type:\"initialize_internal\"}))})(window,void 0);window.startPluginApi=function(){var a=window.Asc.plugin;a.executeCommand=function(a,f,e){window.Asc.plugin.info.type=a;window.Asc.plugin.info.data=f;a=\"\";try{a=JSON.stringify(window.Asc.plugin.info)}catch(b){a=JSON.stringify({type:f})}window.Asc.plugin.onCallCommandCallback=e;window.plugin_sendMessage(a)};a.executeMethod=function(a,f,e){if(!0===window.Asc.plugin.isWaitMethod)return void 0===this.executeMethodStack&&(this.executeMethodStack=[]),this.executeMethodStack.push({name:a,params:f,callback:e}),!1;window.Asc.plugin.isWaitMethod=!0;window.Asc.plugin.methodCallback=e;window.Asc.plugin.info.type=\"method\";window.Asc.plugin.info.methodName=a;window.Asc.plugin.info.data=f;a=\"\";try{a=JSON.stringify(window.Asc.plugin.info)}catch(b){return!1}window.plugin_sendMessage(a);return!0};a.resizeWindow=function(a,f,e,b,d,c){void 0===e&&(e=0);void 0===b&&(b=0);void 0===d&&(d=0);void 0===c&&(c=0);a=JSON.stringify({width:a,height:f,minw:e,minh:b,maxw:d,maxh:c});window.Asc.plugin.info.type=\"resize\";window.Asc.plugin.info.data=a;f=\"\";try{f=JSON.stringify(window.Asc.plugin.info)}catch(k){f=JSON.stringify({type:a})}window.plugin_sendMessage(f)};a.callCommand=function(a,f,e,b){a=\"var Asc = {}; Asc.scope = \"+JSON.stringify(window.Asc.scope)+\"; var scope = Asc.scope; (\"+a.toString()+\")();\";window.Asc.plugin.info.recalculate=!1===e?!1:!0;window.Asc.plugin.executeCommand(!0===f?\"close\":\"command\",a,b)};a.callModule=function(a,f,e){var b=new XMLHttpRequest;b.open(\"GET\",a);b.onreadystatechange=function(){if(4==b.readyState&&(200==b.status||0==location.href.indexOf(\"file:\"))){var a=!0===e?\"close\":\"command\";window.Asc.plugin.info.recalculate=!0;window.Asc.plugin.executeCommand(a,b.responseText);f&&f(b.responseText)}};b.send()};a.loadModule=function(a,f){var e=new XMLHttpRequest;e.open(\"GET\",a);e.onreadystatechange=function(){4!=e.readyState||200!=e.status&&0!=location.href.indexOf(\"file:\")||f&&f(e.responseText)};e.send()};a.createInputHelper=function(){window.Asc.plugin.ih=new window.Asc.inputHelper(window.Asc.plugin)};a.getInputHelper=function(){return window.Asc.plugin.ih}};"/*</code>*/);
			var _iframe = document.getElementById(runObject.frameId);
			if (_iframe)
				_iframe.contentWindow.postMessage(pluginData.serialize(), "*");
			return;
		}
		else if ("reload" == name)
		{
			if (true === pluginData.getAttribute("ctrl"))
			{				
				if (AscCommon.c_oEditorId.Presentation === window.g_asc_plugins.api.getEditorId())
				{
					window.g_asc_plugins.api.sendEvent("asc_onStartDemonstration");
				}
			}
			return;
		}
		else if ("close" == name || "command" == name)
		{
			if (runObject.closeAttackTimer != -1)
			{
				clearTimeout(runObject.closeAttackTimer);
				runObject.closeAttackTimer = -1;
			}

			if (value && value != "")
			{
				var _command_callback_send = ("command" == name);
				try
				{
					if (pluginData.getAttribute("interface"))
					{
						var _script = "(function(){ var Api = window.g_asc_plugins.api;\n" + value + "\n})();";
						eval(_script);
					}
					else if (pluginData.getAttribute("resize") || window.g_asc_plugins.api.asc_canPaste())
					{
						var oLogicDocument, i;
						var editorId = window.g_asc_plugins.api.getEditorId();
						if (AscCommon.c_oEditorId.Word === editorId ||
							AscCommon.c_oEditorId.Presentation === editorId)
						{
							oLogicDocument = window.g_asc_plugins.api.WordControl ?
								window.g_asc_plugins.api.WordControl.m_oLogicDocument : null;
							if(AscCommon.c_oEditorId.Word === editorId){
								oLogicDocument.LockPanelStyles();
							}
						}

                        AscFonts.IsCheckSymbols = true;
						var _script = "(function(){ var Api = window.g_asc_plugins.api;\n" + value + "\n})();";
						try
						{
							eval(_script);
						}
						catch (err)
						{
						}
                        AscFonts.IsCheckSymbols = false;

						if (pluginData.getAttribute("recalculate") == true)
						{
							_command_callback_send = false;
							if (AscCommon.c_oEditorId.Word === editorId ||
								AscCommon.c_oEditorId.Presentation === editorId)
							{
								oLogicDocument = window.g_asc_plugins.api.WordControl ?
									window.g_asc_plugins.api.WordControl.m_oLogicDocument : null;
								var _fonts         = oLogicDocument.Document_Get_AllFontNames();
								var _imagesArray   = oLogicDocument.Get_AllImageUrls();
								var _images        = {};
								for (i = 0; i < _imagesArray.length; i++)
								{
									_images[_imagesArray[i]] = _imagesArray[i];
								}

								window.g_asc_plugins.images_rename = _images;
								AscCommon.Check_LoadingDataBeforePrepaste(window.g_asc_plugins.api, _fonts, _images,
									function()
									{
										if (window.g_asc_plugins.api.WordControl &&
											window.g_asc_plugins.api.WordControl.m_oLogicDocument &&
											window.g_asc_plugins.api.WordControl.m_oLogicDocument.Reassign_ImageUrls)
										{
											window.g_asc_plugins.api.WordControl.m_oLogicDocument.Reassign_ImageUrls(
												window.g_asc_plugins.images_rename);
										}
										delete window.g_asc_plugins.images_rename;

										if(AscCommon.c_oEditorId.Word === editorId) {
											oLogicDocument.UnlockPanelStyles(true);
											oLogicDocument.OnEndLoadScript();
										}

										window.g_asc_plugins.api.asc_Recalculate(true);
										window.g_asc_plugins.api.WordControl.m_oLogicDocument.FinalizeAction();

										if (window.g_asc_plugins.api.SaveAfterMacros)
										{
											window.g_asc_plugins.api.asc_Save();
											window.g_asc_plugins.api.SaveAfterMacros = false;
										}

										var pluginData = new CPluginData();
										pluginData.setAttribute("guid", guid);
										pluginData.setAttribute("type", "onCommandCallback");

										var _iframe = document.getElementById(runObject.frameId);
										if (_iframe)
											_iframe.contentWindow.postMessage(pluginData.serialize(), "*");
									});
							}
							else if (AscCommon.c_oEditorId.Spreadsheet === editorId)
							{
								var oApi    = window.g_asc_plugins.api;
								var oFonts  = oApi.wbModel._generateFontMap();
								var aImages = oApi.wbModel.getAllImageUrls();
								var oImages = {};
								for (i = 0; i < aImages.length; i++)
								{
									oImages[aImages[i]] = aImages[i];
								}
								window.g_asc_plugins.images_rename = oImages;
								AscCommon.Check_LoadingDataBeforePrepaste(window.g_asc_plugins.api, oFonts, oImages,
									function(){
										oApi.wbModel.reassignImageUrls(window.g_asc_plugins.images_rename);
										delete window.g_asc_plugins.images_rename;
										window.g_asc_plugins.api.asc_Recalculate(true);
										var wsView = oApi.wb && oApi.wb.getWorksheet();
										if (wsView && wsView.objectRender && wsView.objectRender.controller) {
											wsView.objectRender.controller.recalculate2(undefined);
										}

										if (window.g_asc_plugins.api.SaveAfterMacros)
										{
											window.g_asc_plugins.api.asc_Save();
											window.g_asc_plugins.api.SaveAfterMacros = false;
										}

										var pluginData = new CPluginData();
										pluginData.setAttribute("guid", guid);
										pluginData.setAttribute("type", "onCommandCallback");

										var _iframe = document.getElementById(runObject.frameId);
										if (_iframe)
											_iframe.contentWindow.postMessage(pluginData.serialize(), "*");
									});
							}
						}
						else
						{
							if (AscCommon.c_oEditorId.Spreadsheet === editorId)
							{
								// На asc_canPaste создается точка в истории и startTransaction. Поэтому нужно ее закрыть без пересчета.
								window.g_asc_plugins.api.asc_endPaste();
							}
							else if (AscCommon.c_oEditorId.Word === editorId ||
								AscCommon.c_oEditorId.Presentation === editorId)
							{
								window.g_asc_plugins.api.WordControl.m_oLogicDocument.FinalizeAction();
							}
						}
					}
				} catch (err)
				{
				}

				if (_command_callback_send)
				{
					var pluginData = new CPluginData();
					pluginData.setAttribute("guid", guid);
					pluginData.setAttribute("type", "onCommandCallback");
					var _iframe = document.getElementById(runObject.frameId);
					if (_iframe)
						_iframe.contentWindow.postMessage(pluginData.serialize(), "*");
				}
			}

			if ("close" == name)
			{
				window.g_asc_plugins.close(guid);
			}
		}
		else if ("resize" == name)
		{
			var _sizes = JSON.parse(value);

			window.g_asc_plugins.api.sendEvent("asc_onPluginResize",
				[_sizes["width"], _sizes["height"]],
				[_sizes["minw"], _sizes["minh"]],
				[_sizes["maxw"], _sizes["maxh"]], function() {
				// TODO: send resize end event
			});
		}
		else if ("onmousemove" == name)
		{
			var _pos = JSON.parse(value);
			window.g_asc_plugins.api.sendEvent("asc_onPluginMouseMove", _pos["x"], _pos["y"]);
		}
		else if ("onmouseup" == name)
		{
			var _pos = JSON.parse(value);
			window.g_asc_plugins.api.sendEvent("asc_onPluginMouseUp", _pos["x"], _pos["y"]);
		}
		else if ("method" == name)
		{
			var _apiMethodName = "pluginMethod_" + pluginData.getAttribute("methodName");
			var _return = undefined;

			window.g_asc_plugins.guidAsyncMethod = guid;

			if (window.g_asc_plugins.api[_apiMethodName])
				_return = window.g_asc_plugins.api[_apiMethodName].apply(window.g_asc_plugins.api, value);

			if (!runObject.methodReturnAsync)
			{
				var pluginData = new CPluginData();
				pluginData.setAttribute("guid", guid);
				pluginData.setAttribute("type", "onMethodReturn");
				pluginData.setAttribute("methodReturnData", _return);
				var _iframe = document.getElementById(runObject.frameId);
				if (_iframe)
					_iframe.contentWindow.postMessage(pluginData.serialize(), "*");
			}
			runObject.methodReturnAsync = false;
			window.g_asc_plugins.guidAsyncMethod = "";
			return;
		}
	}

	if (window.addEventListener)
	{
		window.addEventListener("message", onMessage, false);
	}
	else if (window.attachEvent)
	{
		window.attachEvent("onmessage", onMessage);
	}

	window["Asc"]                      = window["Asc"] ? window["Asc"] : {};
	window["Asc"].createPluginsManager = function(api)
	{
		if (window.g_asc_plugins)
			return window.g_asc_plugins;

		window.g_asc_plugins        = new CPluginsManager(api);
		window["g_asc_plugins"]     = window.g_asc_plugins;
		window.g_asc_plugins.api    = api;
		window.g_asc_plugins["api"] = window.g_asc_plugins.api;

		api.asc_registerCallback('asc_onSelectionEnd', function(){
			window.g_asc_plugins.onChangedSelectionData();
		});

		window.g_asc_plugins.api.asc_registerCallback('asc_onDocumentContentReady', function()
		{

			setTimeout(function()
			{
				window.g_asc_plugins.loadExtensionPlugins(window["Asc"]["extensionPlugins"]);
			}, 10);

		});
		window.g_asc_plugins.api.asc_registerCallback('asc_LoadPluginsOrDocument', function()
		{
			window.g_asc_plugins.countEventDocContOrPluginsReady++;
		});

        if (window.location && window.location.search)
        {
            var _langSearch = window.location.search;
            var _pos1 = _langSearch.indexOf("lang=");
            var _pos2 = (-1 != _pos1) ? _langSearch.indexOf("&", _pos1) : -1;
            if (_pos1 >= 0)
            {
                _pos1 += 5;

                if (_pos2 < 0)
                    _pos2 = _langSearch.length;

                var _lang = _langSearch.substr(_pos1, _pos2 - _pos1);
                if (_lang.length == 2)
                {
                    _lang = (_lang.toLowerCase() + "-" + _lang.toUpperCase());
                }

                if (5 == _lang.length)
                    window.g_asc_plugins.language = _lang;
            }
        }

		if (window["AscDesktopEditor"] && window["UpdateSystemPlugins"])
			window["UpdateSystemPlugins"]();

		return window.g_asc_plugins;
	};

	window["Asc"].CPluginData      = CPluginData;
	window["Asc"].CPluginData_wrap = function(obj)
	{
		if (!obj.getAttribute)
			obj.getAttribute = function(name)
			{
				return this[name];
			};
		if (!obj.setAttribute)
			obj.setAttribute = function(name, value)
			{
				return this[name] = value;
			};
	};

    window["Asc"].loadConfigAsInterface = function(url)
	{
        if (url)
        {
            try {
                var xhrObj = new XMLHttpRequest();
                if ( xhrObj )
                {
                    xhrObj.open('GET', url, false);
                    xhrObj.send('');

                    return JSON.parse(xhrObj.responseText);
                }
            } catch (e) {}
        }
        return null;
	};

	window["Asc"].loadPluginsAsInterface = function(api)
	{
		if (window.g_asc_plugins.srcPluginsLoaded)
			return;
        window.g_asc_plugins.srcPluginsLoaded = true;

		var configs = window["Asc"].loadConfigAsInterface("../../../../plugins.json");

        if (!configs)
        	return;

        var pluginsData = configs["pluginsData"];
        if (!pluginsData || pluginsData.length < 1)
        	return;

        var arrPluginsConfigs = [];
        pluginsData.forEach(function(item) {
            var value = window["Asc"].loadConfigAsInterface(item);
            if (value) {
                value["baseUrl"] = item.substring(0, item.lastIndexOf("config.json"));
                arrPluginsConfigs.push(value);
            }
        });

        var arrPlugins = [];
        arrPluginsConfigs.forEach(function(item) {
            var plugin = new Asc.CPlugin();
            plugin["set_Name"](item["name"]);
            plugin["set_Guid"](item["guid"]);
            plugin["set_BaseUrl"](item["baseUrl"]);
            var variations = item["variations"];
        	var variationsArr = [];
            variations.forEach(function(itemVar){
                var variation = new Asc.CPluginVariation();
                variation["set_Description"](itemVar["description"]);
                variation["set_Url"](itemVar["url"]);
                variation["set_Icons"](itemVar["icons"]);
                variation["set_Visual"](itemVar["isVisual"]);
                variation["set_CustomWindow"](itemVar["'isCustomWindow"]);
                variation["set_System"](itemVar["isSystem"]);
                variation["set_Viewer"](itemVar["isViewer"]);
                variation["set_EditorsSupport"](itemVar["EditorsSupport"]);
                variation["set_Modal"](itemVar["isModal"]);
                variation["set_InsideMode"](itemVar["isInsideMode"]);
                variation["set_InitDataType"](itemVar["initDataType"]);
                variation["set_InitData"](itemVar["initData"]);
                variation["set_UpdateOleOnResize"](itemVar["isUpdateOleOnResize"]);
                variation["set_Buttons"](itemVar["buttons"]);
                variation["set_Size"](itemVar["size"]);
                variation["set_InitOnSelectionChanged"](itemVar["initOnSelectionChanged"]);
                variation["set_Events"](itemVar["events"]);
                variationsArr.push(variation);
            });
            plugin["set_Variations"](variationsArr);
            arrPlugins.push(plugin);
        });

        window.g_asc_plugins.srcPlugins = arrPluginsConfigs;
        api.asc_pluginsRegister('', arrPlugins);

        api.asc_registerCallback('asc_onPluginShow', function(plugin, variationIndex, frameId) {

        	var _t = window.g_asc_plugins;

        	var srcPlugin = null;
        	for (var i = 0; i < _t.srcPlugins.length; i++)
			{
				if (plugin.guid == _t.srcPlugins[i]["guid"])
				{
					srcPlugin = _t.srcPlugins[i];
					break;
				}
            }

        	var variation = plugin.get_Variations()[variationIndex];

            var _elem = document.createElement("div");
            _elem.id = "parent_" + frameId;
            _elem.setAttribute("style", "user-select:none;z-index:5000;position:fixed;left:10px;top:10px;right:10px;bottom:10px;box-sizing:border-box;z-index:5000;box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);border-radius: 5px;background-color: #fff;border: solid 1px #cbcbcb;");

            var _elemBody = "";
            _elemBody += "<div style=\"box-sizing:border-box;height: 34px;padding: 5px 6px 6px;left:0;right:0;top:0;border-bottom: solid 1px #cbcbcb;background: #ededed;text-align: center;vertical-align: bottom;\">";
            _elemBody += "<span style=\"color: #848484;text-align: center;font-size: 12px;font-weight:700;text-shadow: 1px 1px #f8f8f8;line-height:26px;vertical-align: bottom;font-family:'Helvetica Neue', Helvetica, Arial, sans-serif;\">";

            var lang = _t.language;
            var lang2 = _t.language.substr(0, 2);

            var _name = plugin.name;
            if (srcPlugin && srcPlugin["nameLocale"])
			{
				if (srcPlugin["nameLocale"][lang])
					_name = srcPlugin["nameLocale"][lang];
				else if (srcPlugin["nameLocale"][lang2])
                    _name = srcPlugin["nameLocale"][lang2];
			}

            _elemBody += _name;
            _elemBody += "</span></div>";

            _elemBody += "<div style=\"position:absolute;box-sizing:border-box;height:calc(100% - 86px);padding: 0;left:0;right:0;top:34px;background:#FFFFFF;\">";

            var _add = plugin.baseUrl == "" ? _t.path : plugin.baseUrl;
            _elemBody += ("<iframe name=\"" + frameId + "\" id=\"" + frameId + "\" src=\"" + (_add + variation.url) + "\" ");
            _elemBody += "style=\"position:absolute;left:0; top:0px; right: 0; bottom: 0; width:100%; height:100%; overflow: hidden;\" frameBorder=\"0\">";
            _elemBody += "</iframe>";

            _elemBody += "</div>";

            _elemBody += "<div style=\"position:absolute;box-sizing:border-box;height:52px;padding: 15px 15px 15px 15px;left:0;right:0;top:calc(100% - 52px);bottom:0;border-top: solid 1px #cbcbcb;background: #ededed;text-align: center;vertical-align: bottom;\">";

            var buttons = variation["get_Buttons"]();

            for (var i = 0; i < buttons.length; i++)
			{
            	_elemBody += ("<button id=\"plugin_button_id_" + i + "\" style=\"border-radius:1px;margin-right:10px;height:22px;font-weight:bold;background-color:#d8dadc;color:#444444;touch-action: manipulation;border: 1px solid transparent;text-align:center;vertical-align: middle;outline:none;font-family:'Helvetica Neue', Helvetica, Arial, sans-serif;font-size: 12px;\">");

                _name = buttons[i]["text"];
                if (srcPlugin && srcPlugin["variations"][variationIndex]["buttons"][i]["textLocale"])
                {
                    if (srcPlugin["variations"][variationIndex]["buttons"][i]["textLocale"][lang])
                        _name = srcPlugin["variations"][variationIndex]["buttons"][i]["textLocale"][lang];
                    else if (srcPlugin["variations"][variationIndex]["buttons"][i]["textLocale"][lang2])
                        _name = srcPlugin["variations"][variationIndex]["buttons"][i]["textLocale"][lang2];
                }

            	_elemBody += _name;

            	_elemBody += "</button>";
			}

			if (0 == buttons.length)
			{
                _elemBody += ("<button id=\"plugin_button_id_0\" style=\"border-radius:1px;margin-right:10px;height:22px;font-weight:bold;background-color:#d8dadc;color:#444444;touch-action: manipulation;border: 1px solid transparent;text-align:center;vertical-align: middle;outline:none;font-family:'Helvetica Neue', Helvetica, Arial, sans-serif;font-size: 12px;\">");
                _elemBody += "Ok</button>";
			}

            _elemBody += "</div>";

            _elem.innerHTML = _elemBody;

            document.body.appendChild(_elem);

            for (var i = 0; i < buttons.length; i++)
            {
            	var _button = document.getElementById("plugin_button_id_" + i);
            	if (_button)
				{
					_button.onclick = function()
					{
						var nId = this.id.substr("plugin_button_id_".length);
                        window.g_asc_plugins.api.asc_pluginButtonClick(parseInt(nId));
					}
				}
            }

            if (0 == buttons.length)
			{
                var _button = document.getElementById("plugin_button_id_0");
                if (_button)
                {
                    _button.onclick = function()
                    {
                       	window.g_asc_plugins.api.asc_pluginButtonClick(-1);
                    }
                }
			}
        });

        api.asc_registerCallback('asc_onPluginClose', function(plugin, variationIndex) {

        	var _elem = document.getElementById("parent_iframe_" + plugin.guid);
        	if (_elem)
        		document.body.removeChild(_elem);
            _elem = null;
        });
	}
})(window, undefined);
