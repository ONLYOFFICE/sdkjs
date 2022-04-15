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

	function CIHelper(plugin)
	{
		this.plugin = plugin;
		this.ps;
		this.items = [];
		this.isVisible = false;
		this.isCurrentVisible = false;
	};

	CIHelper.prototype.createWindow = function()
	{
		var _body = document.body;
		var _head = document.getElementsByTagName('head')[0];
		if (!_body || !_head)
			return;

		var _style = document.createElement('style');
		_style.type      = 'text/css';

		var _style_body = ".ih_main { margin: 0px; padding: 0px; width: 100%; height: 100%; display: inline-block; overflow: hidden; box-sizing: border-box; user-select: none; position: fixed; border: 1px solid #cfcfcf; } ";
		_style_body += "ul { margin: 0px; padding: 0px; width: 100%; height: 100%; list-style-type: none; outline:none; } ";
		_style_body += "li { padding: 5px; font-family: \"Helvetica Neue\", Helvetica, Arial, sans-serif; font-size: 12px; font-weight: 400; color: #373737; } ";
		_style_body += "li:hover { background-color: #D8DADC; } ";
		_style_body += ".li_selected { background-color: #D8DADC; color: #373737; }";
		_style_body += ".li_selected:hover { background-color: #D8DADC; color: #373737; }";
		_style.innerHTML = _style_body;
		_head.appendChild(_style);

		document.body.style.background = "#FFFFFF";
		document.body.style.width = "100%";
		document.body.style.height = "100%";
		document.body.style.margin = "0";
		document.body.style.padding = "0";

		document.body.innerHTML = "<div class=\"ih_main\" id=\"ih_area\"><ul id=\"ih_elements_id\" role=\"listbox\"></ul></div>";

		this.ps = new PerfectScrollbar(document.getElementById("ih_area"), { minScrollbarLength: 20 });
		this.updateScrolls();

		this.createDefaultEvents();
	};

	CIHelper.prototype.setItems = function(items)
	{
		this.items = items;

		var _data = "";
		var _len = items.length;
		for (var i = 0; i < _len; i++)
		{
			if (undefined === items[i].id)
				items[i].id = "" + i;

			_data += "<li role=\"option\"";
			if (0 == i)
				_data += " class=\"li_selected\"";

			_data += " id=\"" + items[i].id + "\"";

			_data += " onclick=\"_private_on_ih_click(event)\">";
			_data += items[i].text;
			_data += "</li>";
		}

		document.getElementById("ih_elements_id").innerHTML = _data;
		this.updateScrolls();
		this.scrollToSelected();
	};

	CIHelper.prototype.createDefaultEvents = function()
	{
		this.plugin.onExternalMouseUp = function()
		{
			var evt = document.createEvent("MouseEvents");
			evt.initMouseEvent("mouseup", true, true, window, 1, 0, 0, 0, 0,
				false, false, false, false, 0, null);

			document.dispatchEvent(evt);
		};

		var _t = this;
		window.onkeydown = function(e) {
			switch (e.keyCode)
			{
				case 27: // Escape
				{
					if (_t.isVisible)
					{
						_t.isVisible = false;
						_t.plugin.executeMethod("UnShowInputHelper", [_t.plugin.info.guid, true]);
					}
					break;
				}
				case 38: // Up
				case 40: // Down
				case 9: // Tab
				case 36: // Home
				case 35: // End
				case 33: // PageUp
				case 34: // PageDown
				{
					var items = document.getElementsByTagName("li");
					var curIndex = -1;
					for (var i = 0; i < items.length; i++)
					{
						if (items[i].className == "li_selected")
						{
							curIndex = i;
							items[i].className = "";
							break;
						}
					}
					if (curIndex == -1)
					{
						curIndex = 0;
					}
					else
					{
						switch (e.keyCode)
						{
							case 38:
							{
								curIndex--;
								if (curIndex < 0)
									curIndex = 0;
								break;
							}
							case 40:
							{
								curIndex++;
								if (curIndex >= items.length)
									curIndex = items.length - 1;
								break;
							}
							case 9:
							{
								curIndex++;
								if (curIndex >= items.length)
									curIndex = 0;
								break;
							}
							case 36:
							{
								curIndex = 0;
								break;
							}
							case 35:
							{
								curIndex = items.length - 1;
								break;
							}
							case 33:
							case 34:
							{
								var _indexDif = 1;
								var _count = (document.getElementById("ih_area").clientHeight / 24) >> 0;
								if (_count > 1)
									_indexDif = _count;

								if (33 == e.keyCode)
								{
									curIndex -= _indexDif;
									if (curIndex < 0)
										curIndex = 0;
								}
								else
								{
									curIndex += _indexDif;
									if (curIndex >= items.length)
										curIndex = curIndex = items.length - 1;;
								}
								break;
							}
						}
					}

					if (curIndex < items.length)
					{
						items[curIndex].className = "li_selected";

						var _currentOffset = items[curIndex].offsetTop;
						var _currentHeight = items[curIndex].offsetHeight;

						var container = document.getElementById("ih_area");
						var _currentScroll = container.scrollTop;
						if (_currentOffset < _currentScroll)
						{
							if (container.scrollTo)
								container.scrollTo(0, _currentOffset);
							else
								container.scrollTop = _currentOffset;
						}
						else if ((_currentScroll + container.offsetHeight) < (_currentOffset + _currentHeight))
						{
							if (container.scrollTo)
								container.scrollTo(0, _currentOffset - (container.offsetHeight - _currentHeight));
							else
								container.scrollTop = _currentOffset - (container.offsetHeight - _currentHeight);
						}
					}
					break;
				}
				case 13:
				{
					_t.onSelectedItem();
					break;
				}
			}

			if (e.preventDefault)
				e.preventDefault();
			if (e.stopPropagation)
				e.stopPropagation();
			return false;
		};

		window.onresize = function(e)
		{
			_t.updateScrolls();
		};

		window._private_on_ih_click = function(e)
		{
			var items = document.getElementsByTagName("li");
			for (var i = 0; i < items.length; i++)
			{
				items[i].className = "";
			}
			e.target.className = "li_selected";
			var _id = e.target.getAttribute("id");
			_t.onSelectedItem();
		};

		this.plugin.event_onKeyDown = function(data)
		{
			window.onkeydown({ keyCode : data.keyCode });
		};
	};

	CIHelper.prototype.updateScrolls = function()
	{
		this.ps.update(); this.ps.update();

		var _elemV = document.getElementsByClassName("ps__rail-y")[0];
		var _elemH = document.getElementsByClassName("ps__rail-x")[0];

		if (!_elemH || !_elemV)
			return;

		var _styleV = window.getComputedStyle(_elemV);
		var _styleH = window.getComputedStyle(_elemH);

		var _visibleV = (_styleV && _styleV.display == "none") ? false : true;
		var _visibleH = (_styleH && _styleH.display == "none") ? false : true;

		if (_visibleH && _visibleV)
		{
			if ("13px" != _elemV.style.marginBottom)
				_elemV.style.marginBottom = "13px";
			if ("13px" != _elemH.style.marginRight)
				_elemH.style.marginRight = "13px";
		}
		else
		{
			if ("2px" != _elemV.style.marginBottom)
				_elemV.style.marginBottom = "2px";
			if ("2px" != _elemH.style.marginRight)
				_elemH.style.marginRight = "2px";
		}
	};

	CIHelper.prototype.scrollToSelected = function()
	{
		var items = document.getElementsByTagName("li");
		var curIndex = -1;
		for (var i = 0; i < items.length; i++)
		{
			if (items[i].className == "li_selected")
			{
				var container = document.getElementById("ih_area");
				if (container.scrollTo)
					container.scrollTo(0, items[i].offsetTop);
				else
					container.scrollTop = items[i].offsetTop;
				return;
			}
		}
	};

	CIHelper.prototype.getSelectedItem = function()
	{
		var items = document.getElementsByTagName("li");
		var curId = -1;
		for (var i = 0; i < items.length; i++)
		{
			if (items[i].className == "li_selected")
			{
				curId = items[i].getAttribute("id");
				break;
			}
		}

		if (-1 == curId)
			return null;

		var len = this.items.length;
		for (var i = 0; i < len; i++)
		{
			if (curId == this.items[i].id)
				return this.items[i];
		}

		return null;
	};

	CIHelper.prototype.onSelectedItem = function()
	{
		if (this.plugin.inputHelper_onSelectItem)
			this.plugin.inputHelper_onSelectItem(this.getSelectedItem());
	};

	CIHelper.prototype.show = function(w, h, isKeyboardTake)
	{
		this.isCurrentVisible = true;
		this.plugin.executeMethod("ShowInputHelper", [this.plugin.info.guid, w, h, isKeyboardTake], function() { window.Asc.plugin.ih.isVisible = true; });
	};

	CIHelper.prototype.unShow = function()
	{
		if (!this.isCurrentVisible && !this.isVisible)
			return;

		this.isCurrentVisible = false;
		window.Asc.plugin.executeMethod("UnShowInputHelper", [this.plugin.info.guid], function() { window.Asc.plugin.ih.isVisible = false; });
	};

	CIHelper.prototype.getItemHeight = function()
	{
		var _sizeItem = 24;
		var _items = document.getElementsByTagName("li");
		if (_items.length > 0 && _items[0].offsetHeight > 0)
			_sizeItem = _items[0].offsetHeight;
		return _sizeItem;
	};

	CIHelper.prototype.getItemsHeight = function(count)
	{
		return 2 + count * this.getItemHeight();
	};

	CIHelper.prototype.getItems = function()
	{
		return this.items;
	};

	CIHelper.prototype.getScrollSizes = function()
	{
		var _size = { w : 0, h : 0 };
		var _sizeItem = this.getItemHeight();

		var _elem = document.getElementById("ih_elements_id");
		if (_elem)
		{
			_size.w = _elem.scrollWidth;
			_size.h = 2 + this.items.length * _sizeItem;
		}
		return _size;
	};

	window.Asc = window.Asc || {};
	window.Asc.inputHelper = CIHelper;

})(window, undefined);

(function(window, undefined){

	function onReloadPage(isCtrl)
	{
		window.parent.postMessage(JSON.stringify({
			type : "reload",
			guid : window.Asc.plugin.guid,
			ctrl : isCtrl
		}), "*");
	}
	function onBaseKeyDown(e)
	{
		var isCtrl = (e.metaKey || e.ctrlKey) ? true : false;
		if (e.keyCode == 116)
		{
			onReloadPage(isCtrl);
			if (e.preventDefault)
				e.preventDefault();
			if (e.stopPropagation)
				e.stopPropagation();
			return false;
		}
	}

	if (window.addEventListener)
		window.addEventListener("keydown", onBaseKeyDown, false);
	else
		window.attachEvent("keydown", onBaseKeyDown);

})(window, undefined);

(function(window, undefined){

	// className => { css property => key in theme object }
	var g_themes_map = {
		"body" : { "color" : "text-normal", "background-color" : "background-toolbar" },
		".defaultlable" : { "color" : "text-normal" },
		".aboutlable" : { "color" : "text-normal" },
		"a.aboutlink" : { "color" : "text-normal" },
		".form-control, .form-control[readonly], .form-control[disabled]" : { "color" : "text-normal", "background-color" : "background-normal", "border-color" : "border-regular-control" },
		".form-control:focus" : { "border-color" : "border-control-focus" },
		".form-control[disabled]" : { "color" : "text-invers" },
		".btn-text-default" : { "background-color" : "background-normal", "border-color" : "border-regular-control", "color" : "text-normal" },
		".btn-text-default:hover" : { "background-color" : "highlight-button-hover" },
		".btn-text-default:active" : { "background-color" : "highlight-button-pressed !important" },
		".btn-text-default[disabled]:hover,\
		.btn-text-default.disabled:hover,\
		.btn-text-default[disabled]:active,\
		.btn-text-default[disabled].active,\
		.btn-text-default.disabled:active,\
		.btn-text-default.disabled.active": {"background-color" : "background-normal !important", "color" : "text-normal"},
		".select2-container--default .select2-selection--single" : { "color" : "text-normal", "background-color" : "background-normal" },
		".select2-container--default .select2-selection--single .select2-selection__rendered" : { "color" : "text-normal" },
		".select2-results" : { "background-color" : "background-normal" },
		".select2-container--default .select2-results__option--highlighted[aria-selected]" : { "background-color" : "border-regular-control !important"},
		".select2-container--default .select2-results__option[aria-selected=true]" : { "background-color" : "border-regular-control !important"},
		".select2-dropdown, .select2-container--default .select2-selection--single" : { "border-color" : "border-regular-control !important"},
		".select2-container--default.select2-container--open .select2-selection--single" : { "border-color" : "border-control-focus !important"},
		".select2-container--default.select2-container--focus:not(.select2-container--open) .select2-selection--single" : { "border-color" : "border-regular-control !important"},
		".select2-container--default.select2-container--open.select2-container--focus .select2-selection--single" : { "border-color" : "border-control-focus !important"},
		".select2-search--dropdown" : { "background-color" : "background-normal !important"},
		".select2-container--default .select2-search--dropdown .select2-search__field" : { "color" : "text-normal", "background-color" : "background-normal", "border-color" : "border-regular-control"},
		".select2-container--default.select2-container--disabled .select2-selection--single" : { "background-color" : "background-normal" },
		".select2-container--default .select2-selection--single .select2-selection__arrow b" : { "border-color" : "text-normal !important" },
		".select2-container--default.select2-container--open .select2-selection__arrow b" : {"border-color" : "text-normal !important"},
		".ps .ps__rail-y:hover" : {"background-color" : "background-toolbar" },
		".ps .ps__rail-y.ps--clicking" : {"background-color" : "background-toolbar" },
		".ps__thumb-y" : { "background-color" : "background-normal", "border-color" : "Border !important" },
		".ps__rail-y:hover > .ps__thumb-y" : {"border-color" : "canvas-scroll-thumb-hover", "background-color" : "canvas-scroll-thumb-hover !important" },
		".ps .ps__rail-x:hover" : {"background-color" : "background-toolbar" },
		".ps .ps__rail-x.ps--clicking" : {"background-color" : "background-toolbar" },
		".ps__thumb-x" : { "background-color" : "background-normal", "border-color" : "Border !important" },
		".ps__rail-x:hover > .ps__thumb-x" : {"border-color" : "canvas-scroll-thumb-hover" },
		"a" : { "color" : "text-link !important" },
		"a:hover" : { "color" : "text-link-hover !important" },
		"a:active" : { "color" : "text-link-active !important" },
		"a:visited" : { "color" : "text-link-visited !important" },
		"*::-webkit-scrollbar-track" : { "background" : "background-normal" },
		"*::-webkit-scrollbar-track:hover" : { "background" : "background-toolbar-additional" },
		"*::-webkit-scrollbar-thumb" : { "background-color" : "background-toolbar", "border-color" : "border-regular-control" },
		"*::-webkit-scrollbar-thumb:hover" : { "background-color" : "canvas-scroll-thumb-hover" },
		".asc-plugin-loader" : { "color" : "text-normal" }
	};

    var g_isMouseSendEnabled = false;
    var g_language = "";

    window.plugin_sendMessage = function sendMessage(data)
    {
        if (window.Asc.plugin.ie_channel)
            window.Asc.plugin.ie_channel.postMessage(data);
        else
            window.parent.postMessage(data, "*");
    };

    window.plugin_onMessage = function(event)
    {
        if (!window.Asc.plugin)
            return;

        if (typeof(event.data) == "string")
        {
            var pluginData = {};
            try
            {
                pluginData = JSON.parse(event.data);
            }
            catch(err)
            {
                pluginData = {};
            }

            var type = pluginData.type;

            if (pluginData.guid != window.Asc.plugin.guid)
            {
                if (undefined !== pluginData.guid)
                    return;

                switch (type)
                {
                    case "onExternalPluginMessage":
                        break;
                    default:
                        return;
                }
            }

            if (type == "init")
                window.Asc.plugin.info = pluginData;

            if (undefined !== pluginData.theme)
			{
				if (!window.Asc.plugin.theme || type === "onThemeChanged")
				{
					window.Asc.plugin.theme = pluginData.theme;

					if (!window.Asc.plugin.onThemeChangedBase)
					{
						window.Asc.plugin.onThemeChangedBase = function (newTheme)
						{
							// correct theme
							var rules = "";
							for (var className in g_themes_map)
							{
								rules += (className + " {");

								var attributes = g_themes_map[className];
								for (var attr in attributes)
								{
									var attrValue = attributes[attr];
									var attrValueImportant = attrValue.indexOf(" !important");
									if (-1 < attrValueImportant)
										attrValue = attrValue.substr(0, attrValueImportant);
									var newVal = newTheme[attrValue];
									if (newVal)
										rules += (attr + " : " + newVal + ((-1 === attrValueImportant) ? ";" : " !important;"));
								}

								rules += " }\n";
							}

							var styleTheme = document.createElement('style');
							styleTheme.type = 'text/css';
							styleTheme.innerHTML = rules;
							document.getElementsByTagName('head')[0].appendChild(styleTheme);
						};
					}

					if (window.Asc.plugin.onThemeChanged)
						window.Asc.plugin.onThemeChanged(window.Asc.plugin.theme);
					else
						window.Asc.plugin.onThemeChangedBase(window.Asc.plugin.theme);
				}
			}

            if (!window.Asc.plugin.tr || !window.Asc.plugin.tr_init)
            {
				window.Asc.plugin.tr_init = true;
                window.Asc.plugin.tr = function(val) {
                    if (!window.Asc.plugin.translateManager || !window.Asc.plugin.translateManager[val])
                        return val;
                    return window.Asc.plugin.translateManager[val];
                };
            }

            var newLang = "";
            if (window.Asc.plugin.info)
                newLang = window.Asc.plugin.info.lang;
            if (newLang == "" || newLang != g_language)
            {
                g_language = newLang;
                if (g_language == "en-EN" || g_language == "")
				{
					window.Asc.plugin.translateManager = {};
					if (window.Asc.plugin.onTranslate)
						window.Asc.plugin.onTranslate();
				}
				else
				{
					var _client = new XMLHttpRequest();
					_client.open("GET", "./translations/" + g_language + ".json");

					_client.onreadystatechange = function ()
					{
						if (_client.readyState == 4)
						{
						    if (_client.status == 200 || location.href.indexOf("file:") == 0)
						    {
                                try
                                {
                                    window.Asc.plugin.translateManager = JSON.parse(_client.responseText);
                                }
                                catch (err)
                                {
                                    window.Asc.plugin.translateManager = {};
                                }
                            }
                            else
                            {
                                window.Asc.plugin.translateManager = {};
                            }


                            if (window.Asc.plugin.onTranslate)
                                window.Asc.plugin.onTranslate();
						}
					};
					_client.send();
				}
            }

            switch (type)
            {
                case "init":
                {
                    pluginStart();
                    window.Asc.plugin.init(window.Asc.plugin.info.data);
                    break;
                }
                case "button":
                {
                    var _buttonId = parseInt(pluginData.button);
                    if (!window.Asc.plugin.button && -1 == _buttonId)
						window.Asc.plugin.executeCommand("close", "");
                    else
                        window.Asc.plugin.button(_buttonId);
                    break;
                }
                case "enableMouseEvent":
                {
                    g_isMouseSendEnabled = pluginData.isEnabled;
					if (window.Asc.plugin.onEnableMouseEvent)
						window.Asc.plugin.onEnableMouseEvent(g_isMouseSendEnabled);
                    break;
                }
                case "onExternalMouseUp":
                {
                    if (window.Asc.plugin.onExternalMouseUp)
                        window.Asc.plugin.onExternalMouseUp();
                    break;
                }
                case "onMethodReturn":
                {
					window.Asc.plugin.isWaitMethod = false;

					if (window.Asc.plugin.methodCallback)
					{
					    var methodCallback = window.Asc.plugin.methodCallback;
                        window.Asc.plugin.methodCallback = null;
                        methodCallback(pluginData.methodReturnData);
                        methodCallback = null;
					}
					else if (window.Asc.plugin.onMethodReturn)
					{
						window.Asc.plugin.onMethodReturn(pluginData.methodReturnData);
					}

					if (window.Asc.plugin.executeMethodStack && window.Asc.plugin.executeMethodStack.length > 0)
                    {
                        var obj = window.Asc.plugin.executeMethodStack.shift();
                        window.Asc.plugin.executeMethod(obj.name, obj.params, obj.callback);
                    }

                    break;
                }
				case "onCommandCallback":
				{
                    if (window.Asc.plugin.onCallCommandCallback)
                    {
                        window.Asc.plugin.onCallCommandCallback();
                        window.Asc.plugin.onCallCommandCallback = null;
                    }
                    else if (window.Asc.plugin.onCommandCallback)
						window.Asc.plugin.onCommandCallback();
					break;
				}
                case "onExternalPluginMessage":
                {
					if (window.Asc.plugin.onExternalPluginMessage && pluginData.data && pluginData.data.type)
						window.Asc.plugin.onExternalPluginMessage(pluginData.data);
					break;
                }
                case "onEvent":
                {
                    if (window.Asc.plugin["event_" + pluginData.eventName])
                        window.Asc.plugin["event_" + pluginData.eventName](pluginData.eventData);
                    break;
                }
                default:
                    break;
            }
        }
    };

    function pluginStart()
    {
        if (window.Asc.plugin.isStarted)
            return;

        window.Asc.plugin.isStarted = true;
        window.startPluginApi();

		var zoomValue = AscCommon.checkDeviceScale();
		AscCommon.retinaPixelRatio = zoomValue.applicationPixelRatio;
		AscCommon.zoom = zoomValue.zoom;
		AscCommon.correctApplicationScale(zoomValue);

		window.Asc.plugin.onEnableMouseEvent = function(isEnabled)
		{
			var _frames = document.getElementsByTagName("iframe");
			if (_frames && _frames[0])
			{
				_frames[0].style.pointerEvents = isEnabled ? "none" : "";
			}
		};
    }

    window.onmousemove = function(e)
    {
        if (!g_isMouseSendEnabled || !window.Asc.plugin || !window.Asc.plugin.executeCommand)
            return;

        var _x = (undefined === e.clientX) ? e.pageX : e.clientX;
        var _y = (undefined === e.clientY) ? e.pageY : e.clientY;

        window.Asc.plugin.executeCommand("onmousemove", JSON.stringify({ x : _x, y : _y }));

    };
    window.onmouseup   = function(e)
    {
        if (!g_isMouseSendEnabled || !window.Asc.plugin || !window.Asc.plugin.executeCommand)
            return;

        var _x = (undefined === e.clientX) ? e.pageX : e.clientX;
        var _y = (undefined === e.clientY) ? e.pageY : e.clientY;

        window.Asc.plugin.executeCommand("onmouseup", JSON.stringify({ x : _x, y : _y }));
    };

    window.plugin_sendMessage(JSON.stringify({ guid : window.Asc.plugin.guid, type : "initialize_internal" }));

})(window, undefined);
