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

(function (window, undefined)
{
	window['AscFonts'] = window['AscFonts'] || {};

	window['AscFonts'].isEngineReady = false;
	window['AscFonts'].api = null;
	window['AscFonts'].onSuccess = null;
	window['AscFonts'].onError = null;
	window['AscFonts'].maxLoadingIndex = 2; // engine (1+1)
	window['AscFonts'].curLoadingIndex = 0;

	window['AscFonts'].allocate = function(size)
	{
		if (typeof(Uint8Array) != 'undefined' && !window.opera)
			return new Uint8Array(size);

		var arr = new Array(size);
		for (var i=0;i<size;i++)
			arr[i] = 0;
		return arr;
	};
	window['AscFonts'].allocateData = function(size)
	{
		return { data : window['AscFonts'].allocate(size) };
	};

	window['AscFonts']['onLoadModule'] = function()
	{
		++window['AscFonts'].curLoadingIndex;

		if (window['AscFonts'].curLoadingIndex === window['AscFonts'].maxLoadingIndex)
		{
			onLoadFontsModule(window, undefined);

			window['AscFonts'].isEngineReady = true;
			window['AscFonts'].onSuccess && window['AscFonts'].onSuccess.call(window['AscFonts'].api);

			delete window['AscFonts'].curLoadingIndex;
			delete window['AscFonts'].maxLoadingIndex;
			delete window['AscFonts'].api;
			delete window['AscFonts'].onSuccess;
			delete window['AscFonts'].onError;
		}
	};

	window['AscFonts'].load = function(api, onSuccess, onError)
	{
		window['AscFonts'].api = api;
		window['AscFonts'].onSuccess = onSuccess;
		window['AscFonts'].onError = onError;

		if (window["NATIVE_EDITOR_ENJINE"] === true || window["IS_NATIVE_EDITOR"] === true || window["native"] !== undefined)
		{
			window['AscFonts'].onSuccess && window['AscFonts'].onSuccess.call(window['AscFonts'].api);
			return;
		}

		var url = "../../../../sdkjs/common/libfont/engine/";
		var useWasm = false;
		var webAsmObj = window["WebAssembly"];
		if (typeof webAsmObj === "object")
		{
			if (typeof webAsmObj["Memory"] === "function")
			{
				if ((typeof webAsmObj["instantiateStreaming"] === "function") || (typeof webAsmObj["instantiate"] === "function"))
					useWasm = true;
			}
		}

		var engine_name_ext = useWasm ? ".js" : "_ie.js";
		var _onSuccess = function(){
		};
		var _onError = function(){
			window['AscFonts'].onError();
		};

		AscCommon.loadScript(url + "fonts" + engine_name_ext, _onSuccess, _onError);
	};

	function FontStream(data, size)
	{
		this.data = data;
		this.size = size;
	}

	window['AscFonts'].FontStream = FontStream;

	window['AscFonts'].FT_Common = {
		UintToInt : function(v)
		{
			return (v>2147483647)?v-4294967296:v;
		},
		UShort_To_Short : function(v)
		{
			return (v>32767)?v-65536:v;
		},
		IntToUInt : function(v)
		{
			return (v<0)?v+4294967296:v;
		},
		Short_To_UShort : function(v)
		{
			return (v<0)?v+65536:v;
		}
	};

	function CPointer()
	{
		this.obj  = null; // TODO: remove
		this.data = null;
		this.pos  = 0;
	}

	function FT_Memory()
	{
		this.canvas = document.createElement('canvas');
		this.canvas.width = 1;
		this.canvas.height = 1;
		this.ctx    = this.canvas.getContext('2d');

		this.Alloc = function(size)
		{
			var p = new CPointer();
			p.data = new Uint8Array(size);
			p.pos = 0;
			return p;
		};
		this.AllocHeap = function()
		{
			// TODO: нужно посмотреть, как эта память будет использоваться.
			// нужно ли здесь делать стек, либо все время от нуля делать??
		};
		this.CreateStream = function(size)
		{
			return new FontStream(new Uint8Array(size), size);
		};
	}

	window['AscFonts'].FT_Memory = FT_Memory;
	window['AscFonts'].g_memory = new FT_Memory();

	// память для растеризации буквы
	function CRasterMemory()
	{
		this.width = 0;
		this.height = 0;
		this.pitch = 0;

		this.m_oBuffer = null;
		this.CheckSize = function(w, h)
		{
			let extra = 10; // с запасом под device pixelratio
			if (this.width < (w + extra) || this.height < (h + extra))
			{
				this.width = Math.max(this.width, w + extra);
				this.pitch = 4 * this.width;
				this.height = Math.max(this.height, h + extra);

				this.m_oBuffer = null;
				this.m_oBuffer = window['AscFonts'].g_memory.ctx.createImageData(this.width, this.height);
			}
		};
	}
	window['AscFonts'].raster_memory = new CRasterMemory();

	window['AscFonts'].registeredFontManagers = [];
	window['AscFonts'].getDefaultBlitting = function()
	{
		var isUseMap = false;
		if (AscCommon.AscBrowser.isAndroidNativeApp)
			isUseMap = true;
		else if (AscCommon.AscBrowser.isIE && !AscCommon.AscBrowser.isArm)
			isUseMap = true;
		return isUseMap;
	};
	window['AscFonts'].setDefaultBlitting = function(value)
	{
		var defaultValue = window['AscFonts'].getDefaultBlitting();
		var newValue = value ? defaultValue : !defaultValue;
		if (window['AscFonts'].use_map_blitting === newValue)
			return;

		window['AscFonts'].use_map_blitting = newValue;
		var arrManagers = window['AscFonts'].registeredFontManagers;
		for (var i = 0, count = arrManagers.length; i < count; i++)
		{
			arrManagers[i].ClearFontsRasterCache();
			arrManagers[i].InitializeRasterMemory();
		}
	};
	window['AscFonts'].use_map_blitting = window['AscFonts'].getDefaultBlitting();

})(window, undefined);
