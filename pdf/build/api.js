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

(function(){

	function setCanvasSize(element, width, height, is_correction)
	{
		if (element.width === width && element.height === height)
			return;

		if (true !== is_correction)
		{
			element.width = width;
			element.height = height;
			return;
		}

		var data = element.getContext("2d").getImageData(0, 0, element.width, element.height);
		element.width = width;
		element.height = height;
		element.getContext("2d").putImageData(data, 0, 0);
	};

	AscCommon.calculateCanvasSize = function(element, is_correction, is_wait_correction)
	{
		if (true !== is_correction && undefined !== element.correctionTimeout)
		{
			clearTimeout(element.correctionTimeout);
			element.correctionTimeout = undefined;
		}

		var scale = AscCommon.AscBrowser.retinaPixelRatio;
		if (Math.abs(scale - (scale >> 0)) < 0.001)
		{
			setCanvasSize(element,
				scale * parseInt(element.style.width),
				scale * parseInt(element.style.height),
				is_correction);
			return;
		}

		var rect = element.getBoundingClientRect();
		var isCorrectRect = (rect.width === 0 && rect.height === 0) ? false : true;
		if (is_wait_correction || !isCorrectRect)
		{
			var isNoVisibleElement = false;
			if (element.style.display === "none")
				isNoVisibleElement = true;
			else if (element.parentNode && element.parentNode.style.display === "none")
				isNoVisibleElement = true;

			if (!isNoVisibleElement)
			{
				element.correctionTimeout = setTimeout(function (){
					AscCommon.calculateCanvasSize(element, true);
				}, 100);
			}

			if (!isCorrectRect)
			{
				var style_width = parseInt(element.style.width);
				var style_height = parseInt(element.style.height);

				rect = {
					x: 0, left: 0,
					y: 0, top: 0,
					width: style_width, right: style_width,
					height: style_height, bottom: style_height
				};
			}
		}

		setCanvasSize(element,
			Math.round(scale * rect.right) - Math.round(scale * rect.left),
			Math.round(scale * rect.bottom) - Math.round(scale * rect.top),
			is_correction);
	};

	function loadScript(url, onSuccess, onError)
	{
		var script = document.createElement('script');
		script.type = 'text/javascript';
		script.src = url;
		script.onload = onSuccess;
		script.onerror = onError;

		document.head.appendChild(script);
	}

	var documentScriptUrl = document.currentScript ? document.currentScript.src : "";

	function CViewer(parent, options)
	{
		this.isLoadAllFonts = (AscFonts.g_font_files.length > 1) ? true : false;
		this.allFontsUrl = "";
		this.fontsErrorCounter = 0;
		this.fontsErrorMax = 3;

		this.options = options;
		this.api = this.checkOptions();

		this.viewer = new AscCommon.CViewer(parent, this.api);
	}

	CViewer.prototype.isInitFonts = false;
	CViewer.prototype["open"] = function(data)
	{
		if (!this.isLoadAllFonts && this.allFontsUrl !== "")
		{
			this.data = data;
			this.checkFonts();
			return;
		}

		this.initFonts();
		this.viewer.open(data);
	};
	CViewer.prototype["createThumbnails"] = function(parent)
	{
		var oThumbnails = new AscCommon.ThumbnailsControl(parent)
		this.viewer.setThumbnailsControl(oThumbnails);
		return oThumbnails;
	};
	CViewer.prototype["registerEvent"] = function(name, callback)
	{
		this.viewer.registerEvent(name, callback);
	};
	CViewer.prototype["resize"] = function()
	{
		this.viewer.resize();
	};
	CViewer.prototype["navigate"] = function(node)
	{
		return this.viewer.navigate(node);
	};
	CViewer.prototype["getEngine"] = function()
	{
		return this.viewer;
	};

	// private
	CViewer.prototype.initFonts = function()
	{
		if (!this.isInitFonts)
		{
			this.isInitFonts = true;
			AscFonts.checkAllFonts();
			AscFonts.g_fontApplication.Init();

			AscCommon.g_font_loader.fontFiles = AscFonts.g_font_files;
			AscCommon.g_font_loader.fontInfos = AscFonts.g_font_infos;
			AscCommon.g_font_loader.map_font_index = AscFonts.g_map_font_index;
		}
	};
	CViewer.prototype.checkFonts = function()
	{
		var onSuccess = function() {
			this.isLoadAllFonts = true;
			this.open(this.data);
			delete this.data;
		};
		var onError = function() {
			this.fontsErrorCounter++;
			if (this.fontsErrorCounter === this.fontsErrorMax)
			{
				console.error("cannot load AllFonts.js");
				return;
			}
			this.checkFonts();
		};

		loadScript(this.allFontsUrl, onSuccess.bind(this), onError.bind(this));
	};
	CViewer.prototype.checkOptions = function()
	{
		if (this.options)
		{
			var sAllFontsUrl = "";
			var sModuleUrl = "";
			var sFontsPath = "";

			if (this.options["sdkjsPath"])
			{
				sAllFontsUrl = this.options["sdkjsPath"] + "/common/AllFonts.js";
				sModuleUrl = this.options["sdkjsPath"] + "/pdf/src/engine/";
				sFontsPath = this.options["sdkjsPath"] + "/../fonts/"
			}

			if (this.options["enginePath"])
				sModuleUrl = this.options["enginePath"];

			if (this.options["fontsPath"])
				sFontsPath = this.options["fontsPath"];

			if (documentScriptUrl !== "")
			{
				var scriptDirectory = documentScriptUrl;
				scriptDirectory = scriptDirectory.substr(0,scriptDirectory.replace(/[?#].*/,"").lastIndexOf("/") + 1);

				if (sAllFontsUrl === "")
					sAllFontsUrl = scriptDirectory + "../../../common/AllFonts.js";
				if (sModuleUrl === "")
					sModuleUrl = scriptDirectory;
				if (sFontsPath === "")
					sFontsPath = scriptDirectory + "../../../../fonts/";
			}

			this.allFontsUrl = sAllFontsUrl;

			if (sModuleUrl !== "")
				window["AscViewer"]["baseEngineUrl"] = sModuleUrl;

			if (sFontsPath !== "")
				AscCommon.g_font_loader.fontFilesPath = sFontsPath;

			if (this.options["theme"])
				AscCommon.updateGlobalSkin(this.options["theme"]);
		}

		return {
			isMobileVersion : false,
			isSeparateModule : true,

			baseFontsPath : (this.options && this.options["fontsPath"]) ? this.options["fontsPath"] : undefined,

			getPageBackgroundColor : function() {
				// TODO: get color from theme
				if (this.isDarkMode)
					return { R: 0x3A, G : 0x3A, B : 0x3A };
				return { R: 0xFF, G : 0xFF, B : 0xFF };
			},

			WordControl : {
				NoneRepaintPages : false
			},

			sendEvent : function() {
			}
		};
	};

	CViewer.prototype["setZoomMode"] = function(value) { this.viewer && this.viewer.setZoomMode(value); };
	CViewer.prototype["getZoom"] = function() { return this.viewer ? (this.viewer.zoom * 100) : 100; };
	CViewer.prototype["setZoom"] = function(value) { this.viewer && this.viewer.setZoom(value / 100); };
	CViewer.prototype["setTargetType"] = function(value) { this.viewer && this.viewer.setTargetType(value); };
	CViewer.prototype["rotatePage"] = function(page, angle, ismultiply) { this.viewer && this.viewer.setRotatePage(page, angle, ismultiply); };
	CViewer.prototype["close"] = function() { this.viewer && this.viewer.close(); };

	window["AscViewer"] = window["AscViewer"] || {};
	window["AscViewer"]["CViewer"] = CViewer;

	window["AscViewer"]["checkApplicationScale"] = function()
	{
		var zoomValue = AscCommon.checkDeviceScale();
		AscCommon.AscBrowser.retinaPixelRatio = zoomValue.applicationPixelRatio;
		AscCommon.AscBrowser.zoom = zoomValue.zoom;
		AscCommon.correctApplicationScale(zoomValue);
	};

})();
