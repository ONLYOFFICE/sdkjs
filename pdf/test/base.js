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

window.onload = function()
{
	var holder = document.getElementById("mainPanel");
	holder.ondragover = function(e) 
	{
		var isFile = false;
		if (e.dataTransfer.types)
		{
			for (var i = 0, length = e.dataTransfer.types.length; i < length; ++i)
			{
				var type = e.dataTransfer.types[i].toLowerCase();
				if (type == "files" && e.dataTransfer.items && e.dataTransfer.items.length == 1)
				{
					var item = e.dataTransfer.items[0];
					if (item.kind && "file" == item.kind.toLowerCase())
					{
						isFile = true;
						break;
					}
				}
			}
		}
		e.dataTransfer.dropEffect = isFile ? "copy" : "none";
		e.preventDefault();
		return false; 
	};
	holder.ondrop = function(e) 
	{ 
		var file = e.dataTransfer.files ? e.dataTransfer.files[0] : null;
		if (!file)
		{
			e.preventDefault();
			return false;
		}
		
		var reader = new FileReader();
		reader.onload = function(e) {
			window.Viewer.open(e.target.result);
		};
		reader.readAsArrayBuffer(file);

		return false; 
	};

	var g_positionSplitter = 200;
	var g_positionSplitterW = 5;
	window.onresize = function()
	{
		var zoomValue = AscCommon.checkDeviceScale();
		AscCommon.AscBrowser.retinaPixelRatio = zoomValue.applicationPixelRatio;
		AscCommon.AscBrowser.zoom = zoomValue.zoom;
		AscCommon.correctApplicationScale(zoomValue);
	
		document.getElementById("leftPanel").style.width = g_positionSplitter + "px";
		document.getElementById("buttonBookmarks").style.left = "10px";
		document.getElementById("buttonBookmarks").style.width = ((g_positionSplitter >> 1) - 20) + "px";
		document.getElementById("buttonThumbnails").style.left = ((g_positionSplitter >> 1) + 10) + "px";
		document.getElementById("buttonThumbnails").style.width = ((g_positionSplitter >> 1) - 20) + "px";
		document.getElementById("panelBookmarks").style.width = g_positionSplitter + "px";
		document.getElementById("panelThumbnails").style.width = g_positionSplitter + "px";
		document.getElementById("mainPanel").style.left = g_positionSplitter + g_positionSplitterW + "px";
		document.getElementById("mainPanel").style.width = window.innerWidth - g_positionSplitter - g_positionSplitterW + "px";
		document.getElementById("mainPanel").style.height = window.innerHeight + "px";

		var trackBarH = 50;
		var tabsBarH = 50;
		var panelH = window.innerHeight - tabsBarH;
		document.getElementById("panelThimbnailsNatural").style.height = panelH - trackBarH + "px";
		document.getElementById("zoomPanel").style.top = panelH - trackBarH + "px";
		document.getElementById("zoomPanel").style.height = trackBarH + "px";
		document.getElementById("zoomPanel").style.left = "10px";
		document.getElementById("zoomPanel").style.width = g_positionSplitter - 20 + "px";

		document.getElementById("splitter").style.left = g_positionSplitter + "px";
		document.getElementById("splitter").style.width = g_positionSplitterW + "px";

		window.Viewer && window.Viewer.resize();
		window.Thumbnails && window.Thumbnails.resize();
	};

	window.onresize();
	window.initTrackBars();

	document.getElementById("buttonBookmarks").onclick = function()
	{
		document.getElementById("panelBookmarks").style.display = "block";
		document.getElementById("panelThumbnails").style.display = "none";
	};
	document.getElementById("buttonThumbnails").onclick = function()
	{
		document.getElementById("panelBookmarks").style.display = "none";
		document.getElementById("panelThumbnails").style.display = "block";
	};

	var g_positionSplitterPosOld = 0;
	document.getElementById("splitter").onmousedown = function(e)
	{
		g_isSplitterMoving = true;
		document.getElementById("leftPanel").style.pointerEvents = "none";
		document.getElementById("mainPanel").style.pointerEvents = "none";
		document.getElementById("splitter").style.pointerEvents = "none";

		g_positionSplitterPosOld = g_positionSplitter;
		var xPos = (e.pageX === undefined) ? e.clientX : e.pageX;

		document.body.style.cursor = "w-resize";
		document.body.onmousemove = function(e)
		{
			document.body.style.cursor = "w-resize";
			var xPosNew = (e.pageX === undefined) ? e.clientX : e.pageX;
			g_positionSplitter = g_positionSplitterPosOld + (xPosNew - xPos);
			window.onresize();
		};

		document.body.onmouseup = function(e)
		{
			document.getElementById("leftPanel").style.pointerEvents = "";
			document.getElementById("mainPanel").style.pointerEvents = "";
			document.getElementById("splitter").style.pointerEvents = "";
			document.body.onmousemove = function(e) {};
			document.body.onmouseup = function(e) {};
			document.body.style.cursor = "default";
		};
	};
	document.getElementById("splitter").onmouseup = function(e)
	{
		document.getElementById("leftPanel").style.pointerEvents = "";
		document.getElementById("mainPanel").style.pointerEvents = "";
		document.getElementById("splitter").style.pointerEvents = "";
		document.body.onmousemove = function(e) {};
		document.body.onmouseup = function(e) {};
	};

	///
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
					calculateCanvasSize(element, true);
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
			new_height = Math.round(scale * rect.bottom) - Math.round(scale * rect.top),
			is_correction);
	};
	///
	AscFonts.g_fontApplication.Init();
	window.Viewer = new AscCommon.CViewer("mainPanel");
	window.Thumbnails = new AscCommon.ThumbnailsControl("panelThimbnailsNatural");
	window.Viewer.setThumbnailsControl(window.Thumbnails);
	
	var trackbar = document.querySelectorAll('.trackbar')[0];
	trackbar.onChangedValue = function(val) {
		window.Thumbnails.setZoom(val);
	};
	
	window.Viewer.registerEvent("onStructure", function(structure){
		AscInterface.updateStructure(structure);
	});

	window.Thumbnails.registerEvent("onZoomChanged", function(value){
		trackbar.setPosition(value);
	});
}
