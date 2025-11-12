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
	var g_mainHeaderHeight = 50;
	window.onresize = function()
	{
		AscViewer.checkApplicationScale();
	
		document.getElementById("leftPanel").style.width = g_positionSplitter + "px";
		document.getElementById("buttonBookmarks").style.left = "10px";
		document.getElementById("buttonBookmarks").style.width = ((g_positionSplitter >> 1) - 20) + "px";
		document.getElementById("buttonThumbnails").style.left = ((g_positionSplitter >> 1) + 10) + "px";
		document.getElementById("buttonThumbnails").style.width = ((g_positionSplitter >> 1) - 20) + "px";
		document.getElementById("panelBookmarks").style.width = g_positionSplitter + "px";
		document.getElementById("panelThumbnails").style.width = g_positionSplitter + "px";
		document.getElementById("mainPanel").style.left = g_positionSplitter + g_positionSplitterW + "px";
		document.getElementById("mainPanel").style.width = window.innerWidth - g_positionSplitter - g_positionSplitterW + "px";
		document.getElementById("mainPanel").style.height = (window.innerHeight - g_mainHeaderHeight) + "px";
		document.getElementById("mainPanelHeader").style.left = g_positionSplitter + g_positionSplitterW + "px";
		document.getElementById("mainPanelHeader").style.width = window.innerWidth - g_positionSplitter - g_positionSplitterW + "px";

		var footerPanelStyle = document.getElementById("headerPanel");
		footerPanelStyle.style.left = ((window.innerWidth - g_positionSplitter - g_positionSplitterW - footerPanelStyle.offsetWidth) >> 1) + "px";
		footerPanelStyle.style.top = ((g_mainHeaderHeight - footerPanelStyle.offsetHeight) >> 1) + "px";

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

	var options = {};
	//options.enginePath = "./../src/engine/"; // FOR NO-MINIMIZED TEST (index.html)
	//options.theme = { type : "dark" };
	//options.fontsPath = "https://url_to_fonts/";
	window.Viewer = new window.AscViewer.CViewer("mainPanel", options);
	window.Thumbnails = window.Viewer.createThumbnails("panelThimbnailsNatural");
	
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

	document.getElementById("zoomMode").addEventListener("change", function(e) {
		var selectElement = e.target;
		var value = selectElement.value;
		switch (value)
		{
			case "zmWidth":
			{
				window.Viewer.setZoomMode(AscCommon.ViewerZoomMode.Width);
				break;
			}
			case "zmPage":
			{
				window.Viewer.setZoomMode(AscCommon.ViewerZoomMode.Page);
				break;
			}
			default:
			{
				if (0 === value.indexOf("custom"))
					window.Viewer.setZoom(parseInt(value.substr(6)));
				break;
			}
		}
	});

	document.getElementById("selectMode").addEventListener("change", function(e) {
		var selectElement = e.target;
		var value = selectElement.value;
		switch (value)
		{
			case "smText":
			{
				window.Viewer.setTargetType("text");
				break;
			}
			case "smHand":
			default:
			{
				window.Viewer.setTargetType("hand");
				break;
			}
		}
	});

	var zoom_values = [50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 320, 340, 360, 380, 400, 425, 450, 475, 500];
	document.getElementById("buttonZoomIn").addEventListener("click", function(e) {
		var count = zoom_values.length;

		var curZoom = (window.Viewer.getZoom() >> 0);
		var newZoom = zoom_values[count - 1];
		for (var i = 0; i < count; i++)
		{
			if (curZoom < zoom_values[i])
			{
				newZoom = zoom_values[i];
				break;
			}
		}

		if (newZoom <= zoom_values[0] || newZoom >= zoom_values[count - 1])
			return;

		window.Viewer.setZoom(newZoom);
	});

	document.getElementById("buttonZoomOut").addEventListener("click", function(e) {
		var count = zoom_values.length;

		var curZoom = (window.Viewer.getZoom() >> 0);
		var newZoom = zoom_values[count - 1];
		for (var i = (count - 1); i >= 0; i--)
		{
			if (curZoom > zoom_values[i])
			{
				newZoom = zoom_values[i];
				break;
			}
		}

		if (newZoom <= zoom_values[0] || newZoom >= zoom_values[count - 1])
			return;

		window.Viewer.setZoom(newZoom);
	});

	document.getElementById("buttonRotateLeft").addEventListener("click", function(e) {
		window.Viewer.rotatePage(undefined, -90, true);
	});

	document.getElementById("buttonRotateRight").addEventListener("click", function(e) {
		window.Viewer.rotatePage(undefined, 90, true);
	});
}
