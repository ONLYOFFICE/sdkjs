/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";
(function (window, undefined) {
	var supportedScaleValues = [1, 1.25, 1.5, 1.75, 2, 2.25, 2.5, 2.75, 3, 3.5, 4, 4.5, 5];
	if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["GetSupportedScaleValues"])
		supportedScaleValues = window["AscDesktopEditor"]["GetSupportedScaleValues"]();

	// uncomment to debug all scales
	//supportedScaleValues = [];

	var isCorrectApplicationScaleEnabled = (function(){

		if (supportedScaleValues.length === 0)
			return false;

		var userAgent = navigator.userAgent.toLowerCase();
		var isAndroid = (userAgent.indexOf("android") > -1);
		var isIE = (userAgent.indexOf("msie") > -1 || userAgent.indexOf("trident") > -1 || userAgent.indexOf("edge") > -1);
		var isChrome = !isIE && (userAgent.indexOf("chrome") > -1);
		var isOperaOld = (!!window.opera);
		var isMobile = /android|avantgo|blackberry|blazer|compal|elaine|fennec|hiptop|iemobile|ip(hone|od|ad)|iris|kindle|lge |maemo|midp|mmp|opera m(ob|in)i|palm( os)?|phone|p(ixi|re)\/|plucker|pocket|psp|symbian|treo|up\.(browser|link)|vodafone|wap|windows (ce|phone)|xda|xiino/i.test(navigator.userAgent || navigator.vendor || window.opera);

		if (isAndroid || !isChrome || isOperaOld || isMobile || !document || !document.firstElementChild || !document.body)
			return false;

		return true;

	})();

	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].checkDeviceScale = function()
	{
		var retValue = {
			zoom: 1,
			devicePixelRatio: window.devicePixelRatio,
			applicationPixelRatio: window.devicePixelRatio,
			correct : false
		};

		if (!isCorrectApplicationScaleEnabled)
			return retValue;

		var systemScaling = window.devicePixelRatio;
		var bestIndex = 0;
		var bestDistance = Math.abs(supportedScaleValues[0] - systemScaling);
		var currentDistance = 0;
		for (var i = 1, len = supportedScaleValues.length; i < len; i++)
		{
			if (true)
			{
				// this is "interface adjustment" - remove after moving this to common code
				if (Math.abs(supportedScaleValues[i] - systemScaling) > 0.0001)
				{
					if (supportedScaleValues[i] > (systemScaling - 0.0001))
						break;
				}
			}

			currentDistance = Math.abs(supportedScaleValues[i] - systemScaling);
			if (currentDistance < (bestDistance - 0.0001))
			{
				bestDistance = currentDistance;
				bestIndex = i;
			}
		}

		retValue.applicationPixelRatio = supportedScaleValues[bestIndex];
		if (Math.abs(retValue.devicePixelRatio - retValue.applicationPixelRatio) > 0.01)
		{
			retValue.zoom = retValue.devicePixelRatio / retValue.applicationPixelRatio;
			retValue.correct = true;
		}
		return retValue;
	};

	var oldZoomValue = 1;
	window['AscCommon'].correctApplicationScale = function(zoomValue)
	{
		if (!zoomValue.correct && Math.abs(zoomValue.zoom - oldZoomValue) < 0.0001)
			return;
		oldZoomValue = zoomValue.zoom;
		var firstElemStyle = document.firstElementChild.style;
		if (Math.abs(oldZoomValue - 1) < 0.001)
			firstElemStyle.zoom = "normal";
		else
			firstElemStyle.zoom = 1.0 / oldZoomValue;
	};
})(window);
