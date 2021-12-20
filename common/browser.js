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
(
/**
* @param {Window} window
* @param {undefined} undefined
*/
function (window, undefined) {
var AscBrowser = {
    userAgent : "",
    isIE : false,
    isMacOs : false,
    isSafariMacOs : false,
    isAppleDevices : false,
    isAndroid : false,
    isMobile : false,
    isGecko : false,
    isChrome : false,
    isOpera : false,
    isOperaOld : false,
    isWebkit : false,
    isSafari : false,
    isArm : false,
    isMozilla : false,
    isLinuxOS : false,
    retinaPixelRatio : 1,
    isVivaldiLinux : false,
    isSailfish : false,
    isEmulateDevicePixelRatio : false,
    isNeedEmulateUpload : false,
    chromeVersion : 70,
    iosVersion : 13,
    isAndroidNativeApp : false
};

// user agent lower case
AscBrowser.userAgent = navigator.userAgent.toLowerCase();

// ie detect
AscBrowser.isIE =  (AscBrowser.userAgent.indexOf("msie") > -1 ||
                    AscBrowser.userAgent.indexOf("trident") > -1 ||
                    AscBrowser.userAgent.indexOf("edge") > -1);

AscBrowser.isIeEdge = (AscBrowser.userAgent.indexOf("edge/") > -1);

AscBrowser.isIE9 =  (AscBrowser.userAgent.indexOf("msie9") > -1 || AscBrowser.userAgent.indexOf("msie 9") > -1);
AscBrowser.isIE10 =  (AscBrowser.userAgent.indexOf("msie10") > -1 || AscBrowser.userAgent.indexOf("msie 10") > -1);

// macOs detect
AscBrowser.isMacOs = (AscBrowser.userAgent.indexOf('mac') > -1);

// chrome detect
AscBrowser.isChrome = !AscBrowser.isIE && (AscBrowser.userAgent.indexOf("chrome") > -1);
if (AscBrowser.isChrome)
{
    var checkVersion = AscBrowser.userAgent.match(/chrom(e|ium)\/([0-9]+)\./);
    if (checkVersion && checkVersion[2])
        AscBrowser.chromeVersion = parseInt(checkVersion[2], 10);
}

// safari detect
AscBrowser.isSafari = !AscBrowser.isIE && !AscBrowser.isChrome && (AscBrowser.userAgent.indexOf("safari") > -1);

// macOs safari detect
AscBrowser.isSafariMacOs = (AscBrowser.isSafari && AscBrowser.isMacOs);

// apple devices detect
AscBrowser.isAppleDevices = (AscBrowser.userAgent.indexOf("ipad") > -1 ||
                             AscBrowser.userAgent.indexOf("iphone") > -1 ||
                             AscBrowser.userAgent.indexOf("ipod") > -1);
if (!AscBrowser.isAppleDevices && AscBrowser.isSafariMacOs && navigator.platform === "MacIntel" && (navigator.maxTouchPoints > 1))
	AscBrowser.isAppleDevices = true;

if (AscBrowser.isAppleDevices)
{
	var iosversion = AscBrowser.iosVersion;
	try
	{
		var v = (navigator.appVersion).match(/OS (\d+)_(\d+)_?(\d+)?/);
		if (!v) v = (navigator.appVersion).match(/Version\/(\d+).(\d+)/);
		//[parseInt(v[1], 10), parseInt(v[2], 10), parseInt(v[3] || 0, 10)];
		iosversion = parseInt(v[1], 10);
	}
	catch (err)
	{
	}
	AscBrowser.iosVersion = iosversion;
}

// android devices detect
AscBrowser.isAndroid = (AscBrowser.userAgent.indexOf("android") > -1);

// mobile detect
AscBrowser.isMobile = /android|avantgo|blackberry|blazer|compal|elaine|fennec|hiptop|iemobile|ip(hone|od|ad)|iris|kindle|lge |maemo|midp|mmp|opera m(ob|in)i|palm( os)?|phone|p(ixi|re)\/|plucker|pocket|psp|symbian|treo|up\.(browser|link)|vodafone|wap|windows (ce|phone)|xda|xiino/i.test(navigator.userAgent || navigator.vendor || window.opera);

// gecko detect
AscBrowser.isGecko = (AscBrowser.userAgent.indexOf("gecko/") > -1);

// opera detect
AscBrowser.isOpera = (!!window.opera || AscBrowser.userAgent.indexOf("opr/") > -1);
AscBrowser.isOperaOld = (!!window.opera);

// webkit detect
AscBrowser.isWebkit = !AscBrowser.isIE && (AscBrowser.userAgent.indexOf("webkit") > -1);

// arm detect
AscBrowser.isArm = (AscBrowser.userAgent.indexOf("arm") > -1);

AscBrowser.isMozilla = !AscBrowser.isIE && (AscBrowser.userAgent.indexOf("firefox") > -1);

AscBrowser.isLinuxOS = (AscBrowser.userAgent.indexOf(" linux ") > -1);

AscBrowser.isVivaldiLinux = AscBrowser.isLinuxOS && (AscBrowser.userAgent.indexOf("vivaldi") > -1);

AscBrowser.isSailfish = (AscBrowser.userAgent.indexOf("sailfish") > -1);

AscBrowser.isEmulateDevicePixelRatio = (AscBrowser.userAgent.indexOf("emulatedevicepixelratio") > -1);

AscBrowser.isNeedEmulateUpload = (AscBrowser.userAgent.indexOf("needemulateupload") > -1);

AscBrowser.isAndroidNativeApp = (AscBrowser.userAgent.indexOf("ascandroidwebview") > -1);

AscBrowser.zoom = 1;

AscBrowser.isCustomScaling = function()
{
    return (Math.abs(AscBrowser.retinaPixelRatio - 1) > 0.001) ? true : false;
};

AscBrowser.isCustomScalingAbove2 = function()
{
    return (AscBrowser.retinaPixelRatio > 1.999) ? true : false;
};

AscBrowser.checkZoom = function()
{
    if (AscBrowser.isSailfish && AscBrowser.isEmulateDevicePixelRatio)
    {
        var scale = 1;
        if (screen.width <= 540)
            scale = 1.5;
        else if (screen.width > 540 && screen.width <= 768)
            scale = 2;
        else if (screen.width > 768)
            scale = 3;

        AscBrowser.retinaPixelRatio = scale;
        window.devicePixelRatio = scale;
        return;
    }

    var zoomValue = AscCommon.checkDeviceScale();
	AscBrowser.retinaPixelRatio = zoomValue.applicationPixelRatio;
	AscBrowser.zoom = zoomValue.zoom;

    AscCommon.correctApplicationScale(zoomValue);
};

AscBrowser.checkZoom();

AscBrowser.convertToRetinaValue = function(value, isScale)
{
	if (isScale === true)
		return ((value * AscBrowser.retinaPixelRatio) + 0.5) >> 0;
	else
		return ((value / AscBrowser.retinaPixelRatio) + 0.5) >> 0;
};

    //--------------------------------------------------------export----------------------------------------------------
    window['AscCommon'] = window['AscCommon'] || {};
    window['AscCommon'].AscBrowser = AscBrowser;
})(window);
