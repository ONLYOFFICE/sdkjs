/*
 *
 * (c) Copyright Ascensio System Limited 2010-2016
 *
 * This program is freeware. You can redistribute it and/or modify it under the terms of the GNU 
 * General Public License (GPL) version 3 as published by the Free Software Foundation (https://www.gnu.org/copyleft/gpl.html). 
 * In accordance with Section 7(a) of the GNU GPL its Section 15 shall be amended to the effect that 
 * Ascensio System SIA expressly excludes the warranty of non-infringement of any third-party rights.
 *
 * THIS PROGRAM IS DISTRIBUTED WITHOUT ANY WARRANTY; WITHOUT EVEN THE IMPLIED WARRANTY OF MERCHANTABILITY OR
 * FITNESS FOR A PARTICULAR PURPOSE. For more details, see GNU GPL at https://www.gnu.org/copyleft/gpl.html
 *
 * You can contact Ascensio System SIA by email at sales@onlyoffice.com
 *
 * The interactive user interfaces in modified source and object code versions of ONLYOFFICE must display 
 * Appropriate Legal Notices, as required under Section 5 of the GNU GPL version 3.
 *
 * Pursuant to Section 7  3(b) of the GNU GPL you must retain the original ONLYOFFICE logo which contains 
 * relevant author attributions when distributing the software. If the display of the logo in its graphic 
 * form is not reasonably feasible for technical reasons, you must include the words "Powered by ONLYOFFICE" 
 * in every copy of the program you distribute. 
 * Pursuant to Section 7  3(e) we decline to grant you any rights under trademark law for use of our trademarks.
 *
*/
"use strict";

/////////////////////////////////////////////////////////
//////////////       FONTS       ////////////////////////
/////////////////////////////////////////////////////////
AscFonts.CFontFileLoader.prototype.LoadFontAsync = function(basePath, _callback, isEmbed)
{
	this.callback = _callback;
    if (-1 != this.Status)
        return true;
		
	var oThis = this;
	this.Status = 2;
	if (window["AscDesktopEditor"] !== undefined && !this.CanUseOriginalFormat)
	{
		this.callback = null;		
		window["AscDesktopEditor"]["LoadFontBase64"](this.Id);
		this._callback_font_load();
		return;
	}

	var xhr = new XMLHttpRequest();
	xhr.open('GET', "ascdesktop://fonts/" + this.Id, true);
	xhr.responseType = 'arraybuffer';

	if (xhr.overrideMimeType)
		xhr.overrideMimeType('text/plain; charset=x-user-defined');
	else
		xhr.setRequestHeader('Accept-Charset', 'x-user-defined');

	xhr.onload = function()
	{
		if (this.status != 200)
		{
			oThis.Status = 1;
			return;
		}

		oThis.Status = 0;

		var fontStreams = AscFonts.g_fonts_streams;
		if (this.response)
		{
			var __font_data_idx = fontStreams.length;
			var _uintData = new Uint8Array(this.response);
			fontStreams[__font_data_idx] = new AscFonts.FT_Stream(_uintData, _uintData.length);
			oThis.SetStreamIndex(__font_data_idx);
		}
		else
		{
			var __font_data_idx = fontStreams.length;
			fontStreams[__font_data_idx] = AscFonts.CreateFontData3(this.responseText);
			oThis.SetStreamIndex(__font_data_idx);

			if (null != oThis.callback)
				oThis.callback();
		}
	};

	xhr.send(null);
};

/////////////////////////////////////////////////////////
//////////////       IMAGES      ////////////////////////
/////////////////////////////////////////////////////////
var prot = AscCommon.DocumentUrls.prototype;
prot.mediaPrefix = 'media/';
prot.init = function(urls) {
};
prot.getUrls = function() {
	return this.urls;
};
prot.addUrls = function(urls){
};
prot.addImageUrl = function(strPath, url){
};
prot.getImageUrl = function(strPath){
	if (0 === strPath.indexOf('theme'))
		return null;

	if (window.editor && window.editor.ThemeLoader && window.editor.ThemeLoader.ThemesUrl != "" && strPath.indexOf(window.editor.ThemeLoader.ThemesUrl) == 0)
		return null;

	return this.documentUrl + "/media/" + strPath;
};
prot.getImageLocal = function(url){
	var _first = this.documentUrl + "/media/";
	if (0 == url.indexOf(_first))
		return url.substring(_first.length);

	if (window.editor && window.editor.ThemeLoader && 0 == url.indexOf(editor.ThemeLoader.ThemesUrlAbs)) {
		return url.substring(editor.ThemeLoader.ThemesUrlAbs.length);
	}

	return null;
};
prot.imagePath2Local = function(imageLocal){
	return this.getImageLocal(imageLocal);
};
prot.getUrl = function(strPath){
	if (0 === strPath.indexOf('theme'))
		return null;

	if (window.editor && window.editor.ThemeLoader && window.editor.ThemeLoader.ThemesUrl != "" && strPath.indexOf(window.editor.ThemeLoader.ThemesUrl) == 0)
		return null;

	return this.documentUrl + "/media/" + strPath;
};
prot.getLocal = function(url){
	return this.getImageLocal(url);
};

AscCommon.sendImgUrls = function(api, images, callback)
{
	var _data = [];
	for (var i = 0; i < images.length; i++)
	{
		var _url = window["AscDesktopEditor"]["LocalFileGetImageUrl"](images[i]);
		_data[i] = { url: images[i], path : AscCommon.g_oDocumentUrls.getImageUrl(_url) };
	}
	callback(_data);
};

/////////////////////////////////////////////////////////
////////////////        SAVE       //////////////////////
/////////////////////////////////////////////////////////
function DesktopOfflineUpdateLocalName(_api)
{
	var _name = window["AscDesktopEditor"]["LocalFileGetSourcePath"]();
	
	var _ind1 = _name.lastIndexOf("\\");
	var _ind2 = _name.lastIndexOf("/");
	
	if (_ind1 == -1)
		_ind1 = 1000000;
	if (_ind2 == -1)
		_ind2 = 1000000;
	
	var _ind = Math.min(_ind1, _ind2);
	if (_ind != 1000000)
		_name = _name.substring(_ind + 1);
	
	_api.documentTitle = _name;
	_api.sendEvent("asc_onDocumentName", _name);
	window["AscDesktopEditor"]["SetDocumentName"](_name);
}

AscCommon.CDocsCoApi.prototype.askSaveChanges = function(callback)
{
    callback({"saveLock": false});
};
AscCommon.CDocsCoApi.prototype.saveChanges = function(arrayChanges, deleteIndex, excelAdditionalInfo)
{
	window["AscDesktopEditor"]["LocalFileSaveChanges"](arrayChanges.join("\",\""), deleteIndex, arrayChanges.length);
	//this.onUnSaveLock();
};

window["NativeCorrectImageUrlOnCopy"] = function(url)
{
	AscCommon.g_oDocumentUrls.getImageUrl(url);
};
window["NativeCorrectImageUrlOnPaste"] = function(url)
{
	return window["AscDesktopEditor"]["LocalFileGetImageUrl"](url);
};

AscCommon.InitDragAndDrop = function(oHtmlElement, callback) {
	if ("undefined" != typeof(FileReader) && null != oHtmlElement) {
		oHtmlElement["ondragover"] = function (e) {
			e.preventDefault();
			e.dataTransfer.dropEffect = AscCommon.CanDropFiles(e) ? 'copy' : 'none';
			return false;
		};
		oHtmlElement["ondrop"] = function (e) {
			e.preventDefault();
			
			var _files = window["AscDesktopEditor"]["GetDropFiles"]();
			for (var i = 0; i < _files.length; i++)
			{
				if (window["AscDesktopEditor"]["IsImageFile"](_files[i]))
				{
					window["DesktopOfflineAppDocumentAddImageEnd"](_files[i]);
					break;
				}
			}
		};
	}
}

// меняем среду
//AscBrowser.isSafari = false;
//AscBrowser.isSafariMacOs = false;
//window.USER_AGENT_SAFARI_MACOS = false;