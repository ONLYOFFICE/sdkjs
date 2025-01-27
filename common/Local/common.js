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
	AscCommon.baseEditorsApi.prototype._openChartOrLocalDocument = function ()
	{
		if (this.isFrameEditor())
		{
			return this._openEmptyDocument();
		}
	};
	AscCommon.baseEditorsApi.prototype.onEndLoadFile2 = AscCommon.baseEditorsApi.prototype.onEndLoadFile;
	AscCommon.baseEditorsApi.prototype.onEndLoadFile = function(result)
	{
		if (this.isFrameEditor())
		{
			return this.onEndLoadFile2(result);
		}

		if (this.isLoadFullApi && this.DocInfo && this._isLoadedModules())
		{
			this.asc_registerCallback('asc_onDocumentContentReady', function(){
				DesktopOfflineUpdateLocalName(Asc.editor || editor);

				setTimeout(function(){window["UpdateInstallPlugins"]();}, 10);
			});

			AscCommon.History.UserSaveMode = true;
			window["AscDesktopEditor"]["LocalStartOpen"]();
		}
	};

	AscCommon.baseEditorsApi.prototype["local_sendEvent"] = function()
	{
		return this.sendEvent.apply(this, arguments);
	};

	AscCommon.baseEditorsApi.prototype["asc_setLocalRestrictions"] = function(value, is_from_app)
	{
		this.localRestrintions = value;
		if (value !== Asc.c_oAscLocalRestrictionType.None)
			this.asc_addRestriction(Asc.c_oAscRestrictionType.View);
		else
			this.asc_removeRestriction(Asc.c_oAscRestrictionType.View);

		if (is_from_app)
			return;

		window["AscDesktopEditor"] && window["AscDesktopEditor"]["SetLocalRestrictions"] && window["AscDesktopEditor"]["SetLocalRestrictions"](value);
	};
	AscCommon.baseEditorsApi.prototype["asc_getLocalRestrictions"] = function()
	{
		if (undefined === this.localRestrintions)
			return Asc.c_oAscLocalRestrictionType.None;
		return this.localRestrintions;
	};

	AscCommon.baseEditorsApi.prototype["startExternalConvertation"] = function(type)
	{
		let params = "";
		try {
			params = JSON.stringify(this["getAdditionalSaveParams"]());
		}
		catch (e) {
			params = "";
		}
		this.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Waiting);
		window["AscDesktopEditor"]["startExternalConvertation"](type, params);
	};
	AscCommon.baseEditorsApi.prototype["endExternalConvertation"] = function()
	{
		this.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Waiting);
	};
})(window);

/////////////////////////////////////////////////////////
//////////////       FONTS       ////////////////////////
/////////////////////////////////////////////////////////
AscFonts.CFontFileLoader.prototype.LoadFontAsync = function(basePath, callback)
{
	this.callback = callback;
    if (-1 !== this.Status)
        return true;
		
	this.Status = 2;

	var xhr = new XMLHttpRequest();
	xhr.fontFile = this;
	xhr.open('GET', "ascdesktop://fonts/" + this.Id, true);
	xhr.responseType = 'arraybuffer';

	if (xhr.overrideMimeType)
		xhr.overrideMimeType('text/plain; charset=x-user-defined');
	else
		xhr.setRequestHeader('Accept-Charset', 'x-user-defined');

	xhr.onload = function()
	{
		if (this.status !== 200)
		{
			xhr.fontFile.Status = 1;
			return;
		}

		this.fontFile.Status = 0;

		let fontStreams = AscFonts.g_fonts_streams;
		let streamIndex = fontStreams.length;
		if (this.response)
		{
			let data = new Uint8Array(this.response);
			fontStreams[streamIndex] = new AscFonts.FontStream(data, data.length);
		}
		else
		{
			fontStreams[streamIndex] = AscFonts.CreateFontData3(this.responseText);
		}

		this.fontFile.SetStreamIndex(streamIndex);

		if (null != this.fontFile.callback)
			this.fontFile.callback();
		if (this.fontFile["externalCallback"])
			this.fontFile["externalCallback"]();
	};

	xhr.send(null);
};
AscFonts.CFontFileLoader.prototype["LoadFontAsync"] = AscFonts.CFontFileLoader.prototype.LoadFontAsync;

window["DesktopOfflineAppDocumentEndLoad"] = function(_url, _data, _len)
{
	var editor = Asc.editor || window.editor;
	AscCommon.g_oDocumentUrls.documentUrl = _url;
	if (AscCommon.g_oDocumentUrls.documentUrl.indexOf("file:") != 0)
	{
		if (AscCommon.g_oDocumentUrls.documentUrl.indexOf("/") != 0)
			AscCommon.g_oDocumentUrls.documentUrl = "/" + AscCommon.g_oDocumentUrls.documentUrl;
		AscCommon.g_oDocumentUrls.documentUrl = "file://" + AscCommon.g_oDocumentUrls.documentUrl;
	}

	editor.setOpenedAt(Date.now());
	AscCommon.g_oIdCounter.m_sUserId = window["AscDesktopEditor"]["CheckUserId"]();
	if (_data === "")
	{
        editor.sendEvent("asc_onError", c_oAscError.ID.ConvertationOpenError, c_oAscError.Level.Critical);
		return;
	}

	var binaryArray = undefined;
	if (0 === _data.indexOf("binary_content://"))
	{
		var bufferArray = window["AscDesktopEditor"]["GetOpenedFile"](_data);
		if (bufferArray)
			binaryArray = new Uint8Array(bufferArray);
		else
		{
			editor.sendEvent("asc_onError", c_oAscError.ID.ConvertationOpenError, c_oAscError.Level.Critical);
			return;
		}
	}

	var file = new AscCommon.OpenFileResult();
	file.data = binaryArray ? binaryArray : getBinaryArray(_data, _len);
	file.bSerFormat = AscCommon.checkStreamSignature(file.data, AscCommon.c_oSerFormat.Signature);
	file.url = _url;
	editor.openDocument(file);

	editor.asc_SetFastCollaborative(false);
	DesktopOfflineUpdateLocalName(editor);

	window["DesktopAfterOpen"](editor);

	// why?
	// this.onUpdateDocumentModified(AscCommon.History.Have_Changes());

	editor.sendEvent("asc_onDocumentPassword", ("" != editor.currentPassword) ? true : false);
};

/////////////////////////////////////////////////////////
//////////////       IMAGES      ////////////////////////
/////////////////////////////////////////////////////////

let isOverrideDocumentUrls = window['Asc']['VisioEditorApi'] ? false : true;

if (isOverrideDocumentUrls)
{
	var prot = AscCommon.DocumentUrls.prototype;
	prot.mediaPrefix = 'media/';
	prot.init = function(urls)
	{
	};
	prot.getUrls = function()
	{
		return this.urls;
	};
	prot.addUrls = function(urls)
	{
	};
	prot.addImageUrl = function(strPath, url)
	{
	};
	prot.getImageUrl = function(strPath)
	{
		if (0 === strPath.indexOf('theme'))
			return null;

		if (window.editor && window.editor.ThemeLoader && window.editor.ThemeLoader.ThemesUrl != "" && strPath.indexOf(window.editor.ThemeLoader.ThemesUrl) == 0)
			return null;

		return this.documentUrl + "/media/" + strPath;
	};
	prot.getImageLocal = function(_url)
	{
		let url = _url.replaceAll("%20", " ");
		var _first = this.documentUrl + "/media/";
		if (0 === url.indexOf(_first))
			return url.substring(_first.length);

		if (window.editor && window.editor.ThemeLoader && 0 === url.indexOf(editor.ThemeLoader.ThemesUrlAbs))
		{
			return url.substring(editor.ThemeLoader.ThemesUrlAbs.length);
		}

		return null;
	};
	prot.imagePath2Local = function(imageLocal)
	{
		if (imageLocal && this.mediaPrefix === imageLocal.substring(0, this.mediaPrefix.length))
			imageLocal = imageLocal.substring(this.mediaPrefix.length);
		return imageLocal;
	};
	prot.getUrl = function(strPath)
	{
		if (0 === strPath.indexOf('theme'))
			return null;

		if (window.editor && window.editor.ThemeLoader && window.editor.ThemeLoader.ThemesUrl != "" && strPath.indexOf(window.editor.ThemeLoader.ThemesUrl) == 0)
			return null;

		if (strPath == "Editor.xlsx")
		{
			var test = this.documentUrl + "/" + strPath;
			if (window["AscDesktopEditor"]["IsLocalFileExist"](test))
				return test;
			return undefined;
		}

		return this.documentUrl + "/media/" + strPath;
	};
	prot.getLocal = function(url)
	{
		return this.getImageLocal(url);
	};
	prot.isThemeUrl = function(sUrl)
	{
		return sUrl && (0 === sUrl.indexOf('theme'));
	};
}

AscCommon.sendImgUrls = function(api, images, callback)
{
	var _data = [];
	for (var i = 0; i < images.length; i++)
	{
		var _url = window["AscDesktopEditor"]["LocalFileGetImageUrl"](images[i]);
		_data[i] = { url: AscCommon.g_oDocumentUrls.getUrl(_url), path : _url };
	}
	callback(_data);
};

window['Asc']["CAscWatermarkProperties"].prototype["showFileDialog"] = window['Asc']["CAscWatermarkProperties"].prototype["asc_showFileDialog"] = function ()
{
    if(!this.Api || !this.DivId)
        return;

    var t = this.Api;
    var _this = this;

    window["AscDesktopEditor"]["OpenFilenameDialog"]("images", false, function(_file) {
		var file = _file;
		if (Array.isArray(file))
			file = file[0];
		if (!file)
			return;

        var url = window["AscDesktopEditor"]["LocalFileGetImageUrl"](file);
        var urls = [AscCommon.g_oDocumentUrls.getImageUrl(url)];

        t.ImageLoader.LoadImagesWithCallback(urls, function(){
            if(urls.length > 0)
            {
                _this.ImageUrl = urls[0];
                _this.Type = Asc.c_oAscWatermarkType.Image;
                _this.drawTexture();
                t.sendEvent("asc_onWatermarkImageLoaded");
            }
        });
    });
};
window["Asc"]["asc_CBullet"].prototype["showFileDialog"] = window["Asc"]["asc_CBullet"].prototype["asc_showFileDialog"] = function()
{
	var Api = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;
	var _this = this;
	window["AscDesktopEditor"]["OpenFilenameDialog"]("images", false, function(_file) {
		var file = _file;
		if (Array.isArray(file))
			file = file[0];
		if (!file)
			return;

		var url = window["AscDesktopEditor"]["LocalFileGetImageUrl"](file);
		var urls = [AscCommon.g_oDocumentUrls.getImageUrl(url)];

		Api.ImageLoader.LoadImagesWithCallback(urls, function(){
			if(urls.length > 0)
			{
				_this.fillBulletImage(urls[0]);
				Api.sendEvent("asc_onBulletImageLoaded", _this);
			}
		});
	});
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
	let count = arrayChanges.length;
	window["AscDesktopEditor"]["LocalFileSaveChanges"]((count > 100000) ? arrayChanges : arrayChanges.join("\",\""), deleteIndex, count);
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

window["UpdateInstallPlugins"] = function()
{
	var _pluginsTmp = JSON.parse(window["AscDesktopEditor"]["GetInstallPlugins"]());
	_pluginsTmp[0]["url"] = _pluginsTmp[0]["url"].split(" ").join("%20");
	_pluginsTmp[1]["url"] = _pluginsTmp[1]["url"].split(" ").join("%20");

	var _plugins = { "url" : _pluginsTmp[0]["url"], "pluginsData" : [] };
	for (var k = 0; k < 2; k++)
	{
		var _pluginsCur = _pluginsTmp[k];

		var _len = _pluginsCur["pluginsData"].length;
		for (var i = 0; i < _len; i++)
		{
			// TODO: здесь нужно прокинуть флаг isSystemInstall, указывающий на то, что этот плагин нельзя удалить, он не в папке пользователя
			//_pluginsCur["pluginsData"][i]["isSystemInstall"] = (k == 0) ? true : false;
			_pluginsCur["pluginsData"][i]["baseUrl"] = _pluginsCur["url"] + _pluginsCur["pluginsData"][i]["guid"].substring(4) + "/";
			_plugins["pluginsData"].push(_pluginsCur["pluginsData"][i]);
		}
	}

	for (var i = 0; i < _plugins["pluginsData"].length; i++)
	{
		var _plugin = _plugins["pluginsData"][i];
		//_plugin["baseUrl"] = _plugins["url"] + _plugin["guid"].substring(4) + "/";

		var isSystem = false;
		for (var j = 0; j < _plugin["variations"].length; j++)
		{
			var _variation = _plugin["variations"][j];
			if (_variation["initDataType"] == "desktop")
			{
				isSystem = true;
				break;
			}
		}

		if (isSystem)
		{
			_plugins["pluginsData"].splice(i, 1);
			--i;
		}
	}

	var _editor = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;

	if (!window.IsFirstPluginLoad)
	{
		_editor.asc_registerCallback("asc_onPluginsReset", function ()
		{

			if (_editor.pluginsManager)
			{
				_editor.pluginsManager.unregisterAll();
			}

		});

		window.IsFirstPluginLoad = true;
	}

	_editor.sendEvent("asc_onPluginsReset");
	window.g_asc_plugins.sendPluginsInit(_plugins);
};

AscCommon.InitDragAndDrop = function(oHtmlElement, callback) {
	if ("undefined" != typeof(FileReader) && null != oHtmlElement) {
		oHtmlElement["ondragover"] = function (e) {
			e.preventDefault();
			e.dataTransfer.dropEffect = AscCommon.CanDropFiles(e) ? 'copy' : 'none';
            if (e.dataTransfer.dropEffect == "copy")
            {
                var editor = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;
                editor.beginInlineDropTarget(e);
            }
			return false;
		};
		oHtmlElement["ondrop"] = function (e) {
			e.preventDefault();

            var editor = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;
            editor.endInlineDropTarget(e);

			var _files = window["AscDesktopEditor"]["GetDropFiles"]();
			let countInserted = 0;
			if (0 !== _files.length)
			{
				let imageFiles = [];
				for (var i = 0; i < _files.length; i++)
				{
					if (window["AscDesktopEditor"]["IsImageFile"](_files[i]))
					{
						if (_files[i] === "")
							continue;

						let resImage = window["AscDesktopEditor"]["LocalFileGetImageUrl"](_files[i]);

						if (resImage)
						{
							imageFiles.push(AscCommon.g_oDocumentUrls.getImageUrl(resImage));
							++countInserted;
						}
						break;
					}
				}

				countInserted = imageFiles.length;
				if (0 !== countInserted)
					editor._addImageUrl(imageFiles);
			}

			if (0 === countInserted)
			{
                // test html
                var htmlValue = e.dataTransfer.getData("text/html");
                if (htmlValue)
                {
                    editor["pluginMethod_PasteHtml"](htmlValue);
                    return;
                }

                var textValue = e.dataTransfer.getData("text/plain");
                if (textValue)
                {
                    editor["pluginMethod_PasteText"](textValue);
                    return;
                }
			}
		};
	}
};

window["DesktopOfflineAppDocumentSignatures"] = function(_json)
{
	var _editor = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;

	_editor.signatures = [];

	var _signatures = [];
	if ("" != _json)
	{
		try
		{
			_signatures = JSON.parse(_json);
		}
		catch (err)
		{
			_signatures = [];
		}
	}

	var _count = _signatures["count"];
	var _data = _signatures["data"];
	var _sign;
	var _add_sign;

	var _images_loading = [];
	for (var i = 0; i < _count; i++)
	{
		_sign = _data[i];
		_add_sign = new window["AscCommon"].asc_CSignatureLine();

		_add_sign.guid = _sign["guid"];
		_add_sign.valid = _sign["valid"];
		_add_sign.image = (_add_sign.valid == 0) ? _sign["image_valid"] : _sign["image_invalid"];
		_add_sign.image = "data:image/png;base64," + _add_sign.image;
		_add_sign.signer1 = _sign["name"];
		_add_sign.id = i;
		_add_sign.date = _sign["date"];
		_add_sign.isvisible = window["asc_IsVisibleSign"](_add_sign.guid);
		_add_sign.correct();

		_editor.signatures.push(_add_sign);

		_images_loading.push(_add_sign.image);
	}

	if (!window.FirstSignaturesCall)
	{
		_editor.asc_registerCallback("asc_onAddSignature", function (guid)
		{

			var _api = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;
			_api.sendEvent("asc_onUpdateSignatures", _api.asc_getSignatures(), _api.asc_getRequestSignatures());

		});
		_editor.asc_registerCallback("asc_onRemoveSignature", function (guid)
		{

			var _api = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;
			_api.sendEvent("asc_onUpdateSignatures", _api.asc_getSignatures(), _api.asc_getRequestSignatures());

		});
		_editor.asc_registerCallback("asc_onUpdateSignatures", function (signatures, requested)
		{
			var _api = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;

			if (0 === signatures.length)
				_api.asc_removeRestriction(Asc.c_oAscRestrictionType.OnlySignatures);
			else
				_api.asc_addRestriction(Asc.c_oAscRestrictionType.OnlySignatures);
		});
	}
	window.FirstSignaturesCall = true;

	_editor.ImageLoader.LoadImagesWithCallback(_images_loading, function() {
		if (this.WordControl)
			this.WordControl.OnRePaintAttack();
		else if (this._onShowDrawingObjects)
			this._onShowDrawingObjects();
	}, null);

	_editor.sendEvent("asc_onUpdateSignatures", _editor.asc_getSignatures(), _editor.asc_getRequestSignatures());
};

window["DesktopSaveQuestionReturn"] = function(isNeedSaved)
{
	if (isNeedSaved)
	{
		var _editor = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;
		_editor.asc_Save(false);
	}
	else
	{
		window.SaveQuestionObjectBeforeSign = null;
	}
};

window["OnNativeReturnCallback"] = function(name, obj)
{
	var _api = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;
	_api.sendEvent(name, obj);
};

window["asc_IsVisibleSign"] = function(guid)
{
	var _editor = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;

	var isVisible = false;
	// detect visible/unvisible
	var _req = _editor.asc_getAllSignatures();
	for (var i = 0; i < _req.length; i++)
	{
		if (_req[i].id == guid)
		{
			isVisible = true;
			break;
		}
	}

	return isVisible;
};

window["asc_LocalRequestSign"] = function(guid, width, height, isView)
{
	if (isView !== true && width === undefined)
	{
		width = 100;
		height = 100;
	}

	var _editor = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;
	if (_editor.isRestrictionView())
		return;

	var _length = _editor.signatures.length;
	for (var i = 0; i < _length; i++)
	{
		if (_editor.signatures[i].guid == guid)
		{
			if (isView === true)
			{
				window["AscDesktopEditor"]["ViewCertificate"](_editor.signatures[i].id);
			}
			return;
		}
	}

	if (!_editor.isDocumentModified())
	{
		_editor.sendEvent("asc_onSignatureClick", guid, width, height, window["asc_IsVisibleSign"](guid));
		return;
	}

	window.SaveQuestionObjectBeforeSign = { guid : guid, width : width, height : height };
	window["AscDesktopEditor"]["SaveQuestion"]();
};

window["DesktopAfterOpen"] = function(_api)
{
	_api.asc_registerCallback("asc_onSignatureDblClick", function (guid, width, height)
	{
		window["asc_LocalRequestSign"](guid, width, height, true);
	});

	let langs = AscCommon.spellcheckGetLanguages();
	let langs_array = [];
	for (let item in langs)
	{
		if (!langs.hasOwnProperty(item))
			continue;
		langs_array.push(item);
	}

	_api.sendEvent('asc_onSpellCheckInit', langs_array);
};

function getBinaryArray(_data, _len)
{
	return AscCommon.Base64.decode(_data, false, _len);
}

// encryption ----------------------------------
var _proto = Asc['asc_docs_api'] || Asc['spreadsheet_api'] || Asc['VisioEditorApi'];
_proto.prototype["pluginMethod_OnEncryption"] = function(obj)
{
	var _editor = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;
	switch (obj.type)
	{
		case "generatePassword":
		{
			if ("" == obj["password"])
			{
				AscCommon.History.UserSavedIndex = _editor.LastUserSavedIndex;

				if (window.editor)
					_editor.UpdateInterfaceState();
				else
					_editor.onUpdateDocumentModified(AscCommon.History.Have_Changes());

				_editor.LastUserSavedIndex = undefined;

				_editor.sendEvent("asc_onError", "There is no connection with the blockchain! End-to-end encryption mode is disabled.", c_oAscError.Level.NoCritical);
				if (window["AscDesktopEditor"])
					window["AscDesktopEditor"]["CryptoMode"] = 0;
				return;
			}

			_editor.currentDocumentInfoNext = obj["docinfo"];

			window["DesktopOfflineAppDocumentStartSave"](window.doadssIsSaveAs, obj["password"], true, obj["docinfo"] ? obj["docinfo"] : "");
            window["AscDesktopEditor"]["buildCryptedStart"]();
			break;
		}
		case "getPasswordByFile":
		{
			if ("" != obj["password"])
			{
				_editor.currentPassword = obj["password"];

                if (window.isNativeOpenPassword)
                {
                    window["AscDesktopEditor"]["NativeViewerOpen"](obj["password"]);
                }
                else
                {
                    var _param = ("<m_sPassword>" + AscCommon.CopyPasteCorrectString(obj["password"]) + "</m_sPassword>");
                    window["AscDesktopEditor"]["SetAdvancedOptions"](_param);
                }
			}
			else
			{
				this._onNeedParams(undefined, true);
			}
			break;
		}
        case "encryptData":
        case "decryptData":
        {
            AscCommon.EncryptionWorker.receiveChanges(obj);
            break;
        }
	}
};

AscCommon.getBinaryArray = getBinaryArray;
// -------------------------------------------

// меняем среду
//AscBrowser.isSafari = false;
//AscBrowser.isSafariMacOs = false;
//window.USER_AGENT_SAFARI_MACOS = false;
