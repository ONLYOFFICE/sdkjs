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

// Import
var c_oAscError = Asc.c_oAscError;

/////////////////////////////////////////////////////////
//////////////        OPEN       ////////////////////////
/////////////////////////////////////////////////////////

(/**
 * @param {jQuery} $
 * @param {Window} window
 * @param {undefined} undefined
 */
	function($, window, undefined) {

	var asc = window["Asc"];
	var prot;

	asc['spreadsheet_api'].prototype._OfflineAppDocumentStartLoad = function()
	{
		this.asc_registerCallback('asc_onDocumentContentReady', function(){
			DesktopOfflineUpdateLocalName(window["Asc"]["editor"]);
		});
	
		window["AscDesktopEditor"]["LocalStartOpen"]();
	};
	
	asc['spreadsheet_api'].prototype._OfflineAppDocumentEndLoad = function(_data)
	{
		AscCommon.g_oIdCounter.m_sUserId = window["AscDesktopEditor"]["CheckUserId"]();
		if (_data == "")
		{
			this.sendEvent("asc_onError", c_oAscError.ID.ConvertationError, c_oAscError.Level.Critical);
			return;
		}

		this.openDocument(_data);
		AscCommon.History.UserSaveMode = true;
		
		DesktopOfflineUpdateLocalName(this);
		
		this.onUpdateDocumentModified(AscCommon.History.Is_Modified());
	};
	
	asc['spreadsheet_api'].prototype._onNeedParams = function(data) 
	{
		var cp = JSON.parse("{\"codepage\":46,\"delimiter\":1}");
		cp['encodings'] = AscCommon.getEncodingParams();
		this.handlers.trigger("asc_onAdvancedOptions", new AscCommon.asc_CAdvancedOptions(Asc.c_oAscAdvancedOptionsID.CSV, cp), AscCommon.c_oAscAdvancedOptionsAction.Open);
	};
	
	asc['spreadsheet_api'].prototype.asc_addImageDrawingObject = function(url)
	{
		var _url = window["AscDesktopEditor"]["LocalFileGetImageUrl"](url);
		
		var ws = this.wb.getWorksheet();
		if (ws) 
		{
			var _url = window["AscDesktopEditor"]["LocalFileGetImageUrl"](url);
			ws.objectRender.addImageDrawingObject(AscCommon.g_oDocumentUrls.getImageUrl(_url) , null);
		}
	};
	asc['spreadsheet_api'].prototype.asc_showImageFileDialog = function()
	{
		window["AscDesktopEditor"]["LocalFileGetImageUrlFromOpenFileDialog"]();
	};
	asc['spreadsheet_api'].prototype.asc_addImage = function()
	{
	  window["AscDesktopEditor"]["LocalFileGetImageUrlFromOpenFileDialog"]();
	};
})(jQuery, window);

window["Asc"]['spreadsheet_api'].prototype.asc_setAdvancedOptions = function(idOption, option) 
{
	window["AscDesktopEditor"]["SetAdvancedOptions"]("" + option.asc_getCodePage(), "" + option.asc_getDelimiter());
};
window["Asc"]['spreadsheet_api'].prototype["asc_setAdvancedOptions"] = window["Asc"]['spreadsheet_api'].prototype.asc_setAdvancedOptions;

window["asc_initAdvancedOptions"] = function()
{	
	window["Asc"]["editor"]._onNeedParams(undefined);
};

window["DesktopOfflineAppDocumentEndLoad"] = function(_url, _data)
{
	AscCommon.g_oDocumentUrls.documentUrl = _url;
	if (AscCommon.g_oDocumentUrls.documentUrl.indexOf("file:") != 0)
	{
		if (AscCommon.g_oDocumentUrls.documentUrl.indexOf("/") != 0)
			AscCommon.g_oDocumentUrls.documentUrl = "/" + AscCommon.g_oDocumentUrls.documentUrl;
		AscCommon.g_oDocumentUrls.documentUrl = "file://" + AscCommon.g_oDocumentUrls.documentUrl;
	}
	
    window["Asc"]["editor"]._OfflineAppDocumentEndLoad(_data);
};

/////////////////////////////////////////////////////////
//////////////        CHANGES       /////////////////////
/////////////////////////////////////////////////////////
AscCommon.CHistory.prototype.Reset_SavedIndex = function(IsUserSave)
{
	if (true === this.Is_UserSaveMode())
	{
		this.SavedIndex = this.Index;
		if (true === IsUserSave)
		{
			this.UserSavedIndex = this.Index;
			this.ForceSave      = false;
		}
	}
	else
	{
		this.SavedIndex = this.Index;
		this.ForceSave  = false;
	}
};

AscCommon.CHistory.prototype.Is_Modified = function(IsNotUserSave, IsNoSavedNoModifyed) 
{
	var checkIndex = (this.Is_UserSaveMode() && !IsNotUserSave) ? this.UserSavedIndex : this.SavedIndex;
	if (-1 === this.Index && null === checkIndex && false === this.ForceSave) 
	{
		if (window["AscDesktopEditor"])
		{
			if (0 != window["AscDesktopEditor"]["LocalFileGetOpenChangesCount"]())
				return true;
			if (!window["AscDesktopEditor"]["LocalFileGetSaved"]() && IsNoSavedNoModifyed !== true)
				return true;
		}
		return false;
	}
	return (this.Index != checkIndex || true === this.ForceSave);
};
	
window["DesktopOfflineAppDocumentApplyChanges"] = function(_changes)
{
	for (var i = 0, l = _changes.length; i < l; ++i) 
	{
		window["Asc"]["editor"].CoAuthoringApi.onSaveChanges(_changes[i], null, true);
    }
};

/////////////////////////////////////////////////////////
////////////////        SAVE       //////////////////////
/////////////////////////////////////////////////////////
window["Asc"]['spreadsheet_api'].prototype.onUpdateDocumentModified = function(bIsModified) 
{
    // Обновляем только после окончания сохранения
    if (this.canSave) {
      this.handlers.trigger("asc_onDocumentModifiedChanged", bIsModified);
      this._onUpdateDocumentCanSave();

      if (undefined !== window["AscDesktopEditor"]) {
        window["AscDesktopEditor"]["onDocumentModifiedChanged"](AscCommon.History ? AscCommon.History.Is_Modified(undefined, true) : bValue);
      }
    }
};
  
window["Asc"]['spreadsheet_api'].prototype.asc_Save = function (isNoUserSave, isSaveAs)
{
	if (this.isChartEditor || AscCommon.c_oAscAdvancedOptionsAction.None !== this.advancedOptionsAction)
		return;
    	
	var t = this;
    if (true !== isNoUserSave)
        this.IsUserSave = true;
	
	if (this.IsUserSave)
	{
		this.LastUserSavedIndex = AscCommon.History.UserSavedIndex;
	}

    if (true === this.canSave && !this.isLongAction())
	{
		var _isNaturalSave = this.IsUserSave;
		this.canSave = false;
		this.CoAuthoringApi.askSaveChanges(function(e){t.onSaveCallback(e);});
		
		if (this.CoAuthoringApi.onUnSaveLock)
			this.CoAuthoringApi.onUnSaveLock();
		
		if (_isNaturalSave === true)
			window["DesktopOfflineAppDocumentStartSave"](isSaveAs);
	}
};
window["Asc"]['spreadsheet_api'].prototype.asc_DownloadAs = function(typeFile, bIsDownloadEvent) 
{
	this.asc_Save(false, true);
};

window["Asc"]['spreadsheet_api'].prototype.asc_isOffline = function()
{
	return true;
};

window["DesktopOfflineAppDocumentStartSave"] = function(isSaveAs)
{
	window["Asc"]["editor"].sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Save);
	
	var _param = "";
	if (isSaveAs === true)
		_param += "saveas=true;";
	if (AscCommon.AscBrowser.isRetina)
		_param += "retina=true;";
	
	window["AscDesktopEditor"]["LocalFileSave"](_param);
};
window["DesktopOfflineAppDocumentEndSave"] = function(error)
{
	window["Asc"]["editor"].sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Save);
	if (0 == error)
		DesktopOfflineUpdateLocalName(window["Asc"]["editor"]);
	else
		AscCommon.History.UserSavedIndex = window["Asc"]["editor"].LastUserSavedIndex;		
	
	window["Asc"]["editor"].onUpdateDocumentModified(AscCommon.History.Is_Modified());
	window["Asc"]["editor"].LastUserSavedIndex = undefined;
	
	if (2 == error)
		window["Asc"]["editor"].sendEvent("asc_onError", c_oAscError.ID.ConvertationError, c_oAscError.Level.NoCritical);
};

window["Asc"]['spreadsheet_api'].prototype["asc_addImageDrawingObject"] = window["Asc"]['spreadsheet_api'].prototype.asc_addImageDrawingObject;
window["Asc"]['spreadsheet_api'].prototype["asc_showImageFileDialog"] = window["Asc"]['spreadsheet_api'].prototype.asc_showImageFileDialog;
window["Asc"]['spreadsheet_api'].prototype["asc_Save"] = window["Asc"]['spreadsheet_api'].prototype.asc_Save;
window["Asc"]['spreadsheet_api'].prototype["asc_DownloadAs"] = window["Asc"]['spreadsheet_api'].prototype.asc_DownloadAs;
window["Asc"]['spreadsheet_api'].prototype["asc_isOffline"] = window["Asc"]['spreadsheet_api'].prototype.asc_isOffline;
window["Asc"]['spreadsheet_api'].prototype["asc_addImage"] = window["Asc"]['spreadsheet_api'].prototype.asc_addImage;

window["DesktopOfflineAppDocumentAddImageEnd"] = function(url)
{
	if (url == "")
		return;
	
	var ws = window["Asc"]["editor"].wb.getWorksheet();
    if (ws) 
	{
		var _url = window["AscDesktopEditor"]["LocalFileGetImageUrl"](url);
        ws.objectRender.addImageDrawingObject(AscCommon.g_oDocumentUrls.getImageUrl(_url) , null);
    }
};

window["on_editor_native_message"] = function(sCommand, sParam)
{
	if (!window["Asc"]["editor"])
		return;
	
	if (sCommand == "save")
		window["Asc"]["editor"].asc_Save();
	else if (sCommand == "saveAs")
		window["Asc"]["editor"].asc_Save(false, true);
	else if (sCommand == "print")
		window["Asc"]["editor"].asc_Print();
};