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

// Import
var c_oAscError = Asc.c_oAscError;
/////////////////////////////////////////////////////////
//////////////        CHANGES       /////////////////////
/////////////////////////////////////////////////////////
AscCommon.CHistory.prototype.Reset_SavedIndex = function(IsUserSave)
{
	this.SavedIndex = (null === this.SavedIndex && -1 === this.Index ? null : this.Index);
	if (true === this.Is_UserSaveMode())
	{
		if (true === IsUserSave)
		{
			this.UserSavedIndex = this.Index;
			this.ForceSave      = false;
		}
	}
	else
	{
		this.ForceSave  = false;
	}
};
AscCommon.CHistory.prototype.Have_Changes = function(IsNotUserSave, IsNoSavedNoModifyed)
{
	if (true === this.Is_UserSaveMode() && true !== IsNotUserSave)
	{
		if (-1 === this.Index && null === this.UserSavedIndex && false === this.ForceSave)
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

		if (this.Index != this.UserSavedIndex || true === this.ForceSave)
			return true;

		return false;
	}
	else
	{
		if (-1 === this.Index && null === this.SavedIndex && false === this.ForceSave)
			return false;

		if (this.Index != this.SavedIndex || true === this.ForceSave)
			return true;

		return false;
	}
};
	
window["DesktopOfflineAppDocumentApplyChanges"] = function(_changes)
{
	editor._coAuthoringSetChanges(_changes, null);
    //editor["asc_nativeApplyChanges"](_changes);
	//editor["asc_nativeCalculateFile"]();
};

let protAPI = window['Asc']['VisioEditorApi'].prototype;

/////////////////////////////////////////////////////////
////////////////        SAVE       //////////////////////
/////////////////////////////////////////////////////////
protAPI.SetDocumentModified = function(bValue)
{
    this.isDocumentModify = bValue;
    this.sendEvent("asc_onDocumentModifiedChanged");

    if (undefined !== window["AscDesktopEditor"])
    {
        window["AscDesktopEditor"]["onDocumentModifiedChanged"](AscCommon.History ? AscCommon.History.Have_Changes(undefined, true) : bValue);
    }
};

protAPI._saveLocalCheck = function()
{
	return true;//this._saveCheck();
};

protAPI.asc_Save = function (isNoUserSave, isSaveAs, isResaveAttack, options)
{
    if (!isResaveAttack && !isSaveAs && !this.asc_isDocumentCanSave())
        return;

    if (true !== isNoUserSave)
        this.IsUserSave = true;
	
	if (this.IsUserSave)
	{
		this.LastUserSavedIndex = AscCommon.History.UserSavedIndex;
	}

    if (true === this.canSave && this._saveLocalCheck())
	{
		var _isNaturalSave = this.IsUserSave;
		/*
		this.canSave = false;

		var t = this;
		this.CoAuthoringApi.askSaveChanges(function(e) {
			t._onSaveCallback(e);
		});

		if (this.CoAuthoringApi.onUnSaveLock)
			this.CoAuthoringApi.onUnSaveLock();
		*/

		if (_isNaturalSave === true)
			window["DesktopOfflineAppDocumentStartSave"](isSaveAs, undefined, undefined, undefined, options);
	}
};
protAPI["getAdditionalSaveParams"] = function()
{
	return {
		"documentLayout" : {
			"openedAt" : this.openedAt
		},
		"locale" : this.asc_getLocale(),
		"translate" : AscCommon.translateManager.mapTranslate
	};
};
window["DesktopOfflineAppDocumentStartSave"] = function(isSaveAs, password, isForce, docinfo, options)
{
	editor.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Save);

	var _param = "";
	if (isSaveAs === true)
		_param += "saveas=true;";

	var jsonOptions = editor["getAdditionalSaveParams"]();

	if (options && options.advancedOptions)
	{
		let nativeOptions = options.advancedOptions.asc_getNativeOptions();
		if (nativeOptions)
		{
			jsonOptions["nativeOptions"] = nativeOptions;
			jsonOptions["nativeOptions"]["currentPage"] = editor.GetCurrentVisiblePage() + 1;
		}
	}

	window["AscDesktopEditor"]["LocalFileSave"](_param, (password === undefined) ? editor.currentPassword : password, docinfo, (options && options.fileType) ? options.fileType : 0, JSON.stringify(jsonOptions));
};
window["DesktopOfflineAppDocumentEndSave"] = function(error, hash, password)
{
	editor.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Save);
	if (2 == error)
		editor.sendEvent("asc_onError", c_oAscError.ID.ConvertationSaveError, c_oAscError.Level.NoCritical);
};
protAPI.asc_DownloadAsNatural = protAPI.asc_DownloadAs;
protAPI.asc_DownloadAs = function(options)
{
    if (options && options.isNaturalDownload)
        return this.asc_DownloadAsNatural(options);
	this.asc_Save(false, true, undefined, options);
};

protAPI["asc_Save"] = protAPI.asc_Save;
protAPI["asc_DownloadAs"] = protAPI.asc_DownloadAs;

protAPI.asc_isOffline = function()
{
	return true;
};
