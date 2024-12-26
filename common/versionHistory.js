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
(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
  function(window, undefined) {
  /** @constructor */
  function asc_CVersionHistory(newObj) {
    this.docId = null;
    this.url = null;
    this.urlChanges = null;
    this.currentChangeId = -1;
    this.newChangeId = -1;
    this.colors = null;
    this.changes = null;
    this.token = null;
    this.isRequested = null;
    this.serverVersion = null;
    this.documentSha256 = null;
    this.userId = null;
    this.userName = null;
    this.userColor = null;
    this.dateOfRevision = null;

    if (newObj) {
      this.update(newObj);
    }
  }

  asc_CVersionHistory.prototype.update = function(newObj)
  {
    let bUpdate =  this.docId !== newObj.docId
                            || this.url !== newObj.url
                            || this.urlChanges !== newObj.urlChanges
                            || this.currentChangeId > newObj.currentChangeId;

    if (bUpdate)
    {
      this.docId            = newObj.docId;
      this.url              = newObj.url;
      this.urlChanges       = newObj.urlChanges;
      this.currentChangeId  = -1;
      this.changes          = null;
	  this.token            = newObj.token;
    }

    this.colors         = newObj.colors;
    this.newChangeId    = newObj.currentChangeId;
	this.isRequested    = newObj.isRequested;
	this.serverVersion  = newObj.serverVersion;
    this.userId         = newObj.userId;
    this.userName       = newObj.userName;
    this.userColor      = newObj.userColor;
    this.dateOfRevision = newObj.dateOfRevision;

	this.documentSha256 = newObj.documentSha256;
    return bUpdate;
  };
  asc_CVersionHistory.prototype.applyChanges = function(editor) {
    //in case of errors in longAction locks, this.changes can be null
    if (!this.changes) {
      return;
    }
    var color;
    this.newChangeId = (null == this.newChangeId) ? (this.changes.length - 1) : this.newChangeId;
    for (let i = this.currentChangeId + 1; i <= this.newChangeId && i < this.changes.length; ++i)
    {
      color = this.colors[i];
      let currentColor = (color ? new CDocumentColor((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF) : new CDocumentColor(191, 255, 199))
      editor._coAuthoringSetChanges(this.changes[i], i !== this.newChangeId ? null : currentColor);
    }
    this.currentChangeId = this.newChangeId;
  };
  asc_CVersionHistory.prototype.asc_setDocId = function(val) {
    this.docId = val;
  };
  asc_CVersionHistory.prototype.asc_setUrl = function(val) {
    this.url = val;
  };
  asc_CVersionHistory.prototype.asc_setUrlChanges = function(val) {
    this.urlChanges = val;
  };
  asc_CVersionHistory.prototype.asc_setCurrentChangeId = function(val) {
    this.currentChangeId = val;
  };
  asc_CVersionHistory.prototype.asc_setArrColors = function(val) {
    this.colors = val;
  };
  asc_CVersionHistory.prototype.asc_setToken = function(val) {
    this.token = val;
  };
  asc_CVersionHistory.prototype.asc_setIsRequested = function(val) {
    this.isRequested = val;
  };
  asc_CVersionHistory.prototype.asc_setServerVersion = function(val) {
    this.serverVersion = val;
  };
  asc_CVersionHistory.prototype.asc_setDocumentSha256 = function(val) {
    this.documentSha256 = val;
  };
  asc_CVersionHistory.prototype.asc_SetUserId = function(val)
  {
    this.userId = val;
  }
  asc_CVersionHistory.prototype.asc_SetUserName = function(val)
  {
    this.userName = val;
  }
  asc_CVersionHistory.prototype.asc_SetDateOfRevision = function(val)
  {
    this.dateOfRevision = val;
  }
  asc_CVersionHistory.prototype.asc_SetUserColor = function (val)
  {
    this.userColor = val;
  }

  window["Asc"].asc_CVersionHistory = window["Asc"]["asc_CVersionHistory"] = asc_CVersionHistory;
  prot = asc_CVersionHistory.prototype;
  prot["asc_setDocId"] = prot.asc_setDocId;
  prot["asc_setUrl"] = prot.asc_setUrl;
  prot["asc_setUrlChanges"] = prot.asc_setUrlChanges;
  prot["asc_setCurrentChangeId"] = prot.asc_setCurrentChangeId;
  prot["asc_setArrColors"] = prot.asc_setArrColors;
  prot["asc_setToken"] = prot.asc_setToken;
  prot["asc_setIsRequested"] = prot.asc_setIsRequested;
  prot["asc_setServerVersion"] = prot.asc_setServerVersion;
  prot["asc_setDocumentSha256"] = prot.asc_setDocumentSha256;
  prot["asc_SetUserId"] = prot.asc_SetUserId;
  prot["asc_SetUserName"] = prot.asc_SetUserName;
  prot["asc_SetUserColor"] = prot.asc_SetUserColor;
  prot["asc_SetDateOfRevision"] = prot.asc_SetDateOfRevision;
})(window);
