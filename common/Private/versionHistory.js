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

    if (newObj) {
      this.update(newObj);
    }
  }

  asc_CVersionHistory.prototype.update = function(newObj) {
    var bUpdate = (this.docId !== newObj.docId || this.url !== newObj.url || this.urlChanges !== newObj.urlChanges || this.currentChangeId > newObj.currentChangeId);
    if (bUpdate) {
      this.docId = newObj.docId;
      this.url = newObj.url;
      this.urlChanges = newObj.urlChanges;
      this.currentChangeId = -1;
      this.changes = null;
    }
    this.colors = newObj.colors;
    this.newChangeId = newObj.currentChangeId;
    return bUpdate;
  };
  asc_CVersionHistory.prototype.applyChanges = function(editor) {
    var color;
    this.newChangeId = (null == this.newChangeId) ? (this.changes.length - 1) : this.newChangeId;
    for (var i = this.currentChangeId + 1; i <= this.newChangeId && i < this.changes.length; ++i) {
      color = this.colors[i];
      editor._coAuthoringSetChanges(this.changes[i], i !== this.newChangeId ? null : (color ? new CDocumentColor((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF) : new CDocumentColor(191, 255, 199)));
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

  window["Asc"].asc_CVersionHistory = window["Asc"]["asc_CVersionHistory"] = asc_CVersionHistory;
  prot = asc_CVersionHistory.prototype;
  prot["asc_setDocId"] = prot.asc_setDocId;
  prot["asc_setUrl"] = prot.asc_setUrl;
  prot["asc_setUrlChanges"] = prot.asc_setUrlChanges;
  prot["asc_setCurrentChangeId"] = prot.asc_setCurrentChangeId;
  prot["asc_setArrColors"] = prot.asc_setArrColors;
})(window);

AscCommon.baseEditorsApi.prototype.asc_showRevision = function(newObj) {
  if (!newObj.docId) {
    return;
  }
  if (this.isCoAuthoringEnable) {
    this.asc_coAuthoringDisconnect();
  }

  var bUpdate = true;
  if (null === this.VersionHistory) {
    this.VersionHistory = new window["Asc"].asc_CVersionHistory(newObj);
  } else {
    bUpdate = this.VersionHistory.update(newObj);
  }
  // ToDo должно быть все общее
  if (bUpdate) {
    this.asc_CloseFile();

    this.DocInfo.put_Id(this.VersionHistory.docId);
    this.DocInfo.put_Url(this.VersionHistory.url);
    this.documentUrlChanges = this.VersionHistory.urlChanges;
    this.asc_setDocInfo(this.DocInfo);
    this.asc_LoadDocument(true);
  } else if (this.VersionHistory.currentChangeId < newObj.currentChangeId) {
    // Нужно только добавить некоторые изменения
    AscCommon.CollaborativeEditing.Clear_CollaborativeMarks();
    editor.VersionHistory.applyChanges(editor);
    AscCommon.CollaborativeEditing.Apply_Changes();
  }
};
AscCommon.baseEditorsApi.prototype['asc_showRevision'] = AscCommon.baseEditorsApi.prototype.asc_showRevision;