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

(function (window) {
    /**
     * A form role. The role owns the list of field full-names that belong to it.
     * The whole collection of CPdfRoleInfo is held by CPdfFormRoles and persisted
     * as a document-level "rolesInfo" section.
     * @constructor
     */
    function CPdfRoleInfo(obj) {
        obj = obj || {};
        this.id         = obj["id"]     !== undefined ? obj["id"]     : null;
        this.name       = obj["name"]   !== undefined ? obj["name"]   : "";
        this.color      = obj["color"]  !== undefined ? obj["color"]  : null; // hex string | null
        this.email      = obj["email"]  !== undefined ? obj["email"]  : null;
        this.userId     = obj["userId"]   !== undefined ? obj["userId"]   : null;
        this.userName   = obj["userName"] !== undefined ? obj["userName"] : null;
        this.filled     = obj["filled"] !== undefined ? obj["filled"] : false;
        this.date       = obj["date"]   !== undefined ? obj["date"]   : null;
        this.isDefault  = obj["isDefault"] !== undefined ? !!obj["isDefault"] : false;
        this.fieldNames = Array.isArray(obj["fieldNames"]) ? obj["fieldNames"].slice() : [];
    }

    CPdfRoleInfo.prototype.GetId     = function () { return this.id; };
    CPdfRoleInfo.prototype.GetName   = function () { return this.name; };
    CPdfRoleInfo.prototype.GetColor  = function () { return this.color; };
    CPdfRoleInfo.prototype.GetEmail  = function () { return this.email; };
    CPdfRoleInfo.prototype.GetUserId = function () { return this.userId; };
    CPdfRoleInfo.prototype.IsFilled  = function () { return !!this.filled; };
    CPdfRoleInfo.prototype.GetDate   = function () { return this.date; };
    CPdfRoleInfo.prototype.IsDefault = function () { return this.isDefault; };

    CPdfRoleInfo.prototype.SetName   = function (v) { this.name = v; };
    CPdfRoleInfo.prototype.SetColor  = function (v) { this.color = v; };
    CPdfRoleInfo.prototype.SetEmail  = function (v) { this.email = v; };
    CPdfRoleInfo.prototype.SetUserId = function (v) { this.userId = v; };
    CPdfRoleInfo.prototype.GetUserName = function () { return this.userName; };
    CPdfRoleInfo.prototype.SetUserName = function (v) { this.userName = v; };
    CPdfRoleInfo.prototype.SetFilled = function (v) { this.filled = !!v; };
    CPdfRoleInfo.prototype.SetDate   = function (v) { this.date = v; };
    CPdfRoleInfo.prototype.SetDefault = function (v) { this.isDefault = !!v; };
    CPdfRoleInfo.prototype.GetFieldNames = function () { return this.fieldNames; };

	CPdfRoleInfo.prototype.GetFieldCount = function () {
        let oDoc   = Asc.editor && Asc.editor.getPDFDoc ? Asc.editor.getPDFDoc() : null;
        let oRoles = oDoc && oDoc.GetFormRoles ? oDoc.GetFormRoles() : null;
        if (oRoles && oRoles.GetEffectiveFieldNames)
            return oRoles.GetEffectiveFieldNames(this).length;
        return this.fieldNames.length;
    };
    CPdfRoleInfo.prototype.HasFieldName  = function (name) { return -1 !== this.fieldNames.indexOf(name); };
    CPdfRoleInfo.prototype.AddFieldName  = function (name) {
        if (name != null && -1 === this.fieldNames.indexOf(name))
            this.fieldNames.push(name);
    };
    CPdfRoleInfo.prototype.RemoveFieldName = function (name) {
        let idx = this.fieldNames.indexOf(name);
        if (-1 !== idx) this.fieldNames.splice(idx, 1);
    };

    CPdfRoleInfo.prototype.ToJSON = function () {
        return {
            "id":         this.id,
            "name":       this.name,
            "color":      this.color,
            "email":      this.email,
            "userId":     this.userId,
            "userName":   this.userName,
            "filled":     this.filled,
            "date":       this.date,
            "isDefault":  this.isDefault,
            "fieldNames": this.fieldNames.slice()
        };
    };

    CPdfRoleInfo.prototype.Copy = function () {
        return new CPdfRoleInfo(this.ToJSON());
    };

    CPdfRoleInfo.prototype.IsEqual = function (other) {
        if (!other) return false;
        if (this.id !== other.id
            || this.name !== other.name
            || this.color !== other.color
            || this.email !== other.email
            || this.userId !== other.userId
            || this.userName !== other.userName
            || (!!this.filled) !== (!!other.filled)
            || this.date !== other.date
            || this.isDefault !== other.isDefault)
            return false;

        let a = this.fieldNames, b = other.fieldNames;
        if (a.length !== b.length) return false;
        for (let i = 0; i < a.length; ++i) {
            if (a[i] !== b[i]) return false;
        }
        return true;
    };

    CPdfRoleInfo.prototype.IsAllFieldsFilled = function (oDoc) {
        if (this.fieldNames.length === 0) return false;
        for (let i = 0; i < this.fieldNames.length; ++i) {
            let oField = oDoc && oDoc.GetField ? oDoc.GetField(this.fieldNames[i]) : null;
            if (oField.GetValue() == undefined)
                return false;
        }
        return true;
    };

    CPdfRoleInfo.prototype.GetSettings = function () {
        let s = new AscPDF.CPdfRoleSettings();
        s.setName(this.name);
        if (this.color) {
            let c = new AscWord.CDocumentColor();
            c.SetFromHexColor(this.color);
            s.setColor(c);
        }
        s.setFieldCount(this.GetFieldCount());
        s.setFilled(!!this.filled);
        s.setDate(this.date);
        s.setUserId(this.userId);
        s.setUserName(this.userName);
        return s;
    };

    CPdfRoleInfo.FromJSON = function (obj) {
        if (!obj) return null;
        return new CPdfRoleInfo(obj);
    };

    window["AscPDF"] = window["AscPDF"] || {};
    window["AscPDF"].CPdfRoleInfo = CPdfRoleInfo;
	CPdfRoleInfo.prototype['asc_getSettings'] = CPdfRoleInfo.prototype.GetSettings;
})(window);
