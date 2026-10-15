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
    function CPdfRoleSettings() {
        this.Name       = "";
        this.Color      = null; // AscWord.CDocumentColor | null
        this.FieldCount = 0;
        this.Date       = undefined;
        this.UserId     = undefined;
        this.UserName   = undefined;
        this.Filled     = false;
    }
    CPdfRoleSettings.prototype.getName = function () {
        return this.Name;
    };
    CPdfRoleSettings.prototype.setName = function (name) {
        this.Name = name;
    };
    CPdfRoleSettings.prototype.setColor = function (color) {
        this.Color = color ? color.Copy() : null;
    };
    CPdfRoleSettings.prototype.getColor = function () {
        return this.Color;
    };
    CPdfRoleSettings.prototype.setAscColor = function (ascColor) {
        if (undefined === ascColor || null === ascColor)
            this.Color = null;
        else
            this.Color = new AscWord.CDocumentColor(ascColor.asc_getR(), ascColor.asc_getG(), ascColor.asc_getB());
    };
    CPdfRoleSettings.prototype.getAscColor = function () {
        if (!this.Color)
            return null;

        return new Asc.asc_CColor(this.Color.r, this.Color.g, this.Color.b);
    };
    CPdfRoleSettings.prototype.haveColor = function () {
        return (!!this.Color);
    };
    CPdfRoleSettings.prototype.setFieldCount = function (count) {
        this.FieldCount = count;
    };
    CPdfRoleSettings.prototype.getFieldCount = function () {
        return this.FieldCount;
    };
    CPdfRoleSettings.prototype.getFilled = function () {
        return this.Filled;
    };
    CPdfRoleSettings.prototype.setFilled = function (isFilled) {
        this.Filled = isFilled;
    };
    CPdfRoleSettings.prototype.getDate = function () {
        return this.Date;
    };
    CPdfRoleSettings.prototype.setDate = function (date) {
        this.Date = date;
    };
    CPdfRoleSettings.prototype.getUserId = function () {
        return this.UserId;
    };
    CPdfRoleSettings.prototype.setUserId = function (userId) {
        this.UserId = userId;
    };
    CPdfRoleSettings.prototype.getUserName = function () {
        return this.UserName;
    };
    CPdfRoleSettings.prototype.setUserName = function (userName) {
        this.UserName = userName;
    };

    //--------------------------------------------------------export----------------------------------------------------
    window["AscPDF"] = window["AscPDF"] || {};
    window["AscPDF"].CPdfRoleSettings = CPdfRoleSettings;

    CPdfRoleSettings.prototype["asc_getName"]       = CPdfRoleSettings.prototype.getName;
    CPdfRoleSettings.prototype["asc_putName"]       = CPdfRoleSettings.prototype.setName;
    CPdfRoleSettings.prototype["asc_getColor"]      = CPdfRoleSettings.prototype.getAscColor;
    CPdfRoleSettings.prototype["asc_putColor"]      = CPdfRoleSettings.prototype.setAscColor;
    CPdfRoleSettings.prototype["asc_getFieldCount"] = CPdfRoleSettings.prototype.getFieldCount;
    CPdfRoleSettings.prototype["asc_getFilled"]     = CPdfRoleSettings.prototype.getFilled;
    CPdfRoleSettings.prototype["asc_getDate"]       = CPdfRoleSettings.prototype.getDate;
    CPdfRoleSettings.prototype["asc_getUserId"]     = CPdfRoleSettings.prototype.getUserId;
    CPdfRoleSettings.prototype["asc_putUserId"]     = CPdfRoleSettings.prototype.setUserId;
    CPdfRoleSettings.prototype["asc_getUserName"]   = CPdfRoleSettings.prototype.getUserName;
    CPdfRoleSettings.prototype["asc_putUserName"]   = CPdfRoleSettings.prototype.setUserName;
})(window);
