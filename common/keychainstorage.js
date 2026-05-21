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

(function(exports){

    exports.AscCrypto = exports.AscCrypto || {};
    var AscCrypto = exports.AscCrypto;

    AscCrypto.Storage = {};

    // command type. "private" - means accessible only to the user making the request
    AscCrypto.Storage.CommandType = {
        private : "private",
        public : "public"
    };

    // command keys.
    AscCrypto.Storage.CommandKey = {
        keySign : "keySign",
        keyCrypt : "keyCrypt"
    };

    /* storage:
    {
        id : { type : AscCrypto.Storage.CommandKey, value { date : ..., ... } },
        ...
    }
    */

    // list of commands.
    AscCrypto.Storage.CommandName = {
        // add a record for the current user
        // { type : add, value : [{value},{value}] }
        add : "add",
        // remove a record by id. not used yet
        // { type : remove, value : [id1, id2, ...] }
        remove : "remove",
        // replace a record by id for the current user (if the record does not exist - do nothing)
        // { type : remove, value : [{rec, rec}] }
        replace : "replace",
        // for the current user. return all records with the specified key
        // { type : get, value : [key1, key2] }
        // for self - return private as well. for others - no
        get : "get",
        // return the user object (their data) by one of the key values.
        // if nothing is specified - return for the current user
        // (in principle, you can also send the key to look up, there will be few records => the key can be omitted)
        getUserInfo : "getUserInfo"
    };

    function CItem()
    {
        this["id"] = undefined;
        this["key"] = undefined;
        this["value"] = undefined;
    }
    CItem.prototype.generate = function(key, value)
    {
        this["id"] = AscCommon.randomBytes(20).base58();
        this["key"] = key;
        this["value"] = {};

        if (value)
        {
            for (let prop in value)
            {
                if (value.hasOwnProperty(prop))
                    this["value"][prop] = value[prop];
            }
        }

        let date = new Date();
        this["value"]["date"] = date.toISOString();
    };
    CItem.prototype.store = function(obj)
    {
        obj[this["id"]] = { "key" : this["key"], "value" : this["value"] };
    };

    function IStorage()
    {
    }
    // interface
    CStorageLocalStorage.prototype.command = function(items, callback) {}

    /**
     * @extends {IStorage}
     */
    function CStorageLocalStorage()
    {
        IStorage.call(this);
    }
    CStorageLocalStorage.prototype = Object.create(IStorage.prototype);
    CStorageLocalStorage.prototype.constructor = CStorageLocalStorage;

    CStorageLocalStorage.prototype.getStorageValue = function()
    {
        try
        {
            return JSON.parse(window.localStorage.getItem("oo-crypto-object"));
        }
        catch (e)
        {
            return {};
        }
    };
    CStorageLocalStorage.prototype.setStorageValue = function(value)
    {
        try
        {
            window.localStorage.setItem("oo-crypto-object", JSON.stringify(value));
            return true;
        }
        catch (e)
        {
        }
        return false;
    };

    CStorageLocalStorage.prototype.command = function(command, callback)
    {
        let localValue = this.getStorageValue();
        if (!localValue)
            localValue = {};

        let isUpdate = false;

        let records = command.value;
        let returnKeys = command.callback;

        switch (command["type"])
        {
            case AscCrypto.Storage.CommandName.add:
            {
                for (let i = 0, len = records.length; i < len; i++)
                {
                    let newItem = new CItem();
                    newItem.generate(records[i]["key"], records[i]["value"]);
                    newItem.store(localValue);
                    isUpdate = true;
                }
                break;
            }
            case AscCrypto.Storage.CommandName.remove:
            {
                for (let i = 0, len = records.length; i < len; i++)
                {
                    if (localValue[records[i]])
                    {
                        delete localValue[records[i]];
                        isUpdate = true;
                    }
                }
                break;
            }
            case AscCrypto.Storage.CommandName.replace:
            {
                for (let prop in records)
                {
                    if (records.hasOwnProperty(prop) && localValue[prop])
                    {
                        delete localValue[prop];
                        localValue[prop] = records[prop];
                        isUpdate = true;
                    }
                }

                for (let i = 0, len = records.length; i < len; i++)
                {
                    if (localValue[records[i]["id"]])
                    {
                        delete localValue[records[i]["id"]];
                        localValue[records[i]["id"]] = { "key" : records[i]["key"], "value" : records[i]["value"] };
                        isUpdate = true;
                    }
                }
                break;
            }
            case AscCrypto.Storage.CommandName.get:
            {
                returnKeys = records;
                break;
            }
            case AscCrypto.Storage.CommandName.getUserInfo:
            {
                // no users in the local version
                break;
            }
            default:
                break;
        }

        if (isUpdate)
            this.setStorageValue(localValue);

        let mapReturnKeys = {};
        for (let i = 0, len = returnKeys.length; i < len; i++)
        {
            mapReturnKeys[returnKeys[i]] = true;
        }

        let returnObj = {};
        for (let prop in localValue)
        {
            if (localValue.hasOwnProperty(prop) && mapReturnKeys[localValue[prop]["key"]] === true)
            {
                returnObj[prop] = localValue[prop];
                // don't remove private ones here (this should be done on users)
            }
        }

        setTimeout(function(){
            callback && callback(returnObj);
        }, 10);
    };

    AscCrypto.Storage.CItem = CItem;
    AscCrypto.Storage.IStorage = IStorage;
    AscCrypto.Storage.CStorageLocalStorage = CStorageLocalStorage;

})(window);
