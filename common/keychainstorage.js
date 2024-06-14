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

(function(exports){

    exports.AscCrypto = exports.AscCrypto || {};
    var AscCrypto = exports.AscCrypto;

    AscCrypto.Storage = {};

    // тип команды. "private" - значит доступно только для того юзера, который делает запрос
    AscCrypto.Storage.CommandType = {
        private : "private",
        public : "public"
    };

    // ключи команды.
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

    // перечень команд.
    AscCrypto.Storage.CommandName = {
        // добавляем к текущему юзеру запись
        // { type : add, value : [{value},{value}] }
        add : "add",
        // удаляем запись по id. пока не используем
        // { type : remove, value : [id1, id2, ...] }
        remove : "remove",
        // заменяем у текущего юзера запись по id (если записи нет - то ничего не делаем)
        // { type : remove, value : [{rec, rec}] }
        replace : "replace",
        // для текущего юзера. отдаем все записи с указанным ключом
        // { type : get, value : [key1, key2] }
        // для себя - отдаем и private. для остальных- нет
        get : "get",
        // вернуть объект юзера (данные его) по одному из значений ключа.
        // если ничего не указано - то вернуть для текущего юзера
        // (в принципе можно присылать еще и ключ, по которому смотреть, записей будет мало => ключ можно упустить)
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
                // в локальной версии юзеров нет
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
                // тут приватные не удаляем (это нужно на юзерах делать)
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
