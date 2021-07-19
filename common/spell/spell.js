/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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

function CSpellchecker(settings)
{
	var useWasm = false;
	var webAsmObj = window["WebAssembly"];
	if (typeof webAsmObj === "object")
	{
		if (typeof webAsmObj["Memory"] === "function")
		{
			if ((typeof webAsmObj["instantiateStreaming"] === "function") || (typeof webAsmObj["instantiate"] === "function"))
				useWasm = true;
		}
	}

	var enginePath = "./spell/";
	if (settings && settings.enginePath)
	{
		enginePath = settings.enginePath;
		if (enginePath.substring(enginePath.length - 1) != "/")
			enginePath += "/";
	}

	var dictionariesPath = "./../dictionaries";
	if (settings && settings.dictionariesPath)
	{
		dictionariesPath = settings.dictionariesPath;
		if (dictionariesPath.substring(dictionariesPath.length - 1) == "/")
			dictionariesPath = dictionariesPath.substr(0, dictionariesPath.length - 1);
	}

	var isUseSharedWorker = !!window.SharedWorker;
	if (isUseSharedWorker && (false === settings.useShared))
		isUseSharedWorker = false;

	var worker_src = useWasm ? "spell.js" : "spell_ie.js";
	worker_src = enginePath + worker_src;

	var _worker = this;
	var _port = null;
	if (isUseSharedWorker)
	{
		this.worker = new SharedWorker(worker_src, "onlyoffice-spellchecker");
		_port = this.worker.port;
	}
	else
	{
		this.worker = new Worker(worker_src);
		_port = this.worker;
	}

	this.languages = {
		"1068" : "az_Latn_AZ",
		"1026" : "bg_BG",
		"1027" : "ca_ES",
		"2051" : "ca_ES_valencia",
		"1029" : "cs_CZ",
		"1030" : "da_DK",
		"3079" : "de_AT",
		"2055" : "de_CH",
		"1031" : "de_DE",
		"1032" : "el_GR",
		"3081" : "en_AU",
		"4105" : "en_CA",
		"2057" : "en_GB",
		"1033" : "en_US",
		"7177" : "en_ZA",
		"3082" : "es_ES",
		"1069" : "eu_ES",
		"1036" : "fr_FR",
		"1110" : "gl_ES",
		"1050" : "hr_HR",
		"1038" : "hu_HU",
		"1057" : "id_ID",
		"1040" : "it_IT",
		"1087" : "kk_KZ",
		"1042" : "ko_KR",
		"1134" : "lb_LU",
		"1063" : "lt_LT",
		"1062" : "lv_LV",
		"1104" : "mn_MN",
		"1044" : "nb_NO",
		"1043" : "nl_NL",
		"2068" : "nn_NO",
		"1045" : "pl_PL",
		"1046" : "pt_BR",
		"2070" : "pt_PT",
		"1048" : "ro_RO",
		"1049" : "ru_RU",
		"1051" : "sk_SK",
		"1060" : "sl_SI",
		"10266" : "sr_Cyrl_RS",
		"9242" : "sr_Latn_RS",
		"1053" : "sv_SE",
		"1055" : "tr_TR",
		"1058" : "uk_UA",
		"1066" : "vi_VN",
		"2067" : "nl_NL" // nl_BE
	};

	_port.onmessage = function(message) {
		_worker.oncommand && _worker.oncommand(message.data);
	};
	_port.postMessage({ "type" : "init", "dictionaries_path" : dictionariesPath, "languages" : this.languages });

	this.stop = function()
	{
		this.worker.terminate();
		this.worker = null;
	};

	this.command = function(message)
	{
		_port && _port.postMessage(message);
	};
	this.oncommand = function(message) { console.log(message); };

	this.checkDictionary = function(lang) {
		return (undefined !== this.languages["" + lang]) ? true : false;
	};
	this.getLanguages = function() {
		var ret = [];
		for (var lang in this.languages)
			ret.push(lang);
		return ret;
	};
}
