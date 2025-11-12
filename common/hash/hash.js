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

(function(window, undefined) {

	function CHashWorker(message, callback)
	{
		this.message = message;
		this.callback = callback;

		this.useWasm = false;
		var webAsmObj = window["WebAssembly"];
		if (typeof webAsmObj === "object")
		{
			if (typeof webAsmObj["Memory"] === "function")
			{
				if ((typeof webAsmObj["instantiateStreaming"] === "function") || (typeof webAsmObj["instantiate"] === "function"))
					this.useWasm = true;
			}
		}

		this.enginePath = "./../../../../sdkjs/common/hash/hash/";
		this.worker = null;

		this.stop = function()
		{
			if (!this.worker)
				return;

			try
			{
				if (this.worker.port)
					this.worker.port.close();
				else if (this.worker.terminate)
					this.worker.terminate();
			}
			catch (err)
			{
			}

			this.worker = null;
		};

		this.restart = function()
		{
			this.stop();

			var worker_src = this.useWasm ? "engine.js" : "engine_ie.js";
			worker_src = this.enginePath + worker_src;

			this.worker = new Worker(worker_src);

			var _t = this;

			// обрабатываем ошибку, чтобы он не влиял на работу редактора
			// и если ошибка из wasm модуля - то просто попробуем js версию - и рестартанем
			this.worker.onerror = function(e) {

				e && e.preventDefault && e.preventDefault();
				e && e.stopPropagation && e.stopPropagation();

				if (_t.useWasm)
				{
					_t.useWasm = false;
					_t.restart();
				}
			};

			this._start(this.worker);
		};

		this.oncommand = function(message)
		{
			if (message && message["hashValue"] !== undefined)
			{
				if (this.callback)
					this.callback.call(window, message["hashValue"]);
				this.stop();
			}
		};

		this._start = function(_port)
		{
			var _worker = this;

			_port.onmessage = function(message) {
				_worker.oncommand && _worker.oncommand(message.data);
			};
			_port.postMessage({ "type" : "hash", "value" : this.message });
		};

		this.restart();
	}

	var currentHashWorker = null;

	window['AscCommon'] = window['AscCommon'] || {};
	var AscCommon = window['AscCommon'];

	window['AscCommon'].calculateProtectHash = function(args, callback) {
		if (window["NATIVE_EDITOR_ENJINE"])
		{
			if (!AscCommon.hashEngine) {
				/** @suppress {checkVars} */
				AscCommon.hashEngine = CreateEmbedObject("CHashEmbed");
			}

			let retArray = [];
			for (var i = 0, len = args.length; i < len; i++)
			{
				retArray.push(AscCommon.Base64.encode(AscCommon.hashEngine["hash2"](args[i].password, args[i].salt, args[i].spinCount, args[i].alg)));
			}

			callback(retArray);
			return;
		}

		var sendedData = [];
		for (var i = 0, len = args.length; i < len; i++)
		{
			sendedData.push({
				"password" : args[i].password,
				"salt" : args[i].salt,
				"spinCount" : args[i].spinCount,
				"alg" : args[i].alg
			});
		}

		currentHashWorker = new CHashWorker(sendedData, callback);
	};

	window['AscCommon'].prepareWordPassword = function(password) {

		let _highOrderWords = [[0xE1, 0xF0], [0x1D, 0x0F], [0xCC, 0x9C], [0x84, 0xC0], [0x11, 0x0C], [0x0E, 0x10], [0xF1, 0xCE], [0x31, 0x3E], [0x18, 0x72], [0xE1, 0x39],
			[0xD4, 0x0F], [0x84, 0xF9], [0x28, 0x0C], [0xA9, 0x6A], [0x4E, 0xC3]];

		let _encryptionMatrix = [[[0xAE, 0xFC], [0x4D, 0xD9], [0x9B, 0xB2], [0x27, 0x45], [0x4E, 0x8A], [0x9D, 0x14], [0x2A, 0x09]],
			[[0x7B, 0x61], [0xF6, 0xC2], [0xFD, 0xA5], [0xEB, 0x6B], [0xC6, 0xF7], [0x9D, 0xCF], [0x2B, 0xBF]],
			[[0x45, 0x63], [0x8A, 0xC6], [0x05, 0xAD], [0x0B, 0x5A], [0x16, 0xB4], [0x2D, 0x68], [0x5A, 0xD0]],
			[[0x03, 0x75], [0x06, 0xEA], [0x0D, 0xD4], [0x1B, 0xA8], [0x37, 0x50], [0x6E, 0xA0], [0xDD, 0x40]],
			[[0xD8, 0x49], [0xA0, 0xB3], [0x51, 0x47], [0xA2, 0x8E], [0x55, 0x3D], [0xAA, 0x7A], [0x44, 0xD5]],
			[[0x6F, 0x45], [0xDE, 0x8A], [0xAD, 0x35], [0x4A, 0x4B], [0x94, 0x96], [0x39, 0x0D], [0x72, 0x1A]],
			[[0xEB, 0x23], [0xC6, 0x67], [0x9C, 0xEF], [0x29, 0xFF], [0x53, 0xFE], [0xA7, 0xFC], [0x5F, 0xD9]],
			[[0x47, 0xD3], [0x8F, 0xA6], [0x0F, 0x6D], [0x1E, 0xDA], [0x3D, 0xB4], [0x7B, 0x68], [0xF6, 0xD0]],
			[[0xB8, 0x61], [0x60, 0xE3], [0xC1, 0xC6], [0x93, 0xAD], [0x37, 0x7B], [0x6E, 0xF6], [0xDD, 0xEC]],
			[[0x45, 0xA0], [0x8B, 0x40], [0x06, 0xA1], [0x0D, 0x42], [0x1A, 0x84], [0x35, 0x08], [0x6A, 0x10]],
			[[0xAA, 0x51], [0x44, 0x83], [0x89, 0x06], [0x02, 0x2D], [0x04, 0x5A], [0x08, 0xB4], [0x11, 0x68]],
			[[0x76, 0xB4], [0xED, 0x68], [0xCA, 0xF1], [0x85, 0xC3], [0x1B, 0xA7], [0x37, 0x4E], [0x6E, 0x9C]],
			[[0x37, 0x30], [0x6E, 0x60], [0xDC, 0xC0], [0xA9, 0xA1], [0x43, 0x63], [0x86, 0xC6], [0x1D, 0xAD]],
			[[0x33, 0x31], [0x66, 0x62], [0xCC, 0xC4], [0x89, 0xA9], [0x03, 0x73], [0x06, 0xE6], [0x0D, 0xCC]],
			[[0x10, 0x21], [0x20, 0x42], [0x40, 0x84], [0x81, 0x08], [0x12, 0x31], [0x24, 0x62], [0x48, 0xC4]]];

		let byteToBits = function (byte) {
			let bitsArr = [];
			for (let i = 0; i < 8; i++) {
				let bit = byte & (1 << i) ? 1 : 0;
				bitsArr.push(bit);
			}
			return bitsArr;
		};

		let splitByHighLowBytes = function (val) {
			return [(val & 0xff00) >> 8, val & 0x00ff];
		};

		let getHexStr = function (arr) {
			let res = "";
			for (let i = 0; i < arr.length; i++) {
				res += AscCommon.ByteToHex(arr[i]);
			}
			return res;
		};

		let getUnionBytes = function (byte1, byte2) {
			return ((byte1 << 8) | byte2);
		};

		if (password == null) {
			return null;
		}

		let passwordBytes = AscCommon.Utf8.encode(password);

		let maxPasswordLength = 15;
		passwordBytes = passwordBytes.slice(0, maxPasswordLength);
		let passwordLength = passwordBytes.length;

		//high word
		let highOrderWord = [0x00, 0x00];
		if (passwordLength > 0) {
			highOrderWord = _highOrderWords[passwordLength - 1];
		}
		for (var i = 0; i < passwordLength; i++) {
			let passwordByte = passwordBytes[i];
			let encryptionMatrixIndex = i + (_encryptionMatrix.length - passwordLength);

			let bitArray = byteToBits(passwordByte);

			for (let j = 0; j < _encryptionMatrix[0].length; j++) {
				let isSet = bitArray[j];

				if (isSet) {
					for (let k = 0; k < _encryptionMatrix[0][0].length; k++) {
						highOrderWord[k] = highOrderWord[k] ^ _encryptionMatrix[encryptionMatrixIndex][j][k];
					}
				}
			}
		}

		if (!highOrderWord) {
			return null;
		}

		//low word
		let lowOrderSequence = 0x00;
		let sequence1 = 0x01;
		let sequence7FFF = 0x7FFF;
		for (let i = passwordLength - 1; i >= 0; i--) {
			let passwordByte = passwordBytes[i];
			lowOrderSequence = (((lowOrderSequence >> 14) & sequence1) | ((lowOrderSequence << 1) & sequence7FFF)) ^ passwordByte;
		}
		lowOrderSequence = (((lowOrderSequence >> 14) & sequence1) | ((lowOrderSequence << 1) & sequence7FFF)) ^ passwordLength ^ getUnionBytes(0xCE, 0x4B);

		let lowOrderWord = splitByHighLowBytes(lowOrderSequence);
		if (!lowOrderWord) {
			return null;
		}

		//full word
		let fullWord = highOrderWord.concat(lowOrderWord);
		fullWord = fullWord && fullWord.reverse();
		if (!fullWord) {
			return null;
		}

		return getHexStr(fullWord);
	};

	window['AscCommon'].HashAlgs = {
		MD2       : 0,
		MD4       : 1,
		MD5       : 2,
		RMD160    : 3,
		SHA1      : 4,
		SHA256    : 5,
		SHA384    : 6,
		SHA512    : 7,
		WHIRLPOOL : 8
	};

})(window);
