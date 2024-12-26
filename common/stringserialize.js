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

(function(window, undefined){

    window["AscCommon"] = window.AscCommon = (window["AscCommon"] || {});

    var charA = "A".charCodeAt(0);
    var charZ = "Z".charCodeAt(0);
    var chara = "a".charCodeAt(0);
    var charz = "z".charCodeAt(0);
    var char0 = "0".charCodeAt(0);
    var char9 = "9".charCodeAt(0);
    var charp = "+".charCodeAt(0);
    var chars = "/".charCodeAt(0);
    var char_break = ";".charCodeAt(0);

    function decodeBase64Char(ch)
    {
        if (ch >= charA && ch <= charZ)
            return ch - charA + 0;
        if (ch >= chara && ch <= charz)
            return ch - chara + 26;
        if (ch >= char0 && ch <= char9)
            return ch - char0 + 52;
        if (ch == charp)
            return 62;
        if (ch == chars)
            return 63;
        return -1;
    }

    var stringBase64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    var arrayBase64  = [];
    for (var index64 = 0; index64 < stringBase64.length; index64++)
    {
        arrayBase64.push(stringBase64.charAt(index64));
    }

    window.AscCommon["Base64"] = window.AscCommon.Base64 = {};

    /**
     * Decode input base64 data to output array
     * @memberof AscCommon.Base64
     * @alias decodeData
     * @param {string|Array|TypedArray} input input data
     * @param {number} [input_offset = undefined] offset in input data. 0 by default
     * @param {number} [input_len = undefined] length input data (not length of needed data, this value does not depend on the offset. input.length by default
     * @param {Array|TypedArray} output output data
     * @param {number} [output_offset = undefined] output data offset. 0 by default
     * @return {number} offset in output data (output_offset + count_write_bytes)
     */
    window.AscCommon.Base64.decodeData = window.AscCommon.Base64["decodeData"] = function(input, input_offset, input_len, output, output_offset)
    {
        var isBase64 = typeof input === "string";
        if (undefined === input_len) input_len = input.length;
        var writeIndex = (undefined === output_offset) ? 0 : output_offset;
        var index = (undefined === input_offset) ? 0 : input_offset;

        while (index < input_len)
        {
            var dwCurr = 0;
            var i;
            var nBits = 0;
            for (i=0; i<4; i++)
            {
                if (index >= input_len)
                    break;
                var nCh = decodeBase64Char(isBase64 ? input.charCodeAt(index) : input[index]);
                index++;
                if (nCh == -1)
                {
                    i--;
                    continue;
                }
                dwCurr <<= 6;
                dwCurr |= nCh;
                nBits += 6;
            }

            dwCurr <<= 24-nBits;
            for (i=0; i<(nBits>>3); i++)
            {
                output[writeIndex++] = ((dwCurr & 0x00ff0000) >>> 16);
                dwCurr <<= 8;
            }
        }
        return writeIndex;
    };

    /**
     * Decode input base64 data to returned Uint8Array
     * @memberof AscCommon.Base64
     * @alias decode
     * @param {string|Array|TypedArray} input input data
     * @param {boolean} [isUsePrefix = undefined] is detect destination size by prefix. false by default
     * @param {number} [dstlen = undefined] destination length
     * @param {number} [offset] offset of input data
     * @return {Uint8Array} decoded data
     */
    window.AscCommon.Base64.decode = window.AscCommon.Base64["decode"] = function(input, isUsePrefix, dstlen, offset)
    {
        var srcLen = input.length;
        var index = (undefined === offset) ? 0 : offset;
        var dstLen = (undefined === dstlen) ? srcLen : dstlen;

        var isBase64 = typeof input === "string";

        if (isUsePrefix && isBase64)
        {
            // ищем длину
            dstLen = 0;
            var maxLen = Math.max(11, srcLen); // > 4 Gb
            while (index < maxLen)
            {
                var c = input.charCodeAt(index++);
                if (c == char_break)
                    break;

                dstLen *= 10;
                dstLen += (c - char0);
            }

            if (index == maxLen)
            {
                // длины нет
                index = 0;
                dstLen = srcLen;
            }
        }

        var dst = new Uint8Array(dstLen);
        var writeIndex = window.AscCommon.Base64.decodeData(input, index, srcLen, dst, 0);

        if (writeIndex == dstLen)
            return dst;

        return new Uint8Array(dst.buffer, 0, writeIndex);
    };

    /**
     * Encode input data to base64 string
     * @memberof AscCommon.Base64
     * @alias encode
     * @param {Array|TypedArray} input input data
     * @param {number} [offset = undefined] offset of input data. 0 by default
     * @param {number} [length = undefined] length input data (last index: offset + length). input.length by default
     * @param {boolean} [isUsePrefix = undefined] is add destination size by prefix. false by default
     * @return {string} encoded data
     */
    window.AscCommon.Base64.encode = window.AscCommon.Base64["encode"] = function(input, offset, length, isUsePrefix)
    {
        var srcLen = (undefined === length) ? input.length : length;
        var index = (undefined === offset) ? 0 : offset;

        var len1 = (((srcLen / 3) >> 0) * 4);
        var len2 = (len1 / 76) >> 0;
        var len3 = 19;
        var dstArray = [];

        var sTemp = "";
        var dwCurr = 0;
        for (var i = 0; i <= len2; i++)
        {
            if (i == len2)
                len3 = ((len1 % 76) / 4) >> 0;

            for (var j = 0; j < len3; j++)
            {
                dwCurr = 0;
                for (var n = 0; n < 3; n++)
                {
                    dwCurr |= input[index++];
                    dwCurr <<= 8;
                }

                sTemp = "";
                for (var k = 0; k < 4; k++)
                {
                    var b = (dwCurr >>> 26) & 0xFF;
                    sTemp += arrayBase64[b];
                    dwCurr <<= 6;
                    dwCurr &= 0xFFFFFFFF;
                }
                dstArray.push(sTemp);
            }
        }
        len2 = (srcLen % 3 != 0) ? (srcLen % 3 + 1) : 0;
        if (len2)
        {
            dwCurr = 0;
            for (var n = 0; n < 3; n++)
            {
                if (n < (srcLen % 3))
                    dwCurr |= input[index++];
                dwCurr <<= 8;
            }

            sTemp = "";
            for (var k = 0; k < len2; k++)
            {
                var b = (dwCurr >>> 26) & 0xFF;
                sTemp += arrayBase64[b];
                dwCurr <<= 6;
            }

            len3 = (len2 != 0) ? 4 - len2 : 0;
            for (var j = 0; j < len3; j++)
            {
                sTemp += '=';
            }
            dstArray.push(sTemp);
        }

        return isUsePrefix ? (("" + srcLen + ";") + dstArray.join("")) : dstArray.join("");
    };

    window.AscCommon["Hex"] = window.AscCommon.Hex = {};

    /**
     * Decode input hex data to Uint8Array
     * @memberof AscCommon.Hex
     * @alias decode
     * @param {string} input input data
     * @return {Uint8Array} decoded data
     */
    window.AscCommon.Hex.decode = window.AscCommon.Hex["decode"] = function(input)
    {
        var hexToByte = function(c) {
            if (c >= 48 && c <= 57) return c - 48; // 0..9
            if (c >= 97 && c <= 102) return c - 87;
            if (c >= 65 && c <= 70) return c - 55;
            return 0;
        };

        var len = input.length;
        if (len & 0x01) len -= 1;
        var result = new Uint8Array(len >> 1);
        var resIndex = 0;
        for (var i = 0; i < len; i += 2)
        {
            result[resIndex++] = hexToByte(input.charCodeAt(i)) << 4 | hexToByte(input.charCodeAt(i + 1));
        }
        return result;
    };

    /**
     * Encode Uint8Array to hex string
     * @memberof AscCommon.Hex
     * @alias encode
     * @param {Array|TypedArray} input input data
     * @param {boolean} [isUpperCase = false] is use upper case
     * @return {string} encoded data
     */
    window.AscCommon.Hex.encode = window.AscCommon.Hex["encode"] = function(input, isUpperCase)
    {
        var byteToHex = new Array(256);
        for (var i = 0; i < 16; i++)
            byteToHex[i] = "0" + (isUpperCase ? i.toString(16).toUpperCase() : i.toString(16));
        for (var i = 16; i < 256; i++)
            byteToHex[i] = isUpperCase ? i.toString(16).toUpperCase() : i.toString(16);

        var result = "";
        for (var i = 0, len = input.length; i < len; i++)
            result += byteToHex[input[i]];

        return result;
    };

    window.AscCommon["Base58"] = window.AscCommon.Base58 = {};

    /**
     * Encode data to base58 string
     * @memberof AscCommon.Base58
     * @alias encode
     * @param {Array|TypedArray|string} input input data
     * @return {string} encoded data
     */
    window.AscCommon.Base58.encode = function(buf)
    {
        if(typeof buf === "string")
        {
            let old = buf;
            buf = [];
            for (let i = 0, len = old.length; i < len; i++)
                buf.push(old.charCodeAt(i));
        }

        const chars = "123456789ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz";
        const chars_map = [
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
            -1,  0,  1,  2,  3,  4,  5,  6,  7,  8, -1, -1, -1, -1, -1, -1,
            -1,  9, 10, 11, 12, 13, 14, 15, 16, -1, 17, 18, 19, 20, 21, -1,
            22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, -1, -1, -1, -1, -1,
            -1, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, -1, 44, 45, 46,
            47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, -1, -1, -1, -1, -1,
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
            -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1
        ];

        let result = [];
        for (let i = 0, len = buf.length; i < len; i++)
        {
            let carry = buf[i];
            for (let j = 0; j < result.length; ++j)
            {
                const x = (chars_map[result[j]] << 8) + carry;
                result[j] = chars.charCodeAt(x % 58);
                carry = (x / 58) >> 0;
            }
            while (carry)
            {
                result.push(chars.charCodeAt(carry % 58));
                carry = (carry / 58) >> 0;
            }
        }

        let char1 = "1".charCodeAt(0);
        for (let i = 0, len = buf.length; i < len; i++)
        {
            if (buf[i])
                break;
            else
                result.push(char1);
        }

        result.reverse();
        return String.fromCharCode.apply(null, result);
    };

    window.AscCommon["Utf8"] = window.AscCommon.Utf8 = {};

    window.AscCommon.Utf8.encode = function(text, isEndNull)
    {
        //return new (TextEncoder()).encode(password);

        var inputLen = text.length;
        var testLen  = 6 * inputLen + 1;
        var tmpStrings = new ArrayBuffer(testLen);

        var code  = 0;
        var index = 0;

        var outputIndex = 0;
        var outputDataTmp = new Uint8Array(tmpStrings);
        var outputData = outputDataTmp;

        while (index < inputLen)
        {
            code = text.charCodeAt(index++);
            if (code >= 0xD800 && code <= 0xDFFF && index < inputLen)
                code = 0x10000 + (((code & 0x3FF) << 10) | (0x03FF & text.charCodeAt(index++)));

            if (code < 0x80)
                outputData[outputIndex++] = code;
            else if (code < 0x0800)
            {
                outputData[outputIndex++] = 0xC0 | (code >> 6);
                outputData[outputIndex++] = 0x80 | (code & 0x3F);
            }
            else if (code < 0x10000)
            {
                outputData[outputIndex++] = 0xE0 | (code >> 12);
                outputData[outputIndex++] = 0x80 | ((code >> 6) & 0x3F);
                outputData[outputIndex++] = 0x80 | (code & 0x3F);
            }
            else if (code < 0x1FFFFF)
            {
                outputData[outputIndex++] = 0xF0 | (code >> 18);
                outputData[outputIndex++] = 0x80 | ((code >> 12) & 0x3F);
                outputData[outputIndex++] = 0x80 | ((code >> 6) & 0x3F);
                outputData[outputIndex++] = 0x80 | (code & 0x3F);
            }
            else if (code < 0x3FFFFFF)
            {
                outputData[outputIndex++] = 0xF8 | (code >> 24);
                outputData[outputIndex++] = 0x80 | ((code >> 18) & 0x3F);
                outputData[outputIndex++] = 0x80 | ((code >> 12) & 0x3F);
                outputData[outputIndex++] = 0x80 | ((code >> 6) & 0x3F);
                outputData[outputIndex++] = 0x80 | (code & 0x3F);
            }
            else if (code < 0x7FFFFFFF)
            {
                outputData[outputIndex++] = 0xFC | (code >> 30);
                outputData[outputIndex++] = 0x80 | ((code >> 24) & 0x3F);
                outputData[outputIndex++] = 0x80 | ((code >> 18) & 0x3F);
                outputData[outputIndex++] = 0x80 | ((code >> 12) & 0x3F);
                outputData[outputIndex++] = 0x80 | ((code >> 6) & 0x3F);
                outputData[outputIndex++] = 0x80 | (code & 0x3F);
            }
        }

        if (isEndNull === true)
            outputData[outputIndex++] = 0;

        return new Uint8Array(tmpStrings, 0, outputIndex);
    };


})(window);
