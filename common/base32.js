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

(function(window, document){

    function ZBase32Encoder()
    {
        this.EncodingTable = "ybndrfg8ejkmcpqxot1uwisza345h769";
        this.DecodingTable = ("undefined" == typeof Uint8Array) ? new Array(128) : new Uint8Array(128);

        var ii =  0;
        for (ii = 0; ii < 128; ii++)
            this.DecodingTable[ii] = 255;

        var _len_32 = this.EncodingTable.length;
        for (ii = 0; ii < _len_32; ii++)
        {
            this.DecodingTable[this.EncodingTable.charCodeAt(ii)] = ii;
        }

        this.GetUTF16_fromUnicodeChar = function(code)
        {
            if (code < 0x10000)
                return String.fromCharCode(code);
            else
            {
                code -= 0x10000;
                return String.fromCharCode(0xD800 | ((code >> 10) & 0x03FF)) + String.fromCharCode(0xDC00 | (code & 0x03FF));
            }
        };

        this.GetUTF16_fromUTF8 = function(pBuffer)
        {
            var _res = "";

            var lIndex = 0;
            var lCount = pBuffer.length;
            var val = 0;
            while (lIndex < lCount)
            {
                var byteMain = pBuffer[lIndex];
                if (0x00 == (byteMain & 0x80))
                {
                    // 1 byte
                    _res += this.GetUTF16_fromUnicodeChar(byteMain);
                    ++lIndex;
                }
                else if (0x00 == (byteMain & 0x20))
                {
                    // 2 byte
                    val = (((byteMain & 0x1F) << 6) |
                                (pBuffer[lIndex + 1] & 0x3F));
                    _res += this.GetUTF16_fromUnicodeChar(val);
                    lIndex += 2;
                }
                else if (0x00 == (byteMain & 0x10))
                {
                    // 3 byte
                    val = (((byteMain & 0x0F) << 12) |
                                ((pBuffer[lIndex + 1] & 0x3F) << 6) |
                                (pBuffer[lIndex + 2] & 0x3F));

                    _res += this.GetUTF16_fromUnicodeChar(val);
                    lIndex += 3;
                }
                else if (0x00 == (byteMain & 0x08))
                {
                    // 4 byte
                    val = (((byteMain & 0x07) << 18) |
                                ((pBuffer[lIndex + 1] & 0x3F) << 12) |
                                ((pBuffer[lIndex + 2] & 0x3F) << 6) |
                                (pBuffer[lIndex + 3] & 0x3F));

                    _res += this.GetUTF16_fromUnicodeChar(val);
                    lIndex += 4;
                }
                else if (0x00 == (byteMain & 0x04))
                {
                    // 5 byte
                    val = (((byteMain & 0x03) << 24) |
                                ((pBuffer[lIndex + 1] & 0x3F) << 18) |
                                ((pBuffer[lIndex + 2] & 0x3F) << 12) |
                                ((pBuffer[lIndex + 3] & 0x3F) << 6) |
                                (pBuffer[lIndex + 4] & 0x3F));

                    _res += this.GetUTF16_fromUnicodeChar(val);
                    lIndex += 5;
                }
                else
                {
                    // 6 byte
                    val = (((byteMain & 0x01) << 30) |
                                ((pBuffer[lIndex + 1] & 0x3F) << 24) |
                                ((pBuffer[lIndex + 2] & 0x3F) << 18) |
                                ((pBuffer[lIndex + 3] & 0x3F) << 12) |
                                ((pBuffer[lIndex + 4] & 0x3F) << 6) |
                                (pBuffer[lIndex + 5] & 0x3F));

                    _res += this.GetUTF16_fromUnicodeChar(val);
                    lIndex += 5;
                }
            }

            return _res;
        };

        this.GetUTF8_fromUTF16 = function(sData)
        {
            var pCur = 0;
            var pEnd = sData.length;

            var result = [];
            while (pCur < pEnd)
            {
                var code = sData.charCodeAt(pCur++);
                if (code >= 0xD800 && code <= 0xDFFF && pCur < pEnd)
                {
                    code = 0x10000 + (((code & 0x3FF) << 10) | (0x03FF & sData.charCodeAt(pCur++)));
                }

                if (code < 0x80)
                {
                    result.push(code);
                }
                else if (code < 0x0800)
                {
                    result.push(0xC0 | (code >> 6));
                    result.push(0x80 | (code & 0x3F));
                }
                else if (code < 0x10000)
                {
                    result.push(0xE0 | (code >> 12));
                    result.push(0x80 | ((code >> 6) & 0x3F));
                    result.push(0x80 | (code & 0x3F));
                }
                else if (code < 0x1FFFFF)
                {
                    result.push(0xF0 | (code >> 18));
                    result.push(0x80 | ((code >> 12) & 0x3F));
                    result.push(0x80 | ((code >> 6) & 0x3F));
                    result.push(0x80 | (code & 0x3F));
                }
                else if (code < 0x3FFFFFF)
                {
                    result.push(0xF8 | (code >> 24));
                    result.push(0x80 | ((code >> 18) & 0x3F));
                    result.push(0x80 | ((code >> 12) & 0x3F));
                    result.push(0x80 | ((code >> 6) & 0x3F));
                    result.push(0x80 | (code & 0x3F));
                }
                else if (code < 0x7FFFFFFF)
                {
                    result.push(0xFC | (code >> 30));
                    result.push(0x80 | ((code >> 24) & 0x3F));
                    result.push(0x80 | ((code >> 18) & 0x3F));
                    result.push(0x80 | ((code >> 12) & 0x3F));
                    result.push(0x80 | ((code >> 6) & 0x3F));
                    result.push(0x80 | (code & 0x3F));
                }
            }

            return result;
        };

        this.Encode = function(sData)
        {
            var data = this.GetUTF8_fromUTF16(sData);

            var encodedResult = "";
            var len = data.length;
            for (var i = 0; i < len; i += 5)
            {
                var byteCount = Math.min(5, len - i);

                var buffer = 0;
                for (var j = 0; j < byteCount; ++j)
                {
                    buffer *= 256;
                    buffer += data[i + j];
                }

                var bitCount = byteCount * 8;
                while (bitCount > 0)
                {
                    var index = 0;
                    if (bitCount >= 5)
                    {
                        var _del = Math.pow(2, bitCount - 5);
                        //var _del = 1 << (bitCount - 5);
                        index = (buffer / _del) & 0x1f;
                    }
                    else
                    {
                        index = (buffer & (0x1f >> (5 - bitCount)));
                        index <<= (5 - bitCount);
                    }

                    encodedResult += this.EncodingTable.charAt(index);
                    bitCount -= 5;
                }
            }

            return encodedResult;
        };

        this.Decode = function(data)
        {
            var result = [];

            var _len = data.length;
            var obj = { data: data, index : new Array(8) };

            var cur = 0;
            while (cur < _len)
            {
                cur = this.CreateIndexByOctetAndMovePosition(obj, cur);

                var shortByteCount = 0;
                var buffer = 0;
                for (var j = 0; j < 8 && obj.index[j] != -1; ++j)
                {
                    buffer *= 32;
                    buffer += (this.DecodingTable[obj.index[j]] & 0x1f);
                    shortByteCount++;
                }

                var bitCount = shortByteCount * 5;
                while (bitCount >= 8)
                {
                    //var _del = 1 << (bitCount - 8);
                    var _del = Math.pow(2, bitCount - 8);
                    var _res = (buffer / _del) & 0xff;
                    result.push(_res);
                    bitCount -= 8;
                }
            }

            this.GetUTF16_fromUTF8(result);
        };

        this.CreateIndexByOctetAndMovePosition = function(obj, currentPosition)
        {
            var j = 0;
            while (j < 8)
            {
                if (currentPosition >= obj.data.length)
                {
                    obj.index[j++] = -1;
                    continue;
                }

                if (this.IgnoredSymbol(obj.data.charCodeAt(currentPosition)))
                {
                    currentPosition++;
                    continue;
                }

                obj.index[j] = obj.data[currentPosition];
                j++;
                currentPosition++;
            }

            return currentPosition;
        };

        this.IgnoredSymbol = function(checkedSymbol)
        {
            return (checkedSymbol >= 128 || this.DecodingTable[checkedSymbol] == 255);
        };
    }

    AscCommon.ZBase32Encoder = ZBase32Encoder;

})(window, window.document);
