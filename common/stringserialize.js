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
    window.AscCommon.Base64.decodeData = window.AscCommon.Base64["decodeData"] = function(input, input_offset, input_len, output, output_offset)
    {
        var isBase64 = typeof input === "string";
        var writeIndex = (undefined === output_offset) ? 0 : output_offset;
        var index = (undefined === input_offset) ? 0 : input_offset;

        while (index < input_len)
        {
            var dwCurr = 0;
            var i;
            var nBits = 0;
            for (i=0; i<4; i++)
            {
                if (index >= srcLen)
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


})(window);