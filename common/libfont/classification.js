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

// Import
var fontslot_ASCII    = 0x00;
var fontslot_EastAsia = 0x01;
var fontslot_CS       = 0x02;
var fontslot_HAnsi    = 0x03;

var fonthint_Default  = 0x00;
var fonthint_CS       = 0x01;
var fonthint_EastAsia = 0x02;

(function(){
    function CDetectFontUse()
    {
        this.DetectData     = null;

        this.TableChunkLen  = 65536;
        this.TableChunks    = 4;

        this.TableChunkMain     = 0;
        this.TableChunkHintEA   = this.TableChunkLen;
        this.TableChunkHintZH   = 2 * this.TableChunkLen;
        this.TableChunkHintEACS = 3 * this.TableChunkLen;

        this.Init = function()
        {
            this.DetectData = AscFonts.allocate(this.TableChunkLen * this.TableChunks);
            var _data = this.DetectData;
            var i, j;

            // ********************** 1st table *********************** //
            j = 0;
            for (i = 0x0000; i <= 0x007F; i++)
            {
                _data[i + j] = fontslot_ASCII;
            }
            for (i = 0x00A0; i <= 0x04FF; i++)
            {
                _data[i + j] = fontslot_HAnsi;
            }
            for (i = 0x0590; i <= 0x07BF; i++)
            {
                _data[i + j] = fontslot_ASCII;
            }
            for (i = 0x1100; i <= 0x11FF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x1E00; i <= 0x1EFF; i++)
            {
                _data[i + j] = fontslot_HAnsi;
            }
            for (i = 0x1F00; i <= 0x27BF; i++)
            {
                _data[i + j] = fontslot_HAnsi;
            }
            for (i = 0x2E80; i <= 0x319F; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x3200; i <= 0x4D8F; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x4E00; i <= 0x9FAF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xA000; i <= 0xA4CF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xAC00; i <= 0xD7AF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xD800; i <= 0xDFFF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xE000; i <= 0xF8FF; i++)
            {
                _data[i + j] = fontslot_HAnsi;
            }
            for (i = 0xF900; i <= 0xFAFF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xFB00; i <= 0xFB1C; i++)
            {
                _data[i + j] = fontslot_HAnsi;
            }
            for (i = 0xFB1D; i <= 0xFDFF; i++)
            {
                _data[i + j] = fontslot_ASCII;
            }
            for (i = 0xFE30; i <= 0xFE6F; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xFE70; i <= 0xFEFE; i++)
            {
                _data[i + j] = fontslot_ASCII;
            }
            for (i = 0xFF00; i <= 0xFFEF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }

            // ********************** 2nd table *********************** //
            j = this.TableChunkHintEA;
            for (i = 0x0000; i <= 0x007F; i++)
            {
                _data[i + j] = fontslot_ASCII;
            }
            for (i = 0x00A0; i <= 0x04FF; i++)
            {
                _data[i + j] = fontslot_HAnsi;
            }
            _data[0xA1 + j] = fontslot_EastAsia;
            _data[0xA4 + j] = fontslot_EastAsia;
            _data[0xA7 + j] = fontslot_EastAsia;
            _data[0xA8 + j] = fontslot_EastAsia;
            _data[0xAA + j] = fontslot_EastAsia;
            _data[0xAD + j] = fontslot_EastAsia;
            _data[0xAF + j] = fontslot_EastAsia;
            _data[0xB0 + j] = fontslot_EastAsia;
            _data[0xB1 + j] = fontslot_EastAsia;
            _data[0xB2 + j] = fontslot_EastAsia;
            _data[0xB3 + j] = fontslot_EastAsia;
            _data[0xB4 + j] = fontslot_EastAsia;
            _data[0xB6 + j] = fontslot_EastAsia;
            _data[0xB7 + j] = fontslot_EastAsia;
            _data[0xB8 + j] = fontslot_EastAsia;
            _data[0xB9 + j] = fontslot_EastAsia;
            _data[0xBA + j] = fontslot_EastAsia;
            _data[0xBC + j] = fontslot_EastAsia;
            _data[0xBD + j] = fontslot_EastAsia;
            _data[0xBE + j] = fontslot_EastAsia;
            _data[0xBF + j] = fontslot_EastAsia;
            _data[0xD7 + j] = fontslot_EastAsia;
            _data[0xF7 + j] = fontslot_EastAsia;

            for (i = 0x02B0; i <= 0x04FF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x0590; i <= 0x07BF; i++)
            {
                _data[i + j] = fontslot_ASCII;
            }
            for (i = 0x1100; i <= 0x11FF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x1E00; i <= 0x1EFF; i++)
            {
                _data[i + j] = fontslot_HAnsi;
            }
            for (i = 0x1F00; i <= 0x1FFF; i++)
            {
                _data[i + j] = fontslot_HAnsi;
            }
            for (i = 0x2000; i <= 0x27BF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x2E80; i <= 0x319F; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x3200; i <= 0x4D8F; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x4E00; i <= 0x9FAF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xA000; i <= 0xA4CF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xAC00; i <= 0xD7AF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xD800; i <= 0xDFFF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xE000; i <= 0xF8FF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xF900; i <= 0xFAFF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xFB00; i <= 0xFB1C; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xFB1D; i <= 0xFDFF; i++)
            {
                _data[i + j] = fontslot_ASCII;
            }
            for (i = 0xFE30; i <= 0xFE6F; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xFE70; i <= 0xFEFE; i++)
            {
                _data[i + j] = fontslot_ASCII;
            }
            for (i = 0xFF00; i <= 0xFFEF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }

            // ********************** 3rd table *********************** //
            j = this.TableChunkHintZH;
            for (i = 0x0000; i <= 0x007F; i++)
            {
                _data[i + j] = fontslot_ASCII;
            }
            for (i = 0x00A0; i <= 0x00FF; i++)
            {
                _data[i + j] = fontslot_HAnsi;
            }
            _data[0xA1 + j] = fontslot_EastAsia;
            _data[0xA4 + j] = fontslot_EastAsia;
            _data[0xA7 + j] = fontslot_EastAsia;
            _data[0xA8 + j] = fontslot_EastAsia;
            _data[0xAA + j] = fontslot_EastAsia;
            _data[0xAD + j] = fontslot_EastAsia;
            _data[0xAF + j] = fontslot_EastAsia;
            _data[0xB0 + j] = fontslot_EastAsia;
            _data[0xB1 + j] = fontslot_EastAsia;
            _data[0xB2 + j] = fontslot_EastAsia;
            _data[0xB3 + j] = fontslot_EastAsia;
            _data[0xB4 + j] = fontslot_EastAsia;
            _data[0xB6 + j] = fontslot_EastAsia;
            _data[0xB7 + j] = fontslot_EastAsia;
            _data[0xB8 + j] = fontslot_EastAsia;
            _data[0xB9 + j] = fontslot_EastAsia;
            _data[0xBA + j] = fontslot_EastAsia;
            _data[0xBC + j] = fontslot_EastAsia;
            _data[0xBD + j] = fontslot_EastAsia;
            _data[0xBE + j] = fontslot_EastAsia;
            _data[0xBF + j] = fontslot_EastAsia;
            _data[0xD7 + j] = fontslot_EastAsia;
            _data[0xF7 + j] = fontslot_EastAsia;

            _data[0xE0 + j] = fontslot_EastAsia;
            _data[0xE1 + j] = fontslot_EastAsia;
            _data[0xE8 + j] = fontslot_EastAsia;
            _data[0xE9 + j] = fontslot_EastAsia;
            _data[0xEA + j] = fontslot_EastAsia;
            _data[0xEC + j] = fontslot_EastAsia;
            _data[0xED + j] = fontslot_EastAsia;
            _data[0xF2 + j] = fontslot_EastAsia;
            _data[0xF3 + j] = fontslot_EastAsia;
            _data[0xF9 + j] = fontslot_EastAsia;
            _data[0xFA + j] = fontslot_EastAsia;
            _data[0xFC + j] = fontslot_EastAsia;

            for (i = 0x0100; i <= 0x02AF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x02B0; i <= 0x04FF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x0590; i <= 0x07BF; i++)
            {
                _data[i + j] = fontslot_ASCII;
            }
            for (i = 0x1100; i <= 0x11FF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x1E00; i <= 0x1EFF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x1F00; i <= 0x1FFF; i++)
            {
                _data[i + j] = fontslot_HAnsi;
            }
            for (i = 0x2000; i <= 0x27BF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x2E80; i <= 0x319F; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x3200; i <= 0x4D8F; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x4E00; i <= 0x9FAF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xA000; i <= 0xA4CF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xAC00; i <= 0xD7AF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xD800; i <= 0xDFFF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xE000; i <= 0xF8FF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xF900; i <= 0xFAFF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xFB00; i <= 0xFB1C; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xFB1D; i <= 0xFDFF; i++)
            {
                _data[i + j] = fontslot_ASCII;
            }
            for (i = 0xFE30; i <= 0xFE6F; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xFE70; i <= 0xFEFE; i++)
            {
                _data[i + j] = fontslot_ASCII;
            }
            for (i = 0xFF00; i <= 0xFFEF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }

            // ********************** 4rd table *********************** //
            j = this.TableChunkHintEACS;
            for (i = 0x0000; i <= 0x007F; i++)
            {
                _data[i + j] = fontslot_ASCII;
            }
            for (i = 0x00A0; i <= 0x00FF; i++)
            {
                _data[i + j] = fontslot_HAnsi;
            }
            _data[0xA1 + j] = fontslot_EastAsia;
            _data[0xA4 + j] = fontslot_EastAsia;
            _data[0xA7 + j] = fontslot_EastAsia;
            _data[0xA8 + j] = fontslot_EastAsia;
            _data[0xAA + j] = fontslot_EastAsia;
            _data[0xAD + j] = fontslot_EastAsia;
            _data[0xAF + j] = fontslot_EastAsia;
            _data[0xB0 + j] = fontslot_EastAsia;
            _data[0xB1 + j] = fontslot_EastAsia;
            _data[0xB2 + j] = fontslot_EastAsia;
            _data[0xB3 + j] = fontslot_EastAsia;
            _data[0xB4 + j] = fontslot_EastAsia;
            _data[0xB6 + j] = fontslot_EastAsia;
            _data[0xB7 + j] = fontslot_EastAsia;
            _data[0xB8 + j] = fontslot_EastAsia;
            _data[0xB9 + j] = fontslot_EastAsia;
            _data[0xBA + j] = fontslot_EastAsia;
            _data[0xBC + j] = fontslot_EastAsia;
            _data[0xBD + j] = fontslot_EastAsia;
            _data[0xBE + j] = fontslot_EastAsia;
            _data[0xBF + j] = fontslot_EastAsia;
            _data[0xD7 + j] = fontslot_EastAsia;
            _data[0xF7 + j] = fontslot_EastAsia;

            for (i = 0x0100; i <= 0x02AF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x02B0; i <= 0x04FF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x0590; i <= 0x07BF; i++)
            {
                _data[i + j] = fontslot_ASCII;
            }
            for (i = 0x1100; i <= 0x11FF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x1E00; i <= 0x1EFF; i++)
            {
                _data[i + j] = fontslot_HAnsi;
            }
            for (i = 0x1F00; i <= 0x1FFF; i++)
            {
                _data[i + j] = fontslot_HAnsi;
            }
            for (i = 0x2000; i <= 0x27BF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x2E80; i <= 0x319F; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x3200; i <= 0x4D8F; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0x4E00; i <= 0x9FAF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xA000; i <= 0xA4CF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xAC00; i <= 0xD7AF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xD800; i <= 0xDFFF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xE000; i <= 0xF8FF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xF900; i <= 0xFAFF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xFB00; i <= 0xFB1C; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xFB1D; i <= 0xFDFF; i++)
            {
                _data[i + j] = fontslot_ASCII;
            }
            for (i = 0xFE30; i <= 0xFE6F; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
            for (i = 0xFE70; i <= 0xFEFE; i++)
            {
                _data[i + j] = fontslot_ASCII;
            }
            for (i = 0xFF00; i <= 0xFFEF; i++)
            {
                _data[i + j] = fontslot_EastAsia;
            }
        }

        this.Get_FontClass = function(nUnicode, nHint, nEastAsia_lcid, bCS, bRTL)
        {
            var _glyph_slot = fontslot_ASCII;
			if (nUnicode > 0xFFFF)
			{
				if ((nUnicode >= 0x20000 && nUnicode <= 0x2A6DF) ||
					(nUnicode >= 0x2F800 && nUnicode <= 0x2FA1F))
				{
					_glyph_slot = fontslot_EastAsia;
				}
				else if (nUnicode >= 0x1D400 && nUnicode <= 0x1D7FF)
				{
					_glyph_slot = fontslot_ASCII;
				}
				else
				{
					_glyph_slot = fontslot_HAnsi;
				}
			}			
            else if (nHint != fonthint_EastAsia)
            {
                _glyph_slot = this.DetectData[nUnicode];
            }
            else
            {
                if (nEastAsia_lcid == lcid_zh)
                    _glyph_slot = this.DetectData[this.TableChunkHintZH + nUnicode];
                else
                    _glyph_slot = this.DetectData[this.TableChunkHintEA + nUnicode];

                if (_glyph_slot == fontslot_EastAsia)
                    return _glyph_slot;
            }

            if (bCS || bRTL)
                return fontslot_CS;

            return _glyph_slot;
        }
    }

    window.CDetectFontUse = CDetectFontUse;
	window.CDetectFontUse;
})();

var g_font_detector = new window.CDetectFontUse();
g_font_detector.Init();
