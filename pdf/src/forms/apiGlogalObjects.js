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

(function(){
    // please use copy of this object
    // let DEFAULT_SPAN = {
    //     "alignment":        ALIGN_TYPE.left,
    //     "fontFamily":       ["sans-serif"],
    //     "fontStretch":      "normal",
    //     "fontStyle":        "normal",
    //     "fontWeight":       400,
    //     "strikethrough":    false,
    //     "subscript":        false,
    //     "superscript":      false,
    //     "text":             "",
    //     "color":            color["black"],
    //     "textSize":         12.0,
    //     "underline":        false
    // }

    let border = {
        "s": "solid",
        "b": "beveled",
        "d": "dashed",
        "i": "inset",
        "u": "underline"
    }

    let position = {
        "textOnly":   0,
        "iconOnly":   1,
        "iconTextV":  2,
        "textIconV":  3,
        "iconTextH":  4,
        "textIconH":  5,
        "overlay":    6
    }

    let scaleHow = {
        "proportional":   0,
        "anamorphic":     1
    }

    let scaleWhen = {
        "always":     0,
        "never":      1,
        "tooBig":     2,
        "tooSmall":   3
    }

    let display = {
        "visible":  0,
        "hidden":   1,
        "noPrint":  2,
        "noView":   3
    }

    let style = {
        "ch": "check",
        "cr": "cross",
        "di": "diamond",
        "ci": "circle",
        "st": "star",
        "sq": "square"
    }

    let highlight = {
        "n": "none",
        "i": "invert",
        "p": "push",
        "o": "outline"
    }

    let color = {};
    Object.defineProperties(color, {
        "transparent":  {
            get: function() { return [ "T" ] }
        },
        "black": {
            get: function() { return [ "G", 0 ] }
        },
        "white": {
            get: function() { return [ "G", 1 ] }
        },
        "red": {
            get: function() { return [ "RGB", 1,0,0 ] }
        },  
        "green": {
            get: function() { return [ "RGB", 0,1,0 ] }
        },
        "blue": {
            get: function() { return [ "RGB", 0, 0, 1 ] }
        }, 
        "cyan": {
            get: function() { return [ "CMYK", 1,0,0,0 ] }
        }, 
        "magenta": {
            get: function() { return [ "CMYK", 0,1,0,0 ] }
        },      
        "yellow": {
            get: function() { return [ "CMYK", 0,0,1,0 ] }
        },       
        "dkGray": {
            get: function() { return [ "G", 0.25 ] }  // version 4. }
        },       
        "gray": {
            get: function() { return [ "G", 0.5 ] }   // version 4. }
        }, 
        "ltGray": {
            get: function() { return [ "G", 0.75 ] }   // version 4. }
        }
    });

    /**
	 * Converts the color space and color values specified by the color object to the specified color space:
     * Conversion to the gray color space is lossy (in the same fashion that displaying a color TV signal on a black-and-white TV is lossy).
     * The conversion of RGB to CMYK does not take into account any black generation or undercolor removal parameters.
	 * @memberof color
     * @param {Array} colorArray - Array of color values.
     * @param {string} cColorspace - The color space to which to convert.
	 * @typeofeditors ["PDF"]
	 * @returns {Array}
	 */
    color.convert = function(colorArray, cColorspace) {
        let nArrLenght = Array.isArray(colorArray) ? colorArray.length : -1;
            if (nArrLenght == -1)
                return ["T"];

        let fromSpace   = colorArray[0];
        let aComponents = colorArray.slice(1);

        // Convert from 'T' (Tint) to 'RGB'
        if (fromSpace === 'T') {
            return ["T"];
        }

        // Convert from 'RGB' to 'T' (Tint)
        if (fromSpace === 'RGB' && cColorspace === 'T') {
            return ["T"];
        }
        if (fromSpace === 'RGB' && cColorspace === 'RGB') {
            let r = typeof(aComponents[0]) == "number" && aComponents[0] > 0 ? aComponents[0] : 0;
            let g = typeof(aComponents[1]) == "number" && aComponents[1] > 0 ? aComponents[1] : 0;
            let b = typeof(aComponents[2]) == "number" && aComponents[2] > 0 ? aComponents[2] : 0;

            return ["RGB", r, g, b];
        }

        // Convert from 'RGB' to 'G' (Grayscale)
        if (fromSpace === 'RGB' && cColorspace === 'G') {
            let r = typeof(aComponents[0]) == "number" && aComponents[0] > 0 ? aComponents[0] : 0;
            let g = typeof(aComponents[1]) == "number" && aComponents[1] > 0 ? aComponents[1] : 0;
            let b = typeof(aComponents[2]) == "number" && aComponents[2] > 0 ? aComponents[2] : 0;

            return ["G", Math.round((r + g + b) / 3)];
        }

        // Convert from 'RGB' to 'CMYK'
        if (fromSpace === 'RGB' && cColorspace === 'CMYK') {
            let r = typeof(aComponents[0]) == "number" && aComponents[0] > 0 ? aComponents[0] : 0;
            let g = typeof(aComponents[1]) == "number" && aComponents[1] > 0 ? aComponents[1] : 0;
            let b = typeof(aComponents[2]) == "number" && aComponents[2] > 0 ? aComponents[2] : 0;

            let k = 1 - Math.max(r, g, b);
            let c = (1 - r - k) / (1 - k);
            let m = (1 - g - k) / (1 - k);
            let y = (1 - b - k) / (1 - k);
            return ["CMYK", c, m, y, k ];
        }

        // Convert from 'G' (Grayscale) to 'T' (Tint)
        if (fromSpace === 'G' && cColorspace === 'T') {
            return ["T"];
        }

        if (fromSpace === 'G' && cColorspace === 'G') {
            let color = typeof(aComponents[0]) == "number" && aComponents[0] > 0 ? aComponents[0] : 0;
            return ["G", color];
        }

        // Convert from 'G' (Grayscale) to 'RGB'
        if (fromSpace === 'G' && cColorspace === 'RGB') {
            let color = typeof(aComponents[0]) == "number" && aComponents[0] > 0 ? aComponents[0] : 0;
            return ["RGB", color, color, color];
        }

        // Convert from 'G' (Grayscale) to 'CMYK'
        if (fromSpace === 'G' && cColorspace === 'CMYK') {
            let color = typeof(aComponents[0]) == "number" && aComponents[0] > 0 ? aComponents[0] : 0;
            let k = Math.round((color) / 100);

            return ["CMYK", 0, 0, 0, k ];
        }

        // Convert from 'CMYK' to 'T' (Tint)
        if (fromSpace === 'CMYK' && cColorspace === 'T') {
            return ["T"];
        }

        if (fromSpace === 'CMYK' && cColorspace === 'CMYK') {
            let c = typeof(aComponents[0]) == "number" && aComponents[0] > 0 ? aComponents[0] : 0;
            let m = typeof(aComponents[1]) == "number" && aComponents[1] > 0 ? aComponents[1] : 0;
            let y = typeof(aComponents[2]) == "number" && aComponents[2] > 0 ? aComponents[2] : 0;
            let k = typeof(aComponents[3]) == "number" && aComponents[3] > 0 ? aComponents[3] : 0;

            return ["CMYK", c, m, y, k];
        }

        // Convert from 'CMYK' to 'RGB'
        if (fromSpace === 'CMYK' && cColorspace === 'RGB') {
            let c = typeof(aComponents[0]) == "number" && aComponents[0] > 0 ? aComponents[0] : 0;
            let m = typeof(aComponents[1]) == "number" && aComponents[1] > 0 ? aComponents[1] : 0;
            let y = typeof(aComponents[2]) == "number" && aComponents[2] > 0 ? aComponents[2] : 0;
            let k = typeof(aComponents[3]) == "number" && aComponents[3] > 0 ? aComponents[3] : 0;
            const r = Math.round((1 - c) * (1 - k));
            const g = Math.round((1 - m) * (1 - k));
            const b = Math.round((1 - y) * (1 - k));

            return ["RGB", r, g, b];
        }

        // Convert from 'CMYK' to 'G' (Grayscale)
        if (fromSpace === 'CMYK' && cColorspace === 'G') {
            let k = typeof(aComponents[3]) == "number" && aComponents[3] > 0 ? aComponents[3] : 0;
            let gray = Math.round(1 - k / 100);
            return ["G", gray];
        }

        return ["T"];
    };

    color["convert"] = color.convert;
    window["AscPDF"].Api.Objects = {
        border:     border,
        position:   position,
        scaleHow:   scaleHow,
        scaleWhen:  scaleWhen,
        display:    display,
        style:      style,
        highlight:  highlight,
        color:      color
    };
})();

