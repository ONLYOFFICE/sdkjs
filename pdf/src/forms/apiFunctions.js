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

(function(){
    /**
	 * Convert field value to specific number format.
     * @param {number} nDec = number of decimals
     * @param {number} sepStyle = separator style 0 = 1,234.56 / 1 = 1234.56 / 2 = 1.234,56 / 3 = 1234,56 / 4 = 1'234.56
     * @param {number} negStyle = 0 black minus / 1 red minus / 2 parens black / 3 parens red /
     * @param {number} currStyle = reserved
     * @param {string} strCurrency = string of currency to display
     * @param {boolean} bCurrencyPrepend = true = pre pend / false = post pend
	 * @typeofeditors ["PDF"]
	 */
    function AFNumber_Format(nDec, sepStyle, negStyle, currStyle, strCurrency, bCurrencyPrepend) {
        let oDoc = editor.getDocumentRenderer().doc;
        let oCurForm = oDoc.event["target"].field;

        let oInfoObj = {
            decimalPlaces: nDec,
            separator: true,
            symbol: null,
            type: Asc.c_oAscNumFormatType.Number
        }

        let oCultureInfo = {};
        Object.assign(oCultureInfo, AscCommon.g_aCultureInfos[oInfoObj.symbol]);
        switch (sepStyle) {
            case 0:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = ",";
                break;
            case 1:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = "";
                break;
            case 2:
                oCultureInfo.NumberDecimalSeparator = ",";
                oCultureInfo.NumberGroupSeparator = ".";
                break;
            case 3:
                oCultureInfo.NumberDecimalSeparator = ",";
                oCultureInfo.NumberGroupSeparator = "";
                break;
            case 4:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = "'";
                break;
        }

        oCultureInfo.NumberGroupSizes = [3];
        
        let aFormats = AscCommon.getFormatCells(oInfoObj);
        let oNumFormat = AscCommon.oNumFormatCache.get(aFormats[0]);
        let oTargetRun = oCurForm.contentFormat.GetElement(0).GetElement(0);

        let sCurValue;
        if (oCurForm.GetType() == AscPDF.FIELD_TYPES.text)
            sCurValue = oCurForm.GetValue();
        else if (oCurForm.GetType() == AscPDF.FIELD_TYPES.combobox) {
            let nCurIdx = oCurForm.GetCurIdxs();
            sCurValue = nCurIdx == -1 ? oCurForm.GetValue() : oCurForm.api.getItemAt(nCurIdx, false);
        }

        sCurValue = sCurValue.replace(",", ".");
        if (sCurValue == "") {
            oTargetRun.ClearContent();
            return;
        }
            
        let sRes = oNumFormat.format(sCurValue, 0, AscCommon.gc_nMaxDigCount, true, oCultureInfo, true)[0].text;

        if (!sRes) {
            oTargetRun.ClearContent();
            return;
        }

        if (bCurrencyPrepend)
            sRes = strCurrency + sRes;
        else
            sRes = sRes + strCurrency;

        if (sRes.indexOf("-") != - 1) {
            switch (negStyle) {
                case 0:
                    oTargetRun.Pr.Color = private_GetColor(255, 255, 255, true);
                    break;
                case 1:
                    oTargetRun.Pr.Color = private_GetColor(255, 0, 0, false);
                    break;
                case 2:
                    sRes = sRes.replace("-", "");
                    oTargetRun.Pr.Color = private_GetColor(255, 255, 255, true);
                    sRes = "(" + sRes + ")";
                    break;
                case 3:
                    sRes = sRes.replace("-", "");
                    oTargetRun.Pr.Color = private_GetColor(255, 0, 0, false);
                    sRes = "(" + sRes + ")";
                    break;
            }
        }
        else {
            oTargetRun.Pr.Color = private_GetColor(255, 255, 255, true);
        }
        
        oTargetRun.RecalcInfo.TextPr = true
        oTargetRun.ClearContent();
        oTargetRun.AddText(sRes);
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {number} nDec = number of decimals
     * @param {number} sepStyle = separator style 0 = 1,234.56 / 1 = 1234.56 / 2 = 1.234,56 / 3 = 1234,56 / 4 = 1'234.56
     * @param {number} negStyle = 0 black minus / 1 red minus / 2 parens black / 3 parens red /
     * @param {number} currStyle = reserved
     * @param {string} strCurrency = string of currency to display
     * @param {boolean} bCurrencyPrepend = true = pre pend / false = post pend
	 * @typeofeditors ["PDF"]
	 */
    function AFNumber_Keystroke(nDec, sepStyle, negStyle, currStyle, strCurrency, bCurrencyPrepend) {
        let oDoc    = editor.getDocumentRenderer().doc;
        let oForm   = oDoc.event["target"].field;
        let sValue  = oDoc.event["value"];
        let sChange = oDoc.event["change"];
        let nSelStart = oDoc.event["selStart"];
        let nSelEnd = oDoc.event["selEnd"];

        function isValidNumber(str) {
            return !isNaN(str) && isFinite(str);
        }

        let sNewValue = sValue.slice(0, nSelStart) + sChange + sValue.slice(nSelEnd);

        // разделитель дробной части, который можно ввести
        switch (sepStyle) {
            case 0:
            case 1:
            case 4:
                if (sNewValue.indexOf(",") != -1) {
                    oDoc.event["rc"] = false;
                }

                if (isValidNumber(sNewValue) == false) {
                    oDoc.event["rc"] = false;
                }
                break;
            case 2:
            case 3:
                if (sNewValue.indexOf(".") != -1) {
                    oDoc.event["rc"] = false;
                }
                
                sNewValue = sNewValue.replace(/\,/g, ".");
                if (isValidNumber(sNewValue) == false) {
                    oDoc.event["rc"] = false;
                }
                break;
        }

        if (oDoc.event["rc"] == false && oDoc.event["willCommit"]) {
            oDoc.SetWarningInfo({
                "target": oForm
            });
        }
    }

    /**
	 * Convert field value to specific percent format.
	 * @memberof CTextField
     * @param {number} nDec = number of decimals
     * @param {number} sepStyle = separator style 0 = 1,234.56 / 1 = 1234.56 / 2 = 1.234,56 / 3 = 1234,56 / 4 = 1'234.56
	 * @typeofeditors ["PDF"]
	 */
    function AFPercent_Format(nDec, sepStyle) {
        let oDoc = editor.getDocumentRenderer().doc;
        let oCurForm = oDoc.event["target"].field;
        
        let oInfoObj = {
            decimalPlaces: nDec,
            separator: true,
            symbol: null,
            type: Asc.c_oAscNumFormatType.Number
        }

        let oCultureInfo = {};
        Object.assign(oCultureInfo, AscCommon.g_aCultureInfos[oInfoObj.symbol]);
        switch (sepStyle) {
            case 0:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = ",";
                break;
            case 1:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = "";
                break;
            case 2:
                oCultureInfo.NumberDecimalSeparator = ",";
                oCultureInfo.NumberGroupSeparator = ".";
                break;
            case 3:
                oCultureInfo.NumberDecimalSeparator = ",";
                oCultureInfo.NumberGroupSeparator = "";
                break;
            case 4:
                oCultureInfo.NumberDecimalSeparator = ".";
                oCultureInfo.NumberGroupSeparator = "'";
                break;
        }
        oCultureInfo.NumberGroupSizes = [3];

        let aFormats = AscCommon.getFormatCells(oInfoObj);
        let oNumFormat = AscCommon.oNumFormatCache.get(aFormats[0]);
        let oTargetRun = oCurForm.contentFormat.GetElement(0).GetElement(0);

        let sCurValue;
        if (oCurForm.GetType() == AscPDF.FIELD_TYPES.text)
            sCurValue = oCurForm.GetValue();
        else if (oCurForm.GetType() == AscPDF.FIELD_TYPES.combobox) {
            let nCurIdx = oCurForm.GetCurIdxs();
            sCurValue = nCurIdx == -1 ? oCurForm.GetValue() : oCurForm.api.getItemAt(nCurIdx, false);
        }

        sCurValue = sCurValue.replace(",", ".");
        if (sCurValue == "")
            sCurValue = 0;
            
        sCurValue = (parseFloat(sCurValue) * 100).toString();
        let sRes = oNumFormat.format(sCurValue, 0, AscCommon.gc_nMaxDigCount, true, oCultureInfo, true)[0].text;
        if (!sRes) {
            oTargetRun.ClearContent();
            return;
        }

        sRes = sRes + "%";

        oTargetRun.ClearContent();
        oTargetRun.AddText(sRes);
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {number} nDec = number of decimals
     * @param {number} sepStyle = separator style 0 = 1,234.56 / 1 = 1234.56 / 2 = 1.234,56 / 3 = 1234,56 / 4 = 1'234.56
	 * @typeofeditors ["PDF"]
	 */
    function AFPercent_Keystroke(nDec, sepStyle) {
        let oDoc    = editor.getDocumentRenderer().doc;
        let oForm   = oDoc.event["target"].field;
        let sValue  = oDoc.event["value"];
        let sChange = oDoc.event["change"];
        let nSelStart = oDoc.event["selStart"];
        let nSelEnd = oDoc.event["selEnd"];

        let sNewValue = sValue.slice(0, nSelStart) + sChange + sValue.slice(nSelEnd);

        function isValidNumber(str) {
            return !isNaN(str) && isFinite(str);
        }

        // разделитель дробной части, который можно ввести
        switch (sepStyle) {
            case 0:
            case 1:
            case 4:
                if (sNewValue.indexOf(",") != -1) {
                    oDoc.event["rc"] = false;
                }

                if (isValidNumber(sNewValue) == false) {
                    oDoc.event["rc"] = false;
                }
                break;
            case 2:
            case 3:
                if (sNewValue.indexOf(".") != -1) {
                    oDoc.event["rc"] = false;
                }

                sNewValue = sNewValue.replace(/\,/g, ".");
                if (isValidNumber(sNewValue) == false) {
                    oDoc.event["rc"] = false;
                }

                break;
        }

        if (oDoc.event["rc"] == false) {
            oDoc.SetWarningInfo({
                "target": oForm
            });
        }
    }

    /**
	 * Convert field value to specific date format.
	 * @memberof CTextField
     * @param {string} cFormat - date format
	 * @typeofeditors ["PDF"]
	 */
    function AFDate_Format(cFormat) {
        let oDoc = editor.getDocumentRenderer().doc;
        let oCurForm = oDoc.event["target"].field;
        oDoc.event["format"] = cFormat;
        
        let oNumFormat = AscCommon.oNumFormatCache.get(cFormat, AscCommon.NumFormatType.PDFFormDate);
        oNumFormat.oNegativeFormat.bAddMinusIfNes = false;
        
        let oTargetRun = oCurForm.contentFormat.GetElement(0).GetElement(0);

        let sCurValue;
        if (oCurForm.GetType() == AscPDF.FIELD_TYPES.text)
            sCurValue = oCurForm.GetValue();
        else if (oCurForm.GetType() == AscPDF.FIELD_TYPES.combobox) {
            let nCurIdx = oCurForm.GetCurIdxs();
            sCurValue = nCurIdx == -1 ? oCurForm.GetValue() : oCurForm.api.getItemAt(nCurIdx, false);
        }

        let oFormatParser = AscCommon.g_oFormatParser;

        function getShortPattern(aRawFormat) {
            let dayDone     = false;
            let monthDone   = false;
            let yearDone    = false;
            
            let sPattern = "";

            let numFormat_Year = 12;
            let numFormat_Month = 13;
            let numFormat_Day = 16;

            for (let i = 0; i < aRawFormat.length; i++) {
                let obj = aRawFormat[i];
                switch (obj.type) {
                    case numFormat_Day:
                        if (dayDone == false) {
                            sPattern += 1;
                            dayDone = true;
                        }
                        break;
                    case numFormat_Month:
                        if (monthDone == false) {
                            sPattern += 3;
                            monthDone = true;
                        }
                        break;
                    case numFormat_Year:
                        if (yearDone == false) {
                            if (obj.val > 2)
                                sPattern += 5;
                            else
                                sPattern += 4;
                            yearDone = true;
                        }
                        break;
                            
                }
            }
            return sPattern;
        }

        let oCultureInfo = {};
        Object.assign(oCultureInfo, AscCommon.g_aCultureInfos[9]);
        if (null == oNumFormat.oTextFormat.ShortDatePattern) {
            oNumFormat.oTextFormat.ShortDatePattern = getShortPattern(oNumFormat.oTextFormat.aRawFormat);
            oNumFormat.oTextFormat._prepareFormatDatePDF();
        }
        oCultureInfo.ShortDatePattern = oNumFormat.oTextFormat.ShortDatePattern;

        if (oCultureInfo.ShortDatePattern.indexOf("1") == -1)
            oNumFormat.oTextFormat.bDay = false;

        oCultureInfo.AbbreviatedMonthNames.length = 12;
        oCultureInfo.MonthNames.length = 12;

        let oResParsed = oFormatParser.parseDatePDF(sCurValue, oCultureInfo, oNumFormat);
                
        if (sCurValue == "")
            oTargetRun.ClearContent();

        if (!oResParsed) {
            oTargetRun.ClearContent();
            return;
        }

        oNumFormat.oTextFormat.formatType = AscCommon.NumFormatType.PDFFormDate;
        let sRes = oNumFormat.oTextFormat.format(oResParsed.value, 0, AscCommon.gc_nMaxDigCount, oCultureInfo)[0].text;

        oTargetRun.ClearContent();
        oTargetRun.AddText(sRes);
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {string} cFormat - date format
	 * @typeofeditors ["PDF"]
	 */
    function AFDate_Keystroke(cFormat) {
        let oDoc        = editor.getDocumentRenderer().doc;
        let oForm       = oDoc.event["target"].field;
        
        if (oDoc.event["willCommit"] == false) {
            oDoc.event["rc"] = true;
            return;
        }

        let oNumFormat = AscCommon.oNumFormatCache.get(cFormat, AscCommon.NumFormatType.PDFFormDate);
        oNumFormat.oNegativeFormat.bAddMinusIfNes = false;
        
        let sCurValue;
        if (oForm.GetType() == AscPDF.FIELD_TYPES.text)
            sCurValue = oForm.GetValue();
        else if (oForm.GetType() == AscPDF.FIELD_TYPES.combobox) {
            let nCurIdx = oForm.GetCurIdxs();
            sCurValue = nCurIdx == -1 ? oForm.GetValue() : oForm.GetFormApi().getItemAt(nCurIdx, false);
        }

        let oFormatParser = AscCommon.g_oFormatParser;

        function getShortPattern(aRawFormat) {
            let dayDone     = false;
            let monthDone   = false;
            let yearDone    = false;
            
            let sPattern = "";

            let numFormat_Year = 12;
            let numFormat_Month = 13;
            let numFormat_Day = 16;

            for (let i = 0; i < aRawFormat.length; i++) {
                let obj = aRawFormat[i];
                switch (obj.type) {
                    case numFormat_Day:
                        if (dayDone == false) {
                            sPattern += 1;
                            dayDone = true;
                        }
                        break;
                    case numFormat_Month:
                        if (monthDone == false) {
                            sPattern += 3;
                            monthDone = true;
                        }
                        break;
                    case numFormat_Year:
                        if (yearDone == false) {
                            if (obj.val > 2)
                                sPattern += 5;
                            else
                                sPattern += 4;
                            yearDone = true;
                        }
                        break;
                            
                }
            }
            return sPattern;
        }

        let oCultureInfo = {};
        Object.assign(oCultureInfo, AscCommon.g_aCultureInfos[9]);
        if (null == oNumFormat.oTextFormat.ShortDatePattern) {
            oNumFormat.oTextFormat.ShortDatePattern = getShortPattern(oNumFormat.oTextFormat.aRawFormat);
            oNumFormat.oTextFormat._prepareFormatDatePDF();
        }
        oCultureInfo.ShortDatePattern = oNumFormat.oTextFormat.ShortDatePattern;

        if (oCultureInfo.ShortDatePattern.indexOf("1") == -1)
            oNumFormat.oTextFormat.bDay = false;

        oCultureInfo.AbbreviatedMonthNames.length = 12;
        oCultureInfo.MonthNames.length = 12;

        let oResParsed = oFormatParser.parseDatePDF(sCurValue, oCultureInfo, oNumFormat);
                
        if (!oResParsed) {
            oDoc.event["rc"] = false;
            oDoc.SetWarningInfo({
                "format": cFormat,
                "target": oForm
            });
            return;
        }
    }
    let AFDate_FormatEx = AFDate_Format;
    let AFDate_KeystrokeEx = AFDate_Keystroke;

    /**
	 * Convert field value to specific time format.
	 * @memberof CTextField
     * @param {number} ptf - time format
     *  0 = 24HR_MM [ 14:30 ]
     *  1 = 12HR_MM [ 2:30 PM ]
     *  2 = 24HR_MM_SS [ 14:30:15 ]
     *  3 = 12HR_MM_SS [ 2:30:15 PM ]
	 * @typeofeditors ["PDF"]
	 */
    function AFTime_Format(ptf) {
        let oDoc = editor.getDocumentRenderer().doc;
        let oForm = oDoc.event["target"].field;
        
        // to do сделать обработку ms 
        let oCultureInfo = {};
        Object.assign(oCultureInfo, AscCommon.g_aCultureInfos[9]);

        let sFormat;
        let fIsValidTime = null;
        switch (ptf) {
            case 0:
                sFormat = "HH:MM";
                fIsValidTime = function isValidTime(time) {
                    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]|[0-9])\s*$/;
                    return pattern.test(time);
                }
                break;
            case 1:
                sFormat = "h:MM tt";
                fIsValidTime = function isValidTime(time) {
                    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]|[0-9])\s*([APap][mM])?$/;
                    return pattern.test(time);
                }
                break;
            case 2:
                sFormat = "HH:MM:ss";
                fIsValidTime = function isValidTime(time) {
                    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]|[0-9])(:([0-5][0-9]|[0-9]))?\s*$/;
                    return pattern.test(time);
                }
                break;
            case 3:
                sFormat = "h:MM:ss tt";
                fIsValidTime = function isValidTime(time) {
                    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]|[0-9])(:([0-5][0-9]|[0-9]))?\s*([APap][mM])?$/;
                    return pattern.test(time);
                }
                break;
        }

        let oNumFormat = AscCommon.oNumFormatCache.get(sFormat, AscCommon.NumFormatType.PDFFormDate);
        oNumFormat.oNegativeFormat.bAddMinusIfNes = false;
        
        let oTargetRun = oForm.contentFormat.GetElement(0).GetElement(0);

        let sCurValue;
        if (oForm.GetType() == AscPDF.FIELD_TYPES.text)
            sCurValue = oForm.GetValue();
        else if (oForm.GetType() == AscPDF.FIELD_TYPES.combobox) {
            let nCurIdx = oForm.GetCurIdxs();
            sCurValue = nCurIdx == -1 ? oForm.GetValue() : oForm.api.getItemAt(nCurIdx, false);
        }
        
        if (sCurValue == "")
            oTargetRun.ClearContent();
        else if (fIsValidTime(sCurValue) == false) {
            oTargetRun.ClearContent();
            return;
        }
        
        let oFormatParser = AscCommon.g_oFormatParser;
        let oResParsed = oFormatParser.parseDatePDF(sCurValue, AscCommon.g_aCultureInfos[9]);

        if (!oResParsed) {
            oTargetRun.ClearContent();
            return;
        }
        
        oNumFormat.oTextFormat.formatType = AscCommon.NumFormatType.PDFFormDate;
        let sRes = oNumFormat.format(oResParsed.value, 0, AscCommon.gc_nMaxDigCount, true, undefined, true)[0].text;

        oTargetRun.ClearContent();
        oTargetRun.AddText(sRes);
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {string} ptf is the time format:
     *  0 = 24HR_MM [ 14:30 ]
     *  1 = 12HR_MM [ 2:30 PM ]
     *  2 = 24HR_MM_SS [ 14:30:15 ]
     *  3 = 12HR_MM_SS [ 2:30:15 PM ]
	 * @typeofeditors ["PDF"]
	 */
    function AFTime_Keystroke(ptf) {
        let oDoc = editor.getDocumentRenderer().doc;
        let oForm = oDoc.event["target"].field;
        
        if (oDoc.event["willCommit"] == false) {
            oDoc.event["rc"] = true;
            return;
        }

        // to do сделать обработку ms
        let oCultureInfo = {};
        Object.assign(oCultureInfo, AscCommon.g_aCultureInfos[9]);

        let sFormat;
        let fIsValidTime = null;
        switch (ptf) {
            case 0:
                sFormat = "HH:MM";
                fIsValidTime = function isValidTime(time) {
                    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]|[0-9])\s*$/;
                    return pattern.test(time);
                }
                break;
            case 1:
                sFormat = "h:MM tt";
                fIsValidTime = function isValidTime(time) {
                    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]|[0-9])\s*([APap][mM])?$/;
                    return pattern.test(time);
                }
                break;
            case 2:
                sFormat = "HH:MM:ss";
                fIsValidTime = function isValidTime(time) {
                    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]|[0-9])(:([0-5][0-9]|[0-9]))?\s*$/;
                    return pattern.test(time);
                }
                break;
            case 3:
                sFormat = "h:MM:ss tt";
                fIsValidTime = function isValidTime(time) {
                    const pattern = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):([0-5][0-9]|[0-9])(:([0-5][0-9]|[0-9]))?\s*([APap][mM])?$/;
                    return pattern.test(time);
                }
                break;
            default:
                sFormat = ptf;
                break;
        }

        let oNumFormat = AscCommon.oNumFormatCache.get(sFormat, AscCommon.NumFormatType.PDFFormDate);
        oNumFormat.oNegativeFormat.bAddMinusIfNes = false;
        
        let sCurValue;
        if (oForm.GetType() == AscPDF.FIELD_TYPES.text)
            sCurValue = oForm.GetValue();
        else if (oForm.GetType() == AscPDF.FIELD_TYPES.combobox) {
            let nCurIdx = oForm.GetCurIdxs();
            sCurValue = nCurIdx == -1 ? oForm.GetValue() : oForm.api.getItemAt(nCurIdx, false);
        }
        
       if (fIsValidTime && fIsValidTime(sCurValue) == false) {
            oDoc.event["rc"] = false;
            oDoc.SetWarningInfo({
                "format": ptf,
                "target": oForm
            });
            return;
        }
        
        let oFormatParser = AscCommon.g_oFormatParser;
        let oResParsed = oFormatParser.parseDatePDF(sCurValue, AscCommon.g_aCultureInfos[9]);

        if (!oResParsed) {
            oDoc.event["rc"] = false;
            oDoc.SetWarningInfo({
                "format": ptf,
                "target": oForm
            });
            return;
        }
    }

    let AFTime_FormatEx = AFDate_FormatEx;
    let AFTime_KeystrokeEx = AFTime_Keystroke;

    /**
	 * Convert field value to specific special format.
	 * @memberof CTextField
     * @param {number} psf – psf is the type of formatting to use:
     *  0 = zip code
     *  1 = zip + 4
     *  2 = phone
     *  3 = SSN
	 * @typeofeditors ["PDF"]
	 */
    function AFSpecial_Format(psf) {
        let oDoc = editor.getDocumentRenderer().doc;
        let oCurForm = oDoc.event["target"].field;
        
        let sCurValue;
        if (oCurForm.GetType() == AscPDF.FIELD_TYPES.text)
            sCurValue = oCurForm.GetValue();
        else if (oCurForm.GetType() == AscPDF.FIELD_TYPES.combobox) {
            let nCurIdx = oCurForm.GetCurIdxs();
            sCurValue = nCurIdx == -1 ? oCurForm.GetValue() : oCurForm.api.getItemAt(nCurIdx, false);
        }

        let oTargetRun = oCurForm.contentFormat.GetElement(0).GetElement(0);

        function isValidZipCode(zipCode) {
            let regex = /^\d{5}$/;
            return regex.test(zipCode);
        }
        function isValidZipCode4(zip) {
            let regex = /^\d{5}[-\s.]?(\d{4})?$/;
            return regex.test(zip);
        }
        function isValidPhoneNumber(number) {
            let regex = /^\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}$/;
            return regex.test(number);
        }
        function isValidSSN(ssn) {
            let regex = /^\d{3}[-\s.]?\d{2}[-\s.]?\d{4}$/;
            return regex.test(ssn);
        }

        if (sCurValue == "")
            oTargetRun.ClearContent();
            
        switch (psf) {
            case 0:
                if (isValidZipCode(sCurValue) == false) {
                    oTargetRun.ClearContent();
                    return;
                }
                break;
            case 1:
                if (isValidZipCode4(sCurValue) == false) {
                    oTargetRun.ClearContent();
                    return;
                }
                break;
            case 2:
                if (isValidPhoneNumber(sCurValue) == false) {
                    oTargetRun.ClearContent();
                    return;
                }
                break;
            case 3:
                if (isValidSSN(sCurValue) == false) {
                    oTargetRun.ClearContent();
                    return;
                }
                break;
        }

        sCurValue = sCurValue.replace(/\D/g, ""); // delete all except no digit chars
        let sFormatValue = "";

        oTargetRun.ClearContent();

        switch (psf) {
            case 0:
                sFormatValue = sCurValue.substring(0, 5);
                break;
            case 1:
                sFormatValue = sCurValue.substring(0, 9);
                if (sFormatValue[4])
                    sFormatValue = sCurValue.substring(0, 5) + "-" + sCurValue.substring(5);
                break;
            case 2: 
                let x = sCurValue.substring(0, 10);
                sFormatValue = x.length > 6
                    ? "(" + x.substring(0, 3) + ") " + x.substring(3, 6) + "-" + x.substring(6, 10)
                    : x.length > 3
                    ? "(" + x.substring(0, 3) + ") " + x.substring(3)
                    : x;
                break;
            case 3:
                let y = sCurValue.substring(0, 9);
                sFormatValue = y.length > 5
                ? y.substring(0, 3) + "-" + y.substring(3, 5) + "-" + y.substring(5, 9)
                : y.length > 2
                ? y.substring(0, 3) + "-" + x.substring(3)
                : x;
                break;
        }

        oTargetRun.AddText(sFormatValue);
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {number} psf – psf is the type of formatting to use:
     *  0 = zip code
     *  1 = zip + 4
     *  2 = phone
     *  3 = SSN
	 * @typeofeditors ["PDF"]
	 */
    function AFSpecial_Keystroke(psf) {
        let oDoc    = editor.getDocumentRenderer().doc;
        let sValue  = oDoc.event["value"];
        let sChange = oDoc.event["change"];
        let nSelStart = oDoc.event["selStart"];
        let nSelEnd = oDoc.event["selEnd"];

        let isValidZipCode, isValidZipCode4, isValidPhoneNumber, isValidSSN;

        if (oDoc.event["willCommit"]) {
            isValidZipCode = function(zipCode) {
                let regex = /^\d{5}$/;
                return regex.test(zipCode);
            };
            isValidZipCode4 = function(zip) {
                let regex = /^\d{5}[-\s.]?(\d{4})?$/;
                return regex.test(zip);
            };
            isValidPhoneNumber = function(number) {
                let regex = /^\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}$/;
                return regex.test(number);
            };
            isValidSSN = function(ssn) {
                let regex = /^\d{3}[-\s.]?\d{2}[-\s.]?\d{4}$/;
                return regex.test(ssn);
            };
        }
        else {
            isValidZipCode = function(zipCode) {
                let regex = /^\d{0,5}$/;
                return regex.test(zipCode);
            };
            isValidZipCode4 = function(zip) {
                let regex = /^\d{0,5}[-\s.]?(\d{0,4})?$/;
                return regex.test(zip);
            };
            isValidPhoneNumber = function(number) {
                let regex = /^\(?\d{0,3}?\)?[\s.-]?\d{0,3}?[\s.-]?\d{0,4}?$/;
                return regex.test(number);
            };
            isValidSSN = function(ssn) {
                let regex = /^\d{0,3}?[-\s.]?\d{0,2}?[-\s.]?\d{0,4}$/;
                return regex.test(ssn);
            };
        }
        
        let sNewValue = sValue.slice(0, nSelStart) + sChange + sValue.slice(nSelEnd);

        let canAdd;
        switch (psf) {
            case 0:
                canAdd = isValidZipCode(sNewValue);
                break;
            case 1:
                canAdd = isValidZipCode4(sNewValue);
                break;
            case 2:
                canAdd = isValidPhoneNumber(sNewValue);
                break;
            case 3:
                canAdd = isValidSSN(sNewValue);
                break;
        }

        oDoc.event["rc"] = canAdd;
        if (canAdd == false && oDoc.event["willCommit"]) {
            oDoc.SetWarningInfo({
                "target": oDoc.event["target"].field
            });
        }
    }
    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {number} mask - the special mask
	 * @typeofeditors ["PDF"]
	 */
    function AFSpecial_KeystrokeEx(mask) {
        let oDoc    = editor.getDocumentRenderer().doc;
        let oForm   = oDoc.event["target"].field;
        let sValue  = oDoc.event["value"];
        let sChange = oDoc.event["change"];
        let nSelStart = oDoc.event["selStart"];
        let nSelEnd = oDoc.event["selEnd"];

        if (typeof(mask) != "string" || mask == "") {
            oDoc.event["rc"] = false;
            oDoc.SetWarningInfo({
                "format": mask,
                "target": oForm
            });
            return;
        }
            
        let oFormMask   = new AscWord.CTextFormMask();
        let oTextFormat = new AscWord.CTextFormFormat();
        oFormMask.Set(mask);

        // если текущее значение не подходит по маске, то ввод запрещаем
        let isCanEnter;
        if (oDoc.event["willCommit"])
            isCanEnter = sValue != "" ? oFormMask.Check(oTextFormat.GetBuffer(sValue), true) : true;
        else
            isCanEnter = oFormMask.Check(oTextFormat.GetBuffer(sValue));

        if (isCanEnter == false) {
            oDoc.event["rc"] = false;
            if (oDoc.event["willCommit"]) {
                oDoc.SetWarningInfo({
                    "format": mask,
                    "target": oForm
                });

                return;
            }
        }

        let aChangeChars = oTextFormat.GetBuffer(sChange);
        let inOnlyRemove = aChangeChars.length == 0 && nSelEnd != nSelStart;

        let sNewValue = sValue.slice(0, nSelStart) + sChange + sValue.slice(nSelEnd);
        let arrBuffer = oTextFormat.GetBuffer(sNewValue);
        
        // проверяем подходит ли по маске, если нет пытаемся скорректировать и проверяем снова
        isCanEnter = oFormMask.Check(arrBuffer);
        if (!isCanEnter) {
            if (inOnlyRemove == true) {
                oDoc.event["rc"] = false;
                return;
            }
                
            let sCorrected = oFormMask.Correct(sNewValue);
            arrBuffer = oTextFormat.GetBuffer(sCorrected);
            isCanEnter = oFormMask.Check(arrBuffer);

            if (isCanEnter) {
                // находим скорректированные символы для вставки после коррекции
                let nCount = sCorrected.length - sValue.length; // кол-во вставленных скорректированных символов
                let sFinalChanges = sCorrected.slice(nSelStart, nSelStart + nCount);
                oDoc.event["rc"] = true;
                oDoc.event["change"] = sFinalChanges;
            }
            else {
                oDoc.event["rc"] = false;
                oDoc.SetWarningInfo({
                    "format": mask,
                    "target": oForm
                });
                return;
            }
        }
        else {
            oDoc.event["rc"] = true;
            oDoc.SetWarningInfo({
                "format": mask,
                "target": oForm
            });
            return;
        }            
    }

    /**
	 * Check can the field accept the char or not.
	 * @memberof CTextField
     * @param {string} sFunction -  is one of "AVG", "SUM", "PRD", "MIN", "MAX"
     * @param {string[]} aFieldsNames - is the list of the fields to use in the calculation
	 * @typeofeditors ["PDF"]
	 */
    function AFSimple_Calculate(sFunction, aFieldsNames) {
        let oViewer = editor.getDocumentRenderer();
        let oDoc    = oViewer.doc;
        let oField  = oDoc.event["target"].field;

        let aFields = [];
        aFieldsNames.forEach(function(name) {
            let aTmpFields = oDoc.GetFields(name);

            // смысл в том, что одно поле из многих с таким именем может быть изменено, но еще не применено, ищем его
            // если нет, то просто берем первое из массива
            let oChangedField = aTmpFields.find(function(field) {
                return field.IsNeedCommit();
            });

            if (oChangedField)
                aFields.push(oChangedField);
            else if (aTmpFields[0])
                aFields.push(aTmpFields[0]);
        });

        function extractNumber(str) {
            let resultStr = str.replace(/^[^\d]*(\d+)/, "$1").replace(/[^0-9]+/, '.');
            return parseFloat(resultStr);
        }

        let aValues = [];
        aFields.forEach(function(field) {
            if (field.GetType() == AscPDF.FIELD_TYPES.button)
                return;

            let fieldValue = field.GetValue();
            
            // multi select listbox
            if (Array.isArray(fieldValue)) {
                aValues = aValues.concat(fieldValue);
            }
            else {
                aValues.push(fieldValue !== "Off" && fieldValue != "" ? fieldValue : "0");
            }
        });

        let nResult;
        switch (sFunction) {
            case "SUM":
                nResult = aValues.reduce(function(sum, current) {
                    let nParsedNumber = extractNumber(current);
                    let nFracSum = sum.toString().split('.')[1] ? sum.toString().split('.')[1].length : 0;
                    let nFracCurr = nParsedNumber.toString().split('.')[1] ? nParsedNumber.toString().split('.')[1].length : 0;
                    let nMaxFrac = Math.max(nFracSum, nFracCurr);
                    
                    return Math.round((sum + nParsedNumber) * (Math.pow(10, nMaxFrac))) / (Math.pow(10, nMaxFrac)); // исправляем беды с дробной частью
                }, 0);
                break;
            case "PRD":
                nResult = aValues.reduce(function(sum, current) {
                    let nParsedNumber = extractNumber(current);
                    let nFracSum = sum.toString().split('.')[1] ? sum.toString().split('.')[1].length : 0;
                    let nFracCurr = nParsedNumber.toString().split('.')[1] ? nParsedNumber.toString().split('.')[1].length : 0;
                    let nMaxFrac = Math.max(nFracSum, nFracCurr);
                    
                    return Math.round((sum * nParsedNumber) * (Math.pow(10, nMaxFrac))) / (Math.pow(10, nMaxFrac)); // исправляем беды с дробной частью
                }, 1);
                break;
            case "AVG":
                nResult = aValues.reduce(function(sum, current) {
                    let nParsedNumber = extractNumber(current);
                    let nFracSum = sum.toString().split('.')[1] ? sum.toString().split('.')[1].length : 0;
                    let nFracCurr = nParsedNumber.toString().split('.')[1] ? nParsedNumber.toString().split('.')[1].length : 0;
                    let nMaxFrac = Math.max(nFracSum, nFracCurr);

                    return Math.round((sum + nParsedNumber) * (Math.pow(10, nMaxFrac))) / (Math.pow(10, nMaxFrac)); // исправляем беды с дробной частью
                }, 0);
                nResult = nResult / aValues.length;
                break;
            case "MIN":
                let nMin = aValues[0];
                for (let i = 1; i < aValues.length; i++) {
                    if (aValues[i] < nMin)
                        nMin = aValues[i];
                }
                nResult = nMin;
                break;
            case "MAX":
                let nMax = aValues[0];
                for (let i = 1; i < aValues.length; i++) {
                    if (aValues[i] > nMax)
                        nMax = aValues[i];
                }

                nResult = nMax;
                break;
        }

        oField.DoValidateAction(String(nResult));
        
        if (oDoc.event["rc"]) {
            oField.SetValue(String(nResult));
            oField.SetNeedCommit(true);
        }
    }

    /**
	 * Function for validation value.
	 * @memberof CTextField
	 * @typeofeditors ["PDF"]
     * @returns {boolean}
	 */
    function AFRange_Validate(bGreaterThan, nGreaterThan, bLessThan, nLessThan) {
        let oDoc = editor.getDocumentRenderer().doc;

        if (bGreaterThan && bLessThan) {
            oDoc.event["rc"] = oDoc.event["value"] >= nGreaterThan && oDoc.event["value"] <= nLessThan;
            if (oDoc.event["rc"] == false) {
                oDoc.SetWarningInfo({
                    "greater": nGreaterThan,
                    "less": nLessThan,
                    "target": oDoc.event["target"].field
                });
            }
        }
        else if (bGreaterThan) {
            oDoc.event["rc"] = oDoc.event["value"] >= nGreaterThan;
            if (oDoc.event["rc"] == false) {
                oDoc.SetWarningInfo({
                    "greater": nGreaterThan,
                    "target": oDoc.event["target"].field
                });
            }
        }
        else if (bLessThan) {
            oDoc.event["rc"] = oDoc.event["value"] <= nLessThan;
            if (oDoc.event["rc"] == false) {
                oDoc.SetWarningInfo({
                    "less": nLessThan,
                    "target": oDoc.event["target"].field
                });
            }
        }
    }
	
    function private_GetColor(r, g, b, Auto)
	{
		return new AscCommonWord.CDocumentColor(r, g, b, Auto ? Auto : false);
	}
    
    window["AscPDF"].Api = {
        Functions: {
            AFNumber_Format:        AFNumber_Format,
            AFNumber_Keystroke:     AFNumber_Keystroke,
            AFPercent_Format:       AFPercent_Format,
            AFPercent_Keystroke:    AFPercent_Keystroke,
            AFDate_Format:          AFDate_Format,
            AFDate_Keystroke:       AFDate_Keystroke,
            AFDate_FormatEx:        AFDate_FormatEx,
            AFDate_KeystrokeEx:     AFDate_KeystrokeEx,
            AFTime_Format:          AFTime_Format,
            AFTime_Keystroke:       AFTime_Keystroke,
            AFTime_FormatEx:        AFTime_FormatEx,
            AFTime_KeystrokeEx:     AFTime_KeystrokeEx,
            AFSpecial_Format:       AFSpecial_Format,
            AFSpecial_Keystroke:    AFSpecial_Keystroke,
            AFSpecial_KeystrokeEx:  AFSpecial_KeystrokeEx,
            AFSimple_Calculate:     AFSimple_Calculate,
            AFRange_Validate:       AFRange_Validate,
        }
    }
})();

